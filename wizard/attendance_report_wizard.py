import base64
import io
from collections import defaultdict
from datetime import date, datetime, timedelta

from odoo import fields, models, _
from odoo.exceptions import UserError


def _default_date_from(self):
    today = date.today()
    first = today.replace(day=1)
    prev = first - timedelta(days=1)
    return prev.replace(day=1)


def _default_date_to(self):
    today = date.today()
    return today.replace(day=1) - timedelta(days=1)


class AttendanceReportWizard(models.TransientModel):
    _name = 'xattendance.report.wizard'
    _description = 'Attendance XLSX Report Wizard'

    date_from = fields.Date(string='Date From', required=True, default=_default_date_from)
    date_to = fields.Date(string='Date To', required=True, default=_default_date_to)
    employee_ids = fields.Many2many('hr.employee', string='Employees')
    department_ids = fields.Many2many('hr.department', string='Departments')
    file_data = fields.Binary(string='XLSX File', readonly=True)
    filename = fields.Char(readonly=True)

    def action_generate(self):
        self.ensure_one()
        try:
            import xlsxwriter
        except ImportError:
            raise UserError(_('Please install xlsxwriter: pip install xlsxwriter'))

        domain = [
            ('check_in', '>=', fields.Datetime.from_string(str(self.date_from) + ' 00:00:00')),
            ('check_in', '<=', fields.Datetime.from_string(str(self.date_to) + ' 23:59:59')),
        ]
        if self.employee_ids:
            domain.append(('employee_id', 'in', self.employee_ids.ids))
        elif self.department_ids:
            # department_id filter via employee relation
            domain.append(('employee_id.department_id', 'in', self.department_ids.ids))

        attendances = self.env['hr.attendance'].sudo().search(
            domain, order='employee_id, check_in'
        )
        if not attendances:
            raise UserError(_('No attendance records found for the selected filters.'))

        emp_date = defaultdict(lambda: defaultdict(list))
        for att in attendances:
            ci_local = fields.Datetime.context_timestamp(att, att.check_in)
            emp_date[att.employee_id][ci_local.date()].append(att)

        all_dates = sorted({d for ed in emp_date.values() for d in ed})

        # ------------------------------------------------------------------
        # Gazetted / Public Holidays in the date range
        # resource.calendar.leaves with resource_id = False are company-wide
        # ------------------------------------------------------------------
        dt_from = fields.Datetime.from_string(str(self.date_from) + ' 00:00:00')
        dt_to = fields.Datetime.from_string(str(self.date_to) + ' 23:59:59')

        public_leaves = self.env['resource.calendar.leaves'].sudo().search([
            ('resource_id', '=', False),
            ('date_from', '<=', dt_to),
            ('date_to', '>=', dt_from),
        ])
        gazetted_dates = set()
        for lv in public_leaves:
            lv_start = fields.Datetime.context_timestamp(self, lv.date_from).date()
            lv_end = fields.Datetime.context_timestamp(self, lv.date_to).date()
            cur = max(lv_start, self.date_from)
            while cur <= min(lv_end, self.date_to):
                gazetted_dates.add(cur)
                cur += timedelta(days=1)
        gazetted_count = len(gazetted_dates)

        # ------------------------------------------------------------------
        # Paid Leaves per employee (validated leave requests in date range)
        # ------------------------------------------------------------------
        employee_ids_in_report = [emp.id for emp in emp_date]
        paid_leaves_recs = self.env['hr.leave'].sudo().search([
            ('state', '=', 'validate'),
            ('employee_id', 'in', employee_ids_in_report),
            ('date_from', '<=', str(self.date_to) + ' 23:59:59'),
            ('date_to', '>=', str(self.date_from) + ' 00:00:00'),
        ])
        emp_paid_days = defaultdict(float)
        for leave in paid_leaves_recs:
            emp_paid_days[leave.employee_id.id] += leave.number_of_days

        # ------------------------------------------------------------------
        # Joining Date – earliest contract date_start per employee
        # ------------------------------------------------------------------
        emp_joining_date = {}
        try:
            contracts = self.env['hr.contract'].sudo().search([
                ('employee_id', 'in', employee_ids_in_report),
            ], order='date_start asc')
            for contract in contracts:
                eid = contract.employee_id.id
                if eid not in emp_joining_date:
                    emp_joining_date[eid] = contract.date_start
        except Exception:
            pass  # hr.contract may not be installed

        # ------------------------------------------------------------------
        # Build workbook
        # ------------------------------------------------------------------
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('Attendance')

        # --- Formats ---
        hdr = wb.add_format({
            'bold': True, 'bg_color': '#1F497D', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        dhdr = wb.add_format({
            'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
        })
        summary_hdr = wb.add_format({
            'bold': True, 'bg_color': '#375623', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        cell = wb.add_format({'border': 1, 'valign': 'vcenter'})
        date_cell = wb.add_format({
            'border': 1, 'valign': 'vcenter', 'num_format': 'dd-mmm-yyyy',
        })
        tfmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': 'hh:mm:ss'})
        dur_fmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '[h]:mm:ss'})
        grand_time_fmt = wb.add_format({
            'bold': True, 'border': 1, 'valign': 'vcenter',
            'num_format': '[h]:mm:ss', 'bg_color': '#E2EFDA',
        })
        int_cell = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '#,##0'})
        dec_cell = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '#,##0.00'})
        summary_int = wb.add_format({
            'bold': True, 'border': 1, 'valign': 'vcenter',
            'num_format': '#,##0', 'bg_color': '#E2EFDA',
        })
        summary_dec = wb.add_format({
            'bold': True, 'border': 1, 'valign': 'vcenter',
            'num_format': '#,##0.00', 'bg_color': '#E2EFDA',
        })

        # --- Row heights ---
        ws.set_row(0, 32)
        ws.set_row(1, 20)

        # --- Fixed columns (rows 0‑1 merged) ---
        fixed_cols = [
            ('#',            5),
            ('Emp ID',      14),
            ('Name',        24),
            ('Designation', 20),
            ('Department',  20),
            ('Joining\nDate', 14),
        ]
        for i, (label, width) in enumerate(fixed_cols):
            ws.merge_range(0, i, 1, i, label, hdr)
            ws.set_column(i, i, width)

        # --- Per-date columns ---
        col = len(fixed_cols)
        date_col_map = {}
        for d in all_dates:
            date_col_map[d] = col
            ws.merge_range(0, col, 0, col + 2, d.strftime('%d-%b-%Y'), dhdr)
            ws.write(1, col,     'First Check In',  hdr)
            ws.write(1, col + 1, 'Last Check Out',  hdr)
            ws.write(1, col + 2, 'Total Time',      hdr)
            ws.set_column(col,     col + 1, 16)
            ws.set_column(col + 2, col + 2, 12)
            col += 3

        # --- Summary columns (rows 0‑1 merged, green header) ---
        summary_defs = [
            # (key,               label,                   width, is_time)
            ('grand_total',      'Total Time\n(All Days)',  14,   True),
            ('present_days',     'Present\nDays',           10,   False),
            ('gazetted',         'Gazetted\nHolidays',      12,   False),
            ('paid_leaves',      'Paid\nLeaves',            10,   False),
            ('total_hours',      'Total\nHours',            10,   False),
            ('total_worked_days','Total Worked\nDays',      13,   False),
            ('avg_worked_hours', 'Avg Worked\nHours',       13,   False),
        ]
        summary_col = {}
        for i, (key, label, width, _) in enumerate(summary_defs):
            c = col + i
            summary_col[key] = c
            ws.merge_range(0, c, 1, c, label, summary_hdr)
            ws.set_column(c, c, width)

        ws.freeze_panes(2, len(fixed_cols))

        # ------------------------------------------------------------------
        # Data rows
        # ------------------------------------------------------------------
        row = 2
        for seq, (emp, date_map) in enumerate(
            sorted(emp_date.items(), key=lambda x: x[0].name), start=1
        ):
            # Fixed columns
            ws.write(row, 0, seq,                                        cell)
            ws.write(row, 1, emp.x_zk_user_id or '',                    cell)
            ws.write(row, 2, emp.name,                                   cell)
            ws.write(row, 3, emp.job_id.name if emp.job_id else '',      cell)
            ws.write(row, 4, emp.department_id.name if emp.department_id else '', cell)

            joining = emp_joining_date.get(emp.id)
            if joining:
                ws.write_datetime(
                    row, 5,
                    datetime.combine(joining, datetime.min.time()),
                    date_cell,
                )
            else:
                ws.write(row, 5, '', cell)

            grand_total_seconds = 0
            present_days = 0

            # Per-date columns
            for d, start_col in date_col_map.items():
                recs = date_map.get(d, [])
                if recs:
                    first_check_in = min(r.check_in for r in recs)
                    last_check_out = max(
                        (r.check_out for r in recs if r.check_out),
                        default=None,
                    )
                    ci_local = fields.Datetime.context_timestamp(self, first_check_in)
                    ws.write_datetime(row, start_col, ci_local.replace(tzinfo=None), tfmt)
                    if last_check_out:
                        co_local = fields.Datetime.context_timestamp(self, last_check_out)
                        ws.write_datetime(row, start_col + 1, co_local.replace(tzinfo=None), tfmt)
                        day_seconds = (last_check_out - first_check_in).total_seconds()
                        grand_total_seconds += day_seconds
                        present_days += 1
                        # Excel time = fraction of a day
                        ws.write_number(row, start_col + 2, day_seconds / 86400.0, dur_fmt)
                    else:
                        ws.write(row, start_col + 1, '', cell)
                        ws.write(row, start_col + 2, '', cell)
                else:
                    ws.write(row, start_col,     '', cell)
                    ws.write(row, start_col + 1, '', cell)
                    ws.write(row, start_col + 2, '', cell)

            # Derived summary values
            total_hours = grand_total_seconds / 3600.0
            avg_worked_hours = total_hours / present_days if present_days else 0.0
            paid_days = emp_paid_days.get(emp.id, 0)

            # Write summary columns
            if grand_total_seconds:
                ws.write_number(
                    row, summary_col['grand_total'],
                    grand_total_seconds / 86400.0, grand_time_fmt,
                )
            else:
                ws.write(row, summary_col['grand_total'], '', grand_time_fmt)

            ws.write(row, summary_col['present_days'],      present_days,              summary_int)
            ws.write(row, summary_col['gazetted'],          gazetted_count,            int_cell)
            ws.write(row, summary_col['paid_leaves'],       round(paid_days, 2),       dec_cell)
            ws.write(row, summary_col['total_hours'],       round(total_hours, 2),     summary_dec)
            ws.write(row, summary_col['total_worked_days'], present_days,              summary_int)
            ws.write(row, summary_col['avg_worked_hours'],  round(avg_worked_hours, 2), summary_dec)

            row += 1

        wb.close()
        self.write({
            'file_data': base64.b64encode(output.getvalue()),
            'filename': 'Attendance_%s_%s.xlsx' % (self.date_from, self.date_to),
        })
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }