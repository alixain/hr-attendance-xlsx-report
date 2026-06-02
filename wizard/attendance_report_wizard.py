import base64
import io
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone

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

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('Attendance')

        hdr = wb.add_format({
            'bold': True, 'bg_color': '#1F497D', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        dhdr = wb.add_format({
            'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
        })
        total_hdr = wb.add_format({
            'bold': True, 'bg_color': '#375623', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        cell = wb.add_format({'border': 1, 'valign': 'vcenter'})
        tfmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': 'hh:mm:ss'})
        dur_fmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '[h]:mm:ss'})
        grand_fmt = wb.add_format({
            'bold': True, 'border': 1, 'valign': 'vcenter',
            'num_format': '[h]:mm:ss', 'bg_color': '#E2EFDA',
        })

        ws.set_row(0, 30)
        ws.set_row(1, 20)
        fixed = ['#', 'Emp ID', 'Name', 'Designation']
        for i, h in enumerate(fixed):
            ws.merge_range(0, i, 1, i, h, hdr)
        ws.set_column(0, 0, 5)
        ws.set_column(1, 1, 14)
        ws.set_column(2, 2, 24)
        ws.set_column(3, 3, 20)

        col = len(fixed)
        date_col_map = {}
        for d in all_dates:
            date_col_map[d] = col
            ws.merge_range(0, col, 0, col + 2, d.strftime('%d-%b-%Y'), dhdr)
            ws.write(1, col, 'First Check In', hdr)
            ws.write(1, col + 1, 'Last Check Out', hdr)
            ws.write(1, col + 2, 'Total Time', hdr)
            ws.set_column(col, col + 1, 16)
            ws.set_column(col + 2, col + 2, 12)
            col += 3

        # Grand total column
        grand_total_col = col
        ws.merge_range(0, grand_total_col, 1, grand_total_col, 'Total Time\n(All Days)', total_hdr)
        ws.set_column(grand_total_col, grand_total_col, 14)

        ws.freeze_panes(2, len(fixed))

        row = 2
        for seq, (emp, date_map) in enumerate(
            sorted(emp_date.items(), key=lambda x: x[0].name), start=1
        ):
            ws.write(row, 0, seq, cell)
            ws.write(row, 1, emp.x_zk_user_id or '', cell)
            ws.write(row, 2, emp.name, cell)
            ws.write(row, 3, emp.job_id.name if emp.job_id else '', cell)

            grand_total_seconds = 0

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
                        # Excel stores time as fraction of a day
                        ws.write_number(row, start_col + 2, day_seconds / 86400.0, dur_fmt)
                    else:
                        ws.write(row, start_col + 1, '', cell)
                        ws.write(row, start_col + 2, '', cell)
                else:
                    ws.write(row, start_col, '', cell)
                    ws.write(row, start_col + 1, '', cell)
                    ws.write(row, start_col + 2, '', cell)

            # Write grand total for this employee
            if grand_total_seconds:
                ws.write_number(row, grand_total_col, grand_total_seconds / 86400.0, grand_fmt)
            else:
                ws.write(row, grand_total_col, '', grand_fmt)

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