import base64
import io
from collections import defaultdict
from datetime import date

from odoo import api, fields, models, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    raise UserError(_("xlsxwriter is required. Install it with: pip install xlsxwriter"))


class AttendanceReportWizard(models.TransientModel):
    _name = 'attendance.report.wizard'
    _description = 'Attendance XLSX Report Wizard'

    date_from = fields.Date(
        string='Date From',
        required=True,
        default=lambda self: date.today().replace(day=1).replace(
            month=date.today().month - 1 if date.today().month > 1 else 12,
            year=date.today().year if date.today().month > 1 else date.today().year - 1,
        ),
    )
    date_to = fields.Date(
        string='Date To',
        required=True,
        default=lambda self: date.today().replace(day=1) - __import__('datetime').timedelta(days=1),
    )
    employee_ids = fields.Many2many(
        'hr.employee',
        string='Employees',
        help='Leave empty for all employees',
    )
    department_ids = fields.Many2many(
        'hr.department',
        string='Departments',
        help='Leave empty for all departments',
    )
    filename = fields.Char(string='Filename', readonly=True)
    file_data = fields.Binary(string='File', readonly=True)

    @api.onchange('date_from', 'date_to')
    def _onchange_dates(self):
        if self.date_from and self.date_to and self.date_from > self.date_to:
            return {'warning': {
                'title': _('Invalid Dates'),
                'message': _('Date From must be before Date To.'),
            }}

    def _get_domain(self):
        domain = [
            ('check_in', '>=', fields.Datetime.from_string(str(self.date_from) + ' 00:00:00')),
            ('check_in', '<=', fields.Datetime.from_string(str(self.date_to) + ' 23:59:59')),
        ]
        if self.employee_ids:
            domain.append(('employee_id', 'in', self.employee_ids.ids))
        if self.department_ids:
            domain.append(('department_id', 'in', self.department_ids.ids))
        return domain

    def action_generate(self):
        self.ensure_one()
        if not self.date_from or not self.date_to:
            raise UserError(_('Please set both Date From and Date To.'))

        attendances = self.env['hr.attendance'].search(
            self._get_domain(),
            order='employee_id, check_in',
        )
        if not attendances:
            raise UserError(_('No attendance records found for the selected filters.'))

        # Group attendances: {employee: {date: [records]}}
        emp_date_map = defaultdict(lambda: defaultdict(list))
        for att in attendances:
            # Convert to user/company TZ
            tz = self.env.user.tz or 'UTC'
            check_in_local = fields.Datetime.context_timestamp(att, att.check_in)
            att_date = check_in_local.date()
            emp_date_map[att.employee_id][att_date].append(att)

        # Collect all unique dates sorted
        all_dates = sorted({d for emp_data in emp_date_map.values() for d in emp_data})
        max_pairs = max(
            (len(records) for emp_data in emp_date_map.values() for records in emp_data.values()),
            default=1,
        )

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet('Attendance')

        # Formats
        hdr_fmt = wb.add_format({
            'bold': True, 'bg_color': '#1F497D', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        date_hdr_fmt = wb.add_format({
            'bold': True, 'bg_color': '#2E75B6', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter',
        })
        cell_fmt = wb.add_format({'border': 1, 'valign': 'vcenter'})
        time_fmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': 'hh:mm:ss'})
        date_fmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': 'yyyy-mm-dd'})

        # Build headers
        fixed_headers = ['#', 'Employee ID', 'Name', 'Designation']
        col = len(fixed_headers)

        ws.set_row(0, 30)
        for i, h in enumerate(fixed_headers):
            ws.write(0, i, h, hdr_fmt)
        ws.set_column(0, 0, 5)
        ws.set_column(1, 1, 15)
        ws.set_column(2, 2, 25)
        ws.set_column(3, 3, 20)

        # Date columns: for each date, pair columns per max_pairs
        date_col_map = {}  # date -> start col
        for d in all_dates:
            date_col_map[d] = col
            date_str = d.strftime('%d-%b-%Y')
            if max_pairs == 1:
                ws.merge_range(0, col, 0, col + 1, date_str, date_hdr_fmt)
                ws.write(1, col, 'Check In', hdr_fmt)
                ws.write(1, col + 1, 'Check Out', hdr_fmt)
                ws.set_column(col, col + 1, 12)
                col += 2
            else:
                ws.merge_range(0, col, 0, col + max_pairs * 2 - 1, date_str, date_hdr_fmt)
                for p in range(max_pairs):
                    ws.write(1, col + p * 2, f'In {p+1}', hdr_fmt)
                    ws.write(1, col + p * 2 + 1, f'Out {p+1}', hdr_fmt)
                ws.set_column(col, col + max_pairs * 2 - 1, 12)
                col += max_pairs * 2

        total_cols = col

        # Write data rows (start at row 2 if double header, else row 1)
        header_rows = 2 if all_dates else 1
        row = header_rows
        seq = 1

        for emp, date_map in sorted(emp_date_map.items(), key=lambda x: x[0].name):
            # Get ZKTeco ID — stored in barcode field on hr.employee
            zkteco_id = emp.x_zk_user_id or ''
            designation = emp.job_id.name if emp.job_id else ''

            ws.write(row, 0, seq, cell_fmt)
            ws.write(row, 1, zkteco_id, cell_fmt)
            ws.write(row, 2, emp.name, cell_fmt)
            ws.write(row, 3, designation, cell_fmt)

            for d, start_col in date_col_map.items():
                records = sorted(date_map.get(d, []), key=lambda r: r.check_in)
                for p, att in enumerate(records[:max_pairs]):
                    tz = self.env.user.tz or 'UTC' #pakistan standard time
                    ci_local = fields.Datetime.context_timestamp(att, att.check_in)
                    co_local = fields.Datetime.context_timestamp(att, att.check_out) if att.check_out else None

                    ci_col = start_col + p * 2
                    co_col = ci_col + 1

                    ws.write_datetime(row, ci_col, ci_local.replace(tzinfo=None), time_fmt)
                    if co_local:
                        ws.write_datetime(row, co_col, co_local.replace(tzinfo=None), time_fmt)
                    else:
                        ws.write(row, co_col, '', cell_fmt)

            row += 1
            seq += 1

        # Freeze panes after fixed headers
        ws.freeze_panes(header_rows, len(fixed_headers))
        wb.close()

        xlsx_data = base64.b64encode(output.getvalue())
        fname = 'Attendance_%s_%s.xlsx' % (self.date_from, self.date_to)
        self.write({'file_data': xlsx_data, 'filename': fname})

        return {
            'type': 'ir.actions.act_window',
            'res_model': 'attendance.report.wizard',
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
            'context': self.env.context,
        }
