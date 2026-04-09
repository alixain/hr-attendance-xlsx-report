import base64
import io
from collections import defaultdict
from datetime import date, timedelta

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
    _name = 'attendance.report.wizard'
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
        max_pairs = max(
            (len(recs) for ed in emp_date.values() for recs in ed.values()),
            default=1,
        )

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
        cell = wb.add_format({'border': 1, 'valign': 'vcenter'})
        tfmt = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': 'hh:mm:ss'})

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
            span = max_pairs * 2
            if span > 1:
                ws.merge_range(0, col, 0, col + span - 1, d.strftime('%d-%b-%Y'), dhdr)
            else:
                ws.write(0, col, d.strftime('%d-%b-%Y'), dhdr)
            for p in range(max_pairs):
                label_in = 'Check In' if max_pairs == 1 else f'In {p + 1}'
                label_out = 'Check Out' if max_pairs == 1 else f'Out {p + 1}'
                ws.write(1, col + p * 2, label_in, hdr)
                ws.write(1, col + p * 2 + 1, label_out, hdr)
            ws.set_column(col, col + span - 1, 12)
            col += span

        ws.freeze_panes(2, len(fixed))

        row = 2
        for seq, (emp, date_map) in enumerate(
            sorted(emp_date.items(), key=lambda x: x[0].name), start=1
        ):
            ws.write(row, 0, seq, cell)
            ws.write(row, 1, emp.x_zk_user_id or '', cell)
            ws.write(row, 2, emp.name, cell)
            ws.write(row, 3, emp.job_id.name if emp.job_id else '', cell)

            for d, start_col in date_col_map.items():
                records = sorted(date_map.get(d, []), key=lambda r: r.check_in)
                for p, att in enumerate(records[:max_pairs]):
                    ci_l = fields.Datetime.context_timestamp(att, att.check_in)
                    co_l = (fields.Datetime.context_timestamp(att, att.check_out)
                            if att.check_out else None)
                    ws.write_datetime(row, start_col + p * 2, ci_l.replace(tzinfo=None), tfmt)
                    if co_l:
                        ws.write_datetime(row, start_col + p * 2 + 1,
                                          co_l.replace(tzinfo=None), tfmt)
                    else:
                        ws.write(row, start_col + p * 2 + 1, '', cell)
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