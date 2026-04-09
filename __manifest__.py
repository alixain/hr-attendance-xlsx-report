{
    'name': 'HR Attendance XLSX Report',
    'version': '18.0.1.0.0',
    'category': 'Human Resources/Attendances',
    'summary': 'Export attendance with daily check-in/out columns to XLSX',
    'description': '''
        <p>This module extends Odoo's HR Attendance functionality by providing comprehensive XLSX export capabilities for attendance reports.</p>
        
        <h3>Key Features:</h3>
        <ul>
            <li><strong>Daily Check-in/Check-out Columns:</strong> Export attendance data with separate columns for each day's check-in and check-out times</li>
            <li><strong>Employee-wise Reports:</strong> Generate detailed attendance reports organized by employee</li>
            <li><strong>Date Range Filtering:</strong> Export attendance data for specific date ranges</li>
            <li><strong>Department-wise Filtering:</strong> Filter reports by departments for better organization</li>
            <li><strong>Excel Format:</strong> Professional XLSX format compatible with Microsoft Excel and other spreadsheet applications</li>
        </ul>
        
        <h3>Use Cases:</h3>
        <ul>
            <li>Monthly attendance reporting for payroll processing</li>
            <li>HR analytics and employee performance tracking</li>
            <li>Compliance reporting for labor regulations</li>
            <li>Management dashboards and KPI tracking</li>
        </ul>
        
        <h3>Easy to Use:</h3>
        <p>The module adds a simple wizard accessible from the HR Attendance menu. Select your date range, departments, and export to XLSX with one click.</p>
        
        <h3>Compatibility:</h3>
        <p>Designed for Odoo 18.0 and compatible with standard HR Attendance module.</p>
    ''',
    'author': 'Aspire Analytica',
    'website': 'https://www.aspireanalytica.com',
    'depends': ['hr_attendance'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/attendance_report_wizard_views.xml',
    ],
    'license': 'LGPL-3',
    'auto_install': False,
    'installable': True,
}
