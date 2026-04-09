# HR Attendance XLSX Report

## Odoo App Store Publishing

This module is prepared for submission to the Odoo App Store. All required files and metadata have been configured according to Odoo publishing guidelines.

### Module Structure
```
hr_attendance_xlsx_report/
├── __manifest__.py          # Module manifest with App Store metadata
├── __init__.py             # Python package initialization
├── README.md               # This documentation file
├── icon.png                # 256x256 module icon
├── security/
│   └── ir.model.access.csv # Access rights
├── static/
│   └── description/        # Marketing assets for App Store
│       ├── index.html      # HTML description page
│       ├── banner.png      # 1920x720 banner image
│       └── screenshot_*.png # Feature screenshots
├── views/
│   └── wizard_views.xml    # Menu and action definitions
└── wizard/                 # Report wizard implementation
    ├── __init__.py
    ├── attendance_report_wizard.py
    └── attendance_report_wizard_views.xml
```

## Marketing Assets for Odoo App Store

### Required Files

#### Icon (`icon.png`)
- **Location**: Module root
- **Dimensions**: 256x256 pixels
- **Format**: PNG
- **Purpose**: Module icon displayed in Odoo App Store and module list

#### Banner (`static/description/banner.png`)
- **Dimensions**: 1920x720 pixels (16:9 aspect ratio)
- **Format**: PNG
- **Purpose**: Hero banner displayed on the module's App Store page

#### Screenshots (`static/description/screenshot_*.png`)
- **Files**: `screenshot_1.png`, `screenshot_2.png`, `screenshot_3.png`
- **Dimensions**: Recommended 1920x1080 or 1280x720 pixels
- **Format**: PNG
- **Purpose**: Demonstrate module functionality in Odoo App Store

#### HTML Description (`static/description/index.html`)
- **Purpose**: Rich HTML description page for the App Store
- **Content**: Detailed feature list, use cases, and screenshots

### Manifest Configuration

The `__manifest__.py` file includes:
- ✅ Comprehensive description with HTML formatting
- ✅ Images array pointing to screenshot files
- ✅ Proper metadata (author, website, license)
- ✅ Correct dependencies and data files

### How to Create Marketing Assets

1. **Design Software**: Use Adobe Photoshop, GIMP, Canva, or Figma
2. **Odoo Branding**: Follow Odoo's design guidelines (clean, professional, blue color scheme)
3. **Content Focus**:
   - Icon: Simple, recognizable symbol for attendance reporting
   - Banner: Eye-catching header showing the module's value proposition
   - Screenshots: Real functionality showing:
     - Wizard interface for report generation
     - Sample XLSX report output
     - Integration with HR Attendance module

### Odoo App Store Submission Checklist

- [ ] Replace placeholder `icon.png` with actual 256x256 PNG icon
- [ ] Replace placeholder `banner.png` with actual 1920x720 PNG banner
- [ ] Replace placeholder screenshots with actual module screenshots
- [ ] Test all images display correctly in Odoo
- [ ] Verify manifest description is comprehensive and accurate
- [ ] Test module installation and functionality in target Odoo version (18.0)
- [ ] Ensure code follows Odoo coding standards
- [ ] Verify license compatibility (LGPL-3)
- [ ] Test module on clean Odoo database
- [ ] Prepare demo data if needed for screenshots

### Module Description

Export HR attendance data to XLSX format with daily check-in/out columns. This module extends the standard HR Attendance functionality by providing comprehensive reporting capabilities for attendance management.

**Key Features:**
- Daily check-in/check-out columns in XLSX export
- Employee-wise and department-wise filtering
- Date range selection
- Professional Excel-compatible format
- Easy-to-use wizard interface

**Compatibility:** Odoo 18.0, requires `hr_attendance` module.