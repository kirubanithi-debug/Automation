"""
Configuration constants, aliases, and styling rules for STAAD.Pro Governing Force Automation.
AI/ML Data Interpretation + Deterministic Engineering Calculations.
"""

from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Application Metadata
APP_NAME = "AI/ML STAAD.Pro Governing Force Extraction System"
APP_VERSION = "4.0.0"

# Target member-force worksheets in STAAD.Pro export
DEFAULT_TARGET_SHEETS = ["R-L", "R-R", "JB-F", "JB-B"]

# Canonical Force Types
FORCE_MOMENT_Z = "Moment-Z (kN-m)"
FORCE_SHEAR_Y = "Shear-Y (kN)"
FORCE_AXIAL = "Axial Force (kN)"

TARGET_FORCE_TYPES = [
    FORCE_MOMENT_Z,
    FORCE_SHEAR_Y,
    FORCE_AXIAL,
]

# Output Sheet Names (Excel prohibits '/' in sheet names)
SHEET1_GOVERNING_SUMMARY = "Governing Summary"
SHEET2_NODE_WISE_UNIQUE_LCS = "Node-Wise Unique LC"
SHEET3_GOVERNING_SOURCE_RECORDS = "Governing Source Records"
SHEET4_COMBINED_SOURCE_DATA = "Combined Source Data"
SHEET5_VALIDATION_REPORT = "Validation Report"

# Aliases for backwards compatibility
SHEET1_MEMBER_GOVERNING_SUMMARY = SHEET1_GOVERNING_SUMMARY
SHEET2_UNIQUE_LCS_BY_NODE = SHEET2_NODE_WISE_UNIQUE_LCS
SHEET2_CONNECTED_MEMBERS = SHEET2_NODE_WISE_UNIQUE_LCS
SHEET3_UNIQUE_LCS = SHEET2_NODE_WISE_UNIQUE_LCS
SHEET4_DETAILED_RECORDS = SHEET3_GOVERNING_SOURCE_RECORDS
SHEET4_COMBINED_ORIGINAL_DATA = SHEET4_COMBINED_SOURCE_DATA
SHEET5_AI_VALIDATION = SHEET5_VALIDATION_REPORT
SHEET5_ML_VALIDATION_REPORT = SHEET5_VALIDATION_REPORT

ALL_OUTPUT_SHEETS = [
    SHEET1_GOVERNING_SUMMARY,
    SHEET2_NODE_WISE_UNIQUE_LCS,
    SHEET3_GOVERNING_SOURCE_RECORDS,
    SHEET4_COMBINED_SOURCE_DATA,
    SHEET5_VALIDATION_REPORT,
]

# Column aliases for STAAD member end forces tables
COLUMN_ALIASES = {
    "member": [
        "beam", "beam no", "beam id", "beam_no", "beam_id",
        "member", "member no", "member id", "member_no", "member_id",
        "element", "elem", "elem no"
    ],
    "lc": [
        "l/c", "lc", "load case", "load_case", "loadcase",
        "loading", "load combination", "load", "l_c"
    ],
    "node": [
        "node", "node no", "node id", "node number", "node_id", "node_no",
        "joint", "joint no", "joint id", "pt", "point"
    ],
    "axial": [
        "axial force (kn)", "axial force", "axial", "axial (kn)",
        "fx", "fx (kn)", "fx(kn)", "fx kn", "p", "p (kn)", "p(kn)", "axial_force"
    ],
    "shear_y": [
        "shear-y (kn)", "shear-y(kn)", "shear y (kn)", "shear-y", "shear y",
        "fy", "fy (kn)", "fy(kn)", "fy kn", "shear_y"
    ],
    "moment_z": [
        "moment-z (kn-m)", "moment-z(kn-m)", "moment z (kn-m)", "moment-z (kn.m)",
        "moment-z", "moment z", "mz", "mz (kn-m)", "mz(kn-m)", "mz (kn.m)",
        "mz (knm)", "mz(knm)", "mz kn-m", "bending-z", "bending z", "mz knm",
        "moment z (kn.m)", "moment_z"
    ]
}

# Excel Styling Configuration
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

SUBHEADER_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
SUBHEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

NODE_BANNER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
NODE_BANNER_FONT = Font(name="Calibri", size=11, bold=True, color="1F4E79")

# Highlighting: Maximum Value (Soft Green)
MAX_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
MAX_FONT = Font(name="Calibri", size=10, bold=True, color="006100")
MAX_ROW_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

# Highlighting: Minimum Value (Soft Red / Coral)
MIN_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
MIN_FONT = Font(name="Calibri", size=10, bold=True, color="9C0006")
MIN_ROW_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

# Video workflow highlighting: Governing Rows (Yellow)
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
YELLOW_FONT = Font(name="Calibri", size=10, bold=True, color="000000")

# Both Max and Min (in case of single record or max == min)
BOTH_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
BOTH_FONT = Font(name="Calibri", size=10, bold=True, color="7F6000")

# Validation Status Styling
VALID_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
VALID_FONT = Font(name="Calibri", size=10, bold=True, color="006100")
INVALID_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
INVALID_FONT = Font(name="Calibri", size=10, bold=True, color="9C0006")

# Standard borders
THIN_GRAY = Side(border_style="thin", color="D9D9D9")
MEDIUM_DARK = Side(border_style="medium", color="1F4E79")
DOUBLE_BOTTOM = Side(border_style="double", color="1F4E79")

CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
HEADER_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=MEDIUM_DARK, bottom=MEDIUM_DARK)
SUMMARY_TOTAL_BORDER = Border(top=THIN_GRAY, bottom=DOUBLE_BOTTOM)

# Standard Alignments
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_HEADER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Standard Number Formats
NUMBER_FORMAT_DECIMAL = "0.000"
NUMBER_FORMAT_INTEGER = "0"
