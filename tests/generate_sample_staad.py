"""
Authentic STAAD.Pro export generator with four worksheets (R-L, R-R, JB-F, JB-B).
Includes exact reference validation benchmark:
  - Node 136, Member 2028:
      • Moment-Z = +1526.629 at L/C 1257
      • Shear-Y  = -146.266  at L/C 1005
  - Multiple structural nodes with 2, 3, 4, 5, 6, 7 connected members.
  - Multi-sheet framing and intentional cross-worksheet ties.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


def create_sample_staad_4sheets(
    output_path: str = "d:\\Automation\\sample_data\\sample_staad_4sheets.xlsx",
) -> str:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font = Font(name="Arial", size=10, bold=True, color="000000")
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    data_font = Font(name="Arial", size=9)

    headers = [
        "Beam",
        "L/C",
        "Node",
        "Axial Force (kN)",
        "Shear-Y (kN)",
        "Shear-Z (kN)",
        "Torsion (kN-m)",
        "Moment-Y (kN-m)",
        "Moment-Z (kN-m)",
    ]

    # Include authentic benchmark load cases from screen recording (1005, 1030, 1115, 1257)
    load_cases = [1, 2, 3, 1005, 1030, 1115, 1257, 101, 104, 105]

    sheet_data_specs = {
        "R-L": [
            (10, 101),
            (20, 101),
            (30, 103),
            (40, 105),
            (50, 107),
            (60, 109),
            (136, 2028),  # Benchmark Member 2028 at Node 136 (R-L)
        ],
        "R-R": [
            (10, 102),
            (30, 104),
            (40, 106),
            (50, 108),
            (60, 110),
            (60, 111),
            (136, 2502),  # Framing into Node 136 (R-R)
        ],
        "JB-F": [
            (20, 201),
            (30, 202),
            (40, 203),
            (40, 204),
            (50, 205),
            (50, 206),
            (60, 207),
            (60, 208),
            (136, 2077),  # Framing into Node 136 (JB-F)
        ],
        "JB-B": [
            (20, 301),
            (30, 302),
            (40, 303),
            (50, 304),
            (50, 305),
            (60, 306),
            (60, 307),
            (136, 2018),  # Framing into Node 136 (JB-B)
        ],
    }

    for sheet_name, assignments in sheet_data_specs.items():
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True

        ws.cell(row=1, column=1, value="STAAD.Pro Output Report").font = Font(name="Arial", size=11, bold=True)
        ws.cell(row=2, column=1, value=f"Member End Forces - Group {sheet_name}").font = Font(name="Arial", size=10, italic=True)

        header_row_idx = 4
        for col_idx, h in enumerate(headers, start=1):
            c = ws.cell(row=header_row_idx, column=col_idx, value=h)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center", vertical="center")

        curr_row = 5
        for node_id, member_id in assignments:
            for lc in load_cases:
                # Base formulas for forces
                base_fx = -200.0 + (node_id * 1.5) - (member_id % 30) * 8.0 + (lc % 100) * 2.5
                base_fy = -80.0 + (node_id * 1.2) + (member_id % 25) * 4.5 - (lc % 100) * 2.0
                base_mz = -110.0 + (node_id * 1.8) - (member_id % 35) * 5.0 + (lc % 100) * 4.0

                # Benchmark Node 136 / Member 2028 (Exact from screen recording: Recording 2026-11-13 162544.mp4):
                if node_id == 136 and member_id == 2028:
                    if lc == 1257:
                        # Row 516 in video: Max Moment-Z AND Max Shear-Y
                        base_mz = 1526.629
                        base_fy = 230.941
                        base_fx = -4.841
                        base_fz = 0.136
                        base_mx = -0.003
                        base_my = -0.661
                    elif lc == 1005:
                        # Row 12 in video: Min Moment-Z AND Min Shear-Y
                        base_mz = -919.990
                        base_fy = -146.266
                        base_fx = -59.469
                        base_fz = -0.117
                        base_mx = 0.001
                        base_my = 0.505
                    elif lc == 1030:
                        # Row 62 in video: Max Axial Force
                        base_fx = 16.094
                        base_fy = 79.836
                        base_mz = 548.898
                        base_fz = 0.536
                        base_mx = -0.001
                        base_my = -0.577
                    elif lc == 1115:
                        # Row 232 in video: Min Axial Force
                        base_fx = -72.126
                        base_fy = 13.049
                        base_mz = 96.224
                        base_fz = 0.032
                        base_mx = 0.000
                        base_my = -0.201
                    else:
                        base_mz = 100.000
                        base_fy = 20.000
                        base_fx = -40.000

                # Node 10 extremes:
                if node_id == 10 and member_id == 101 and lc == 101:
                    base_mz = 820.000
                    base_fy = 320.000
                if node_id == 10 and member_id == 102 and lc == 104:
                    base_mz = -420.000
                    base_fx = -750.000
                if node_id == 10 and member_id == 101 and lc == 105:
                    base_fx = 610.000

                # Node 30 intentional cross-worksheet ties:
                if node_id == 30 and lc == 105 and member_id in (103, 202):
                    base_mz = 650.000
                if node_id == 30 and lc == 105 and member_id in (104, 302):
                    base_fy = -450.000

                base_fz = round((node_id * 0.3 - (lc % 100) * 0.15), 3)
                base_mx = round((member_id * 0.04 - (lc % 100) * 0.08), 3)
                base_my = round((node_id * 0.2 + (lc % 100) * 0.1), 3)

                row_vals = [
                    member_id,
                    lc,
                    node_id,
                    round(base_fx, 3),
                    round(base_fy, 3),
                    base_fz,
                    base_mx,
                    base_my,
                    round(base_mz, 3),
                ]

                for col_idx, val in enumerate(row_vals, start=1):
                    c = ws.cell(row=curr_row, column=col_idx, value=val)
                    c.font = data_font
                    if col_idx in (1, 2, 3):
                        c.alignment = Alignment(horizontal="center")
                    else:
                        c.alignment = Alignment(horizontal="right")
                        c.number_format = "0.000"

                curr_row += 1

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    path = create_sample_staad_4sheets()
    print(f"Sample 4-worksheet STAAD export generated at: {path}")
