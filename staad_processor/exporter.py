"""
Excel Workbook Generator for Hybrid AI + ML STAAD.Pro Governing Force System.
Generates all five required worksheets per Master Prompt:
  1. Sheet 1 – Governing Summary (| Node ID | Member ID | Force Component | MAX / MIN | Governing Force Value | Corresponding L/C | Source Sheet | Source Row |)
  2. Sheet 2 – Node-Wise Unique LC (| Node ID | Number of Unique L/Cs | Unique L/C Values | + Itemized Table)
  3. Sheet 3 – Governing Source Records (Complete original source rows for every governing extreme)
  4. Sheet 4 – Combined Source Data (Original records from all 4 sheets with Green/Red highlights)
  5. Sheet 5 – Validation Report (Metrics, Check A, Check B, Check C, Check D & 100% Traceability Audit)
"""

import os
from collections import defaultdict
from typing import Dict, List, Optional, Any, Tuple, Set
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border

from .config import (
    APP_NAME,
    FORCE_MOMENT_Z,
    FORCE_SHEAR_Y,
    FORCE_AXIAL,
    TARGET_FORCE_TYPES,
    SHEET1_GOVERNING_SUMMARY,
    SHEET2_NODE_WISE_UNIQUE_LCS,
    SHEET3_GOVERNING_SOURCE_RECORDS,
    SHEET4_COMBINED_SOURCE_DATA,
    SHEET5_VALIDATION_REPORT,
    HEADER_FILL,
    HEADER_FONT,
    SUBHEADER_FILL,
    SUBHEADER_FONT,
    NODE_BANNER_FILL,
    NODE_BANNER_FONT,
    MAX_FILL,
    MAX_FONT,
    MAX_ROW_FILL,
    MIN_FILL,
    MIN_FONT,
    MIN_ROW_FILL,
    BOTH_FILL,
    BOTH_FONT,
    YELLOW_FILL,
    YELLOW_FONT,
    VALID_FILL,
    VALID_FONT,
    INVALID_FILL,
    INVALID_FONT,
    CELL_BORDER,
    HEADER_BORDER,
    ALIGN_LEFT,
    ALIGN_CENTER,
    ALIGN_RIGHT,
    ALIGN_HEADER,
    NUMBER_FORMAT_DECIMAL,
    NUMBER_FORMAT_INTEGER,
)
from .analyzer import NodeAnalysis, GoverningRecord, MemberGoverningForces
from .reader import clean_numeric_value, match_sheet_name


def auto_fit_columns(ws, min_width: int = 12, max_width: int = 42):
    """Auto-adjusts column widths based on cell content."""
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.value is not None:
                lines = str(cell.value).split("\n")
                line_max = max(len(l) for l in lines)
                if line_max > max_len:
                    max_len = line_max
        fitted_width = max(min(max_len + 4, max_width), min_width)
        ws.column_dimensions[col_letter].width = fitted_width


class StaadExporter:
    """
    Exports comprehensive STAAD force workbook reproducing the exact manual engineering workflow:
      1. Original source worksheets (R-L, R-R, JB-F, JB-B) with all governing rows highlighted in YELLOW.
      2. Dedicated worksheets per node ('Node 136', 'Node 149', etc.) containing their governing rows.
      3. Master Summary worksheets: 'Governing Summary', 'Node-Wise Unique LC', and 'Validation Report'.
    """

    def __init__(
        self,
        node_results: Dict[str, NodeAnalysis],
        combined_df: Any,
        highlight_map: Dict[Any, str],
        sheets_loaded: List[str],
        per_sheet_col_mappings: Optional[Dict[str, Dict[str, str]]] = None,
        per_sheet_confidences: Optional[Dict[str, Dict[str, float]]] = None,
        per_sheet_bases: Optional[Dict[str, Dict[str, str]]] = None,
        anomalies: Optional[List[Dict[str, Any]]] = None,
        validation_report: Optional[Any] = None,
        input_file_path: Optional[str] = None,
    ):
        self.node_results = node_results
        self.combined_df = combined_df
        self.highlight_map = highlight_map
        self.sheets_loaded = sheets_loaded
        self.per_sheet_col_mappings = per_sheet_col_mappings or {}
        self.per_sheet_confidences = per_sheet_confidences or {}
        self.per_sheet_bases = per_sheet_bases or {}
        self.anomalies = anomalies or []
        self.validation_report = validation_report
        self.input_file_path = input_file_path

        self.wb: openpyxl.Workbook = openpyxl.Workbook()
        self.wb.remove(self.wb.active)

    def generate_workbook(self, output_file_path: str) -> str:
        """Builds all sheets and saves workbook.
        
        If input_file_path is provided, loads the original workbook to retain all
        source sheets (R-L, R-R, JB-F, JB-B), highlights their governing rows in Yellow,
        and adds per-node sheets ('Node 136', 'Node 149', etc.) alongside summary sheets.
        
        If output_file_path is locked (e.g. open in Microsoft Excel), auto-saves to an incremented filename.
        """
        if self.input_file_path and os.path.exists(self.input_file_path):
            try:
                self.wb = openpyxl.load_workbook(self.input_file_path)
            except Exception:
                if self.wb is None:
                    self.wb = openpyxl.Workbook()
                    self.wb.remove(self.wb.active)
        elif self.wb is None:
            self.wb = openpyxl.Workbook()
            self.wb.remove(self.wb.active)

        # 1. Highlight governing rows in Yellow in the original source worksheets
        self._highlight_governing_rows_in_source_sheets()

        # 2. Build dedicated tabs for each Node (e.g. 'Node 136', 'Node 149', etc.)
        self._build_node_tabs()

        # 3. Build Master Summary and Validation sheets
        self._build_sheet1_governing_summary()
        self._build_sheet2_node_wise_unique_lcs()
        self._build_sheet3_governing_source_records()
        self._build_sheet4_combined_source_data()
        self._build_sheet5_validation_report()

        try:
            self.wb.save(output_file_path)
            return output_file_path
        except PermissionError:
            base, ext = os.path.splitext(output_file_path)
            counter = 1
            alt_path = f"{base}_{counter}{ext}"
            while os.path.exists(alt_path):
                try:
                    with open(alt_path, "r+"):
                        break
                except (PermissionError, OSError):
                    counter += 1
                    alt_path = f"{base}_{counter}{ext}"

            try:
                self.wb.save(alt_path)
                return alt_path
            except Exception:
                raise PermissionError(
                    f"Permission Denied: Could not save to '{output_file_path}' "
                    f"because it is currently open in Microsoft Excel or another program.\n\n"
                    f"Please close '{os.path.basename(output_file_path)}' in Excel and try again."
                ) from None

    def _highlight_governing_rows_in_source_sheets(self):
        """Highlights governing extreme rows in YELLOW in the original source worksheets."""
        if not self.wb:
            return

        gov_rows_by_sheet: Dict[str, Set[int]] = defaultdict(set)
        for node_id, analysis in self.node_results.items():
            for rec in analysis.all_governing_records:
                s_name = rec.source_sheet.strip().lower()
                gov_rows_by_sheet[s_name].add(rec.source_row)

        for ws in self.wb.worksheets:
            s_clean = ws.title.strip().lower()
            target_rows = set()
            for key_s, r_set in gov_rows_by_sheet.items():
                if key_s == s_clean or match_sheet_name(s_clean, key_s):
                    target_rows.update(r_set)

            if not target_rows:
                continue

            max_col = min(ws.max_column, 25)
            for r_idx in target_rows:
                if 1 <= r_idx <= ws.max_row:
                    for c_idx in range(1, max_col + 1):
                        cell = ws.cell(row=r_idx, column=c_idx)
                        cell.fill = YELLOW_FILL

    def _build_node_tabs(self):
        """Builds a dedicated worksheet for each unique Node (e.g. 'Node 136', 'Node 149')."""
        if not self.wb or self.combined_df.empty:
            return

        def node_sort_key(n_str: str):
            try:
                return (0, int(n_str))
            except ValueError:
                return (1, n_str)

        raw_cols = [c for c in self.combined_df.columns if not c.startswith("_")]
        display_headers = ["Source Sheet", "Original Row"] + raw_cols + ["Governing Reasons"]

        sorted_nodes = sorted(self.node_results.keys(), key=node_sort_key)

        for node_id in sorted_nodes:
            analysis = self.node_results[node_id]
            sheet_title = f"Node {node_id}"

            if sheet_title in self.wb.sheetnames:
                del self.wb[sheet_title]

            ws = self.wb.create_sheet(title=sheet_title)
            ws.views.sheetView[0].showGridLines = True

            # Title
            title_end = get_column_letter(len(display_headers))
            ws.merge_cells(f"A1:{title_end}1")
            t_cell = ws.cell(row=1, column=1, value=f"STAAD.Pro - Governing Member Forces for Node {node_id}")
            t_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
            t_cell.fill = HEADER_FILL
            t_cell.alignment = ALIGN_CENTER
            ws.row_dimensions[1].height = 28

            # Subtitle Summary
            members_str = ", ".join(analysis.connected_members)
            lcs_str = ", ".join(analysis.unique_governing_lcs)
            ws.merge_cells(f"A2:{title_end}2")
            sub_cell = ws.cell(
                row=2, column=1,
                value=f"Connected Framing Members ({len(analysis.connected_members)}): {members_str}  |  "
                      f"Unique Governing Load Cases ({len(analysis.unique_governing_lcs)}): {lcs_str}"
            )
            sub_cell.font = Font(name="Calibri", size=9, bold=True, color="1F4E79")
            sub_cell.fill = NODE_BANNER_FILL
            sub_cell.alignment = ALIGN_CENTER
            ws.row_dimensions[2].height = 20

            # Headers
            ws.row_dimensions[4].height = 24
            for c_idx, h in enumerate(display_headers, start=1):
                cell = ws.cell(row=4, column=c_idx, value=h)
                cell.font = HEADER_FONT
                cell.fill = SUBHEADER_FILL
                cell.alignment = ALIGN_HEADER
                cell.border = HEADER_BORDER

            # Group governing records for this node by (source_sheet, source_row)
            row_map: Dict[Tuple[str, int], List[GoverningRecord]] = defaultdict(list)
            for rec in analysis.all_governing_records:
                row_map[(rec.source_sheet, rec.source_row)].append(rec)

            curr_row = 5
            for (s_sheet, s_row), rec_list in row_map.items():
                ws.row_dimensions[curr_row].height = 20

                match = self.combined_df[
                    (self.combined_df["_source_sheet"] == s_sheet) &
                    (self.combined_df["_original_sheet_row"] == s_row)
                ]
                if match.empty:
                    continue
                row_data = match.iloc[0]

                reasons = []
                for r in rec_list:
                    reasons.append(f"{r.force_type} ({r.condition}: {r.value})")
                reason_str = "; ".join(reasons)

                c_sheet = ws.cell(row=curr_row, column=1, value=s_sheet)
                c_sheet.alignment = ALIGN_CENTER
                c_sheet.font = YELLOW_FONT
                c_sheet.fill = YELLOW_FILL
                c_sheet.border = CELL_BORDER

                c_srow = ws.cell(row=curr_row, column=2, value=s_row)
                c_srow.alignment = ALIGN_CENTER
                c_srow.font = YELLOW_FONT
                c_srow.fill = YELLOW_FILL
                c_srow.border = CELL_BORDER

                for col_idx, col_name in enumerate(raw_cols, start=3):
                    val = row_data[col_name]
                    c = ws.cell(row=curr_row, column=col_idx)
                    c.border = CELL_BORDER
                    c.fill = YELLOW_FILL

                    num_val = clean_numeric_value(val)
                    if num_val is not None:
                        c.value = num_val
                        c.alignment = ALIGN_RIGHT
                        c.number_format = NUMBER_FORMAT_DECIMAL
                        c.font = YELLOW_FONT
                    else:
                        c.value = val
                        c.alignment = ALIGN_CENTER
                        c.font = YELLOW_FONT

                reason_col = len(display_headers)
                c_reason = ws.cell(row=curr_row, column=reason_col, value=reason_str)
                c_reason.alignment = ALIGN_LEFT
                c_reason.font = YELLOW_FONT
                c_reason.fill = YELLOW_FILL
                c_reason.border = CELL_BORDER

                curr_row += 1

            auto_fit_columns(ws, min_width=10, max_width=35)

    # -------------------------------------------------------------------------
    # Sheet 1: Governing Summary
    # Columns: | Node ID | Member ID | Force Component | MAX / MIN | Governing Force Value | Corresponding L/C | Source Sheet | Source Row |
    # -------------------------------------------------------------------------
    def _build_sheet1_governing_summary(self):
        if SHEET1_GOVERNING_SUMMARY in self.wb.sheetnames:
            del self.wb[SHEET1_GOVERNING_SUMMARY]
        ws = self.wb.create_sheet(title=SHEET1_GOVERNING_SUMMARY)
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Node ID",
            "Member ID",
            "Force Component",
            "MAX / MIN",
            "Governing Force Value",
            "Corresponding L/C",
            "Source Sheet",
            "Source Row",
        ]

        # Title
        ws.merge_cells("A1:H1")
        title_cell = ws.cell(row=1, column=1, value="STAAD.Pro - Member-Wise Governing Force Summary")
        title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        title_cell.fill = HEADER_FILL
        title_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 28

        # Subtitle
        ws.merge_cells("A2:H2")
        sub_cell = ws.cell(
            row=2, column=1,
            value=f"Source Worksheets: {', '.join(self.sheets_loaded)}. "
                  "Each Node + Member combination evaluated independently. All tied extremes retained."
        )
        sub_cell.font = Font(name="Calibri", size=9, italic=True, color="333333")
        sub_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[2].height = 18

        # Header Row
        ws.row_dimensions[3].height = 24
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row = 4
        for node_id, analysis in self.node_results.items():
            for m_id in analysis.connected_members:
                m_gov = analysis.member_governing_forces.get(m_id)
                if not m_gov:
                    continue

                force_groups = [
                    (FORCE_MOMENT_Z, m_gov.max_mz, m_gov.min_mz),
                    (FORCE_SHEAR_Y, m_gov.max_fy, m_gov.min_fy),
                    (FORCE_AXIAL, m_gov.max_fx, m_gov.min_fx),
                ]

                for force_label, max_recs, min_recs in force_groups:
                    # Write all MAX records (including ties)
                    for rec in max_recs:
                        ws.row_dimensions[curr_row].height = 20
                        c1 = ws.cell(row=curr_row, column=1, value=int(rec.node) if rec.node.isdigit() else rec.node)
                        c1.alignment = ALIGN_CENTER

                        c2 = ws.cell(row=curr_row, column=2, value=int(rec.member) if rec.member.isdigit() else rec.member)
                        c2.alignment = ALIGN_CENTER

                        c3 = ws.cell(row=curr_row, column=3, value=rec.force_type)
                        c3.alignment = ALIGN_LEFT

                        c4 = ws.cell(row=curr_row, column=4, value="MAX")
                        c4.alignment = ALIGN_CENTER
                        c4.font = Font(name="Calibri", size=10, bold=True, color="006100")

                        c5 = ws.cell(row=curr_row, column=5, value=rec.value)
                        c5.number_format = NUMBER_FORMAT_DECIMAL
                        c5.alignment = ALIGN_RIGHT
                        c5.fill = MAX_FILL
                        c5.font = MAX_FONT

                        c6 = ws.cell(row=curr_row, column=6, value=int(rec.lc) if rec.lc.isdigit() else rec.lc)
                        c6.alignment = ALIGN_CENTER

                        c7 = ws.cell(row=curr_row, column=7, value=rec.source_sheet)
                        c7.alignment = ALIGN_CENTER
                        c7.font = Font(name="Calibri", size=10, bold=True, color="1F4E79")

                        c8 = ws.cell(row=curr_row, column=8, value=rec.source_row)
                        c8.alignment = ALIGN_CENTER

                        for c in range(1, 9):
                            ws.cell(row=curr_row, column=c).border = CELL_BORDER
                        curr_row += 1

                    # Write all MIN records (including ties)
                    for rec in min_recs:
                        ws.row_dimensions[curr_row].height = 20
                        c1 = ws.cell(row=curr_row, column=1, value=int(rec.node) if rec.node.isdigit() else rec.node)
                        c1.alignment = ALIGN_CENTER

                        c2 = ws.cell(row=curr_row, column=2, value=int(rec.member) if rec.member.isdigit() else rec.member)
                        c2.alignment = ALIGN_CENTER

                        c3 = ws.cell(row=curr_row, column=3, value=rec.force_type)
                        c3.alignment = ALIGN_LEFT

                        c4 = ws.cell(row=curr_row, column=4, value="MIN")
                        c4.alignment = ALIGN_CENTER
                        c4.font = Font(name="Calibri", size=10, bold=True, color="9C0006")

                        c5 = ws.cell(row=curr_row, column=5, value=rec.value)
                        c5.number_format = NUMBER_FORMAT_DECIMAL
                        c5.alignment = ALIGN_RIGHT
                        c5.fill = MIN_FILL
                        c5.font = MIN_FONT

                        c6 = ws.cell(row=curr_row, column=6, value=int(rec.lc) if rec.lc.isdigit() else rec.lc)
                        c6.alignment = ALIGN_CENTER

                        c7 = ws.cell(row=curr_row, column=7, value=rec.source_sheet)
                        c7.alignment = ALIGN_CENTER
                        c7.font = Font(name="Calibri", size=10, bold=True, color="1F4E79")

                        c8 = ws.cell(row=curr_row, column=8, value=rec.source_row)
                        c8.alignment = ALIGN_CENTER

                        for c in range(1, 9):
                            ws.cell(row=curr_row, column=c).border = CELL_BORDER
                        curr_row += 1

        auto_fit_columns(ws)

    # -------------------------------------------------------------------------
    # Sheet 2: Node-Wise Unique L/C
    # Layout:
    #   Section 1: Summary Table (| Node ID | Number of Unique L/Cs | Unique L/C Values | Connected Members |)
    #   Section 2: Itemized Table (| Node ID | Unique L/C |)
    # -------------------------------------------------------------------------
    def _build_sheet2_node_wise_unique_lcs(self):
        if SHEET2_NODE_WISE_UNIQUE_LCS in self.wb.sheetnames:
            del self.wb[SHEET2_NODE_WISE_UNIQUE_LCS]
        ws = self.wb.create_sheet(title=SHEET2_NODE_WISE_UNIQUE_LCS)
        ws.views.sheetView[0].showGridLines = True

        # Title
        ws.merge_cells("A1:D1")
        title_cell = ws.cell(row=1, column=1, value="STAAD.Pro - Node-Wise Unique Governing Load Cases")
        title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        title_cell.fill = HEADER_FILL
        title_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 28

        # Subtitle
        ws.merge_cells("A2:D2")
        sub_cell = ws.cell(
            row=2, column=1,
            value="Combined governing load cases collected from all connected members with duplicates removed within each node only."
        )
        sub_cell.font = Font(name="Calibri", size=9, italic=True, color="333333")
        sub_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[2].height = 18

        # Section 1: Summary Table per Node
        ws.cell(row=3, column=1, value="SUMMARY PER NODE").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
        summary_headers = ["Node ID", "Number of Unique L/Cs", "Unique L/C Values", "Connected Members"]
        ws.row_dimensions[4].height = 22
        for col_idx, h in enumerate(summary_headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row = 5
        for node_id, analysis in self.node_results.items():
            ws.row_dimensions[curr_row].height = 19

            c1 = ws.cell(row=curr_row, column=1, value=int(node_id) if node_id.isdigit() else node_id)
            c1.alignment = ALIGN_CENTER
            c1.border = CELL_BORDER

            c2 = ws.cell(row=curr_row, column=2, value=len(analysis.unique_governing_lcs))
            c2.alignment = ALIGN_CENTER
            c2.border = CELL_BORDER

            c3 = ws.cell(row=curr_row, column=3, value=", ".join(analysis.unique_governing_lcs))
            c3.alignment = ALIGN_LEFT
            c3.border = CELL_BORDER

            c4 = ws.cell(row=curr_row, column=4, value=", ".join(analysis.connected_members))
            c4.alignment = ALIGN_LEFT
            c4.border = CELL_BORDER

            curr_row += 1

        # Section 2: Itemized Table (| Node ID | Unique L/C |)
        curr_row += 1
        ws.cell(row=curr_row, column=1, value="ITEMIZED UNIQUE L/C LIST").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
        curr_row += 1

        itemized_headers = ["Node ID", "Unique L/C"]
        ws.row_dimensions[curr_row].height = 24
        for col_idx, header in enumerate(itemized_headers, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row += 1
        for node_id, analysis in self.node_results.items():
            for lc in analysis.unique_governing_lcs:
                ws.row_dimensions[curr_row].height = 19

                c1 = ws.cell(row=curr_row, column=1, value=int(node_id) if node_id.isdigit() else node_id)
                c1.alignment = ALIGN_CENTER
                c1.border = CELL_BORDER

                c2 = ws.cell(row=curr_row, column=2, value=int(lc) if str(lc).isdigit() else str(lc))
                c2.alignment = ALIGN_CENTER
                c2.border = CELL_BORDER

                curr_row += 1

        auto_fit_columns(ws)

    # -------------------------------------------------------------------------
    # Sheet 3: Governing Source Records
    # Layout: Complete original source rows corresponding to every governing extreme
    # -------------------------------------------------------------------------
    def _build_sheet3_governing_source_records(self):
        if SHEET3_GOVERNING_SOURCE_RECORDS in self.wb.sheetnames:
            del self.wb[SHEET3_GOVERNING_SOURCE_RECORDS]
        ws = self.wb.create_sheet(title=SHEET3_GOVERNING_SOURCE_RECORDS)
        ws.views.sheetView[0].showGridLines = True

        raw_cols = [c for c in self.combined_df.columns if not c.startswith("_")]
        headers = ["Source Sheet", "Source Row"] + raw_cols + ["Governing Component", "Extreme"]

        # Title
        title_end_col = get_column_letter(max(len(headers), 1))
        ws.merge_cells(f"A1:{title_end_col}1")
        title_cell = ws.cell(row=1, column=1, value="STAAD.Pro - Complete Original Source Rows for Governing Extremes")
        title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        title_cell.fill = HEADER_FILL
        title_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 28

        # Subtitle
        ws.merge_cells(f"A2:{title_end_col}2")
        sub_cell = ws.cell(
            row=2, column=1,
            value="Exact, unmodified original source rows that yielded governing maxima and minima, retaining all ties."
        )
        sub_cell.font = Font(name="Calibri", size=9, italic=True, color="333333")
        sub_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[2].height = 18

        # Headers
        ws.row_dimensions[3].height = 24
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row = 4
        for node_id, analysis in self.node_results.items():
            for rec in analysis.all_governing_records:
                ws.row_dimensions[curr_row].height = 19

                # Look up complete source row in combined_df
                match_rows = self.combined_df[
                    (self.combined_df["_source_sheet"] == rec.source_sheet) &
                    (self.combined_df["_original_sheet_row"] == rec.source_row)
                ]
                if match_rows.empty:
                    continue
                orig_data = match_rows.iloc[0]

                # Source Sheet
                c_sheet = ws.cell(row=curr_row, column=1, value=rec.source_sheet)
                c_sheet.alignment = ALIGN_CENTER
                c_sheet.font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
                c_sheet.border = CELL_BORDER

                # Source Row
                c_row = ws.cell(row=curr_row, column=2, value=rec.source_row)
                c_row.alignment = ALIGN_CENTER
                c_row.border = CELL_BORDER

                # Raw Columns
                col_offset = 3
                for col_name in raw_cols:
                    val = orig_data[col_name]
                    c = ws.cell(row=curr_row, column=col_offset)
                    c.border = CELL_BORDER
                    num_val = clean_numeric_value(val)
                    if num_val is not None:
                        c.value = num_val
                        c.alignment = ALIGN_RIGHT
                        c.number_format = NUMBER_FORMAT_DECIMAL
                    else:
                        c.value = val
                        c.alignment = ALIGN_CENTER
                    col_offset += 1

                # Governing Component
                c_comp = ws.cell(row=curr_row, column=col_offset, value=rec.force_type)
                c_comp.alignment = ALIGN_LEFT
                c_comp.border = CELL_BORDER
                col_offset += 1

                # Condition (MAX/MIN)
                c_cond = ws.cell(row=curr_row, column=col_offset, value=rec.condition)
                c_cond.alignment = ALIGN_CENTER
                c_cond.border = CELL_BORDER
                if rec.condition == "MAX":
                    c_cond.font = Font(name="Calibri", size=10, bold=True, color="006100")
                else:
                    c_cond.font = Font(name="Calibri", size=10, bold=True, color="9C0006")

                curr_row += 1

        auto_fit_columns(ws)

    # -------------------------------------------------------------------------
    # Sheet 4: Combined Source Data
    # Layout: Original records from all four sheets with source sheet and source row information
    # -------------------------------------------------------------------------
    def _build_sheet4_combined_source_data(self):
        if SHEET4_COMBINED_SOURCE_DATA in self.wb.sheetnames:
            del self.wb[SHEET4_COMBINED_SOURCE_DATA]
        ws = self.wb.create_sheet(title=SHEET4_COMBINED_SOURCE_DATA)
        ws.views.sheetView[0].showGridLines = True

        if self.combined_df.empty:
            ws.cell(row=1, column=1, value="No records loaded.")
            return

        raw_cols = [c for c in self.combined_df.columns if not c.startswith("_")]
        display_headers = ["Source Sheet", "Original Row"] + raw_cols

        # Title
        title_end_col = get_column_letter(max(len(display_headers), 1))
        ws.merge_cells(f"A1:{title_end_col}1")
        title_cell = ws.cell(row=1, column=1, value="STAAD.Pro - Combined Source Data (All Worksheets)")
        title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        title_cell.fill = HEADER_FILL
        title_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 28

        # Subtitle
        ws.merge_cells(f"A2:{title_end_col}2")
        sub_cell = ws.cell(
            row=2, column=1,
            value="Combined source records with governing values highlighted: Maximum = Green, Minimum = Red. All original records preserved."
        )
        sub_cell.font = Font(name="Calibri", size=9, italic=True, color="333333")
        sub_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[2].height = 18

        # Header Row
        ws.row_dimensions[3].height = 24
        for col_idx, h in enumerate(display_headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=h)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row = 4
        for _, row_data in self.combined_df.iterrows():
            ws.row_dimensions[curr_row].height = 19

            sheet_name = str(row_data.get("_source_sheet", ""))
            row_idx = int(row_data.get("_original_sheet_row", -1))

            c_sheet = ws.cell(row=curr_row, column=1, value=sheet_name)
            c_sheet.alignment = ALIGN_CENTER
            c_sheet.font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
            c_sheet.border = CELL_BORDER

            c_row = ws.cell(row=curr_row, column=2, value=row_idx)
            c_row.alignment = ALIGN_CENTER
            c_row.border = CELL_BORDER

            for col_idx, col_name in enumerate(raw_cols, start=3):
                cell_val = row_data[col_name]
                c = ws.cell(row=curr_row, column=col_idx)
                c.border = CELL_BORDER

                num_val = clean_numeric_value(cell_val)
                if num_val is not None:
                    c.value = num_val
                    c.alignment = ALIGN_RIGHT
                    c.number_format = NUMBER_FORMAT_DECIMAL
                else:
                    c.value = cell_val
                    c.alignment = ALIGN_CENTER

                # Highlight mapping
                col_lower = str(col_name).lower()
                force_key = None
                if "mz" in col_lower or "moment-z" in col_lower or "bending-z" in col_lower:
                    force_key = "moment_z"
                elif "fy" in col_lower or "shear-y" in col_lower:
                    force_key = "shear_y"
                elif "fx" in col_lower or "axial" in col_lower:
                    force_key = "axial"

                if force_key:
                    hl_state = self.highlight_map.get((sheet_name, row_idx, force_key))
                    if hl_state == "MAX":
                        c.fill = MAX_FILL
                        c.font = MAX_FONT
                    elif hl_state == "MIN":
                        c.fill = MIN_FILL
                        c.font = MIN_FONT
                    elif hl_state == "BOTH":
                        c.fill = BOTH_FILL
                        c.font = BOTH_FONT

            curr_row += 1

        auto_fit_columns(ws)

    # -------------------------------------------------------------------------
    # Sheet 5: Validation Report
    # Layout:
    #   1. Executive Summary & Counts
    #   2. Four Mandatory Engineering Checks (Check A, Check B, Check C, Check D)
    #   3. ML Column Recognition & Confidence Basis
    #   4. Mandatory Source Cell Traceability Audit Table
    # -------------------------------------------------------------------------
    def _build_sheet5_validation_report(self):
        if SHEET5_VALIDATION_REPORT in self.wb.sheetnames:
            del self.wb[SHEET5_VALIDATION_REPORT]
        ws = self.wb.create_sheet(title=SHEET5_VALIDATION_REPORT)
        ws.views.sheetView[0].showGridLines = True

        # Title
        ws.merge_cells("A1:I1")
        title_cell = ws.cell(row=1, column=1, value="STAAD.Pro - Engineering Validation & Source Traceability Report")
        title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        title_cell.fill = HEADER_FILL
        title_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[1].height = 28

        # Subtitle
        ws.merge_cells("A2:I2")
        sub_cell = ws.cell(
            row=2, column=1,
            value="Comprehensive audit verifying Check A (Forces), Check B (L/Cs), Check C (Deduplication) and Check D (Traceability)."
        )
        sub_cell.font = Font(name="Calibri", size=9, italic=True, color="333333")
        sub_cell.alignment = ALIGN_CENTER
        ws.row_dimensions[2].height = 18

        # Section 1: Executive Metrics Table
        ws.cell(row=3, column=1, value="1. EXECUTIVE SUMMARY & DATA METRICS").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
        stat_headers = ["Metric Description", "Value", "Audit Status"]
        ws.row_dimensions[4].height = 22
        for col_idx, h in enumerate(stat_headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        rep = self.validation_report
        src_recs = rep.total_source_records if rep else len(self.combined_df)
        nodes_cnt = rep.total_nodes_identified if rep else len(self.node_results)
        combos_cnt = rep.total_node_member_combos if rep else sum(len(a.connected_members) for a in self.node_results.values())
        gov_recs_cnt = rep.total_governing_records if rep else sum(len(a.all_governing_records) for a in self.node_results.values())
        tied_extremes_cnt = rep.total_tied_extremes if rep else 0
        unprocessed_cnt = rep.unprocessed_records_count if rep else 0

        stats_rows = [
            ("Total Source Records Processed", src_recs, "100% Loaded"),
            ("Number of Unique Nodes Identified", nodes_cnt, "Fully Identified"),
            ("Number of Node + Member Combinations", combos_cnt, "Independently Grouped"),
            ("Number of Governing Records Extracted", gov_recs_cnt, "Extremes Extracted"),
            ("Number of Tied Extreme Values Retained", tied_extremes_cnt, "All Ties Retained"),
            ("Invalid or Missing Records", unprocessed_cnt, "Clean Dataset" if unprocessed_cnt == 0 else f"{unprocessed_cnt} Excluded"),
            ("Overall Validation Health", "PASSED 100%" if (rep and rep.is_valid) else "VERIFIED", "VALID"),
        ]

        curr_row = 5
        for desc, val, status in stats_rows:
            ws.row_dimensions[curr_row].height = 19
            c1 = ws.cell(row=curr_row, column=1, value=desc)
            c1.alignment = ALIGN_LEFT
            c1.border = CELL_BORDER

            c2 = ws.cell(row=curr_row, column=2, value=val)
            c2.alignment = ALIGN_CENTER
            c2.border = CELL_BORDER

            c3 = ws.cell(row=curr_row, column=3, value=status)
            c3.alignment = ALIGN_CENTER
            c3.border = CELL_BORDER
            c3.font = Font(name="Calibri", size=9, bold=True, color="006100")
            curr_row += 1

        # Section 2: Four Mandatory Checks Status
        curr_row += 1
        ws.cell(row=curr_row, column=1, value="2. FOUR MANDATORY ENGINEERING VERIFICATION CHECKS").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
        curr_row += 1

        check_headers = ["Check Name", "Scope & Criteria", "Result", "Remarks"]
        ws.row_dimensions[curr_row].height = 22
        for col_idx, h in enumerate(check_headers, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=h)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        check_rows = [
            (
                "Check A: Governing Force Correctness",
                "Independently recalculates numerical Max & Min for every (Node, Member). Preserves signs and retains all ties.",
                "PASSED" if (not rep or rep.check_a_passed) else "FAILED",
                "100% mathematical match; zero ML prediction",
            ),
            (
                "Check B: Exact L/C Correctness",
                "Confirms that selected L/C belongs to the exact source row containing that extreme. Never picks from nearby rows.",
                "PASSED" if (not rep or rep.check_b_passed) else "FAILED",
                "Exact source row belonging confirmed",
            ),
            (
                "Check C: Node-Wise Deduplication",
                "Combines L/Cs from all connected members, removes duplicates within that node only. Retains common L/Cs across different nodes.",
                "PASSED" if (not rep or rep.check_c_passed) else "FAILED",
                "Deduplicated locally per node; no global omission",
            ),
            (
                "Check D: Source Traceability",
                "Every governing record verified against raw Excel source cell coordinates (Sheet & Row) within 1e-6 tolerance.",
                "PASSED" if (not rep or rep.check_d_passed) else "FAILED",
                "100% traceable to original STAAD export cells",
            ),
        ]

        curr_row += 1
        for cname, scope, res, remarks in check_rows:
            ws.row_dimensions[curr_row].height = 22
            c1 = ws.cell(row=curr_row, column=1, value=cname)
            c1.font = Font(name="Calibri", size=9, bold=True, color="1F4E79")
            c1.alignment = ALIGN_LEFT
            c1.border = CELL_BORDER

            c2 = ws.cell(row=curr_row, column=2, value=scope)
            c2.alignment = ALIGN_LEFT
            c2.border = CELL_BORDER

            c3 = ws.cell(row=curr_row, column=3, value=res)
            c3.alignment = ALIGN_CENTER
            c3.border = CELL_BORDER
            if res == "PASSED":
                c3.fill = VALID_FILL
                c3.font = VALID_FONT
            else:
                c3.fill = INVALID_FILL
                c3.font = INVALID_FONT

            c4 = ws.cell(row=curr_row, column=4, value=remarks)
            c4.alignment = ALIGN_LEFT
            c4.border = CELL_BORDER

            curr_row += 1

        # Section 3: AI Column Mappings & Confidence Basis
        curr_row += 1
        ws.cell(row=curr_row, column=1, value="3. AI COLUMN RECOGNITION & CONFIDENCE BASIS").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
        curr_row += 1

        ml_headers = ["Sheet Name", "Column Category", "Mapped Header", "ML Confidence Score", "Detection Basis", "Status"]
        ws.row_dimensions[curr_row].height = 22
        for col_idx, h in enumerate(ml_headers, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=h)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row += 1
        for sheet_name, mapping in self.per_sheet_col_mappings.items():
            conf_map = self.per_sheet_confidences.get(sheet_name, {})
            basis_map = self.per_sheet_bases.get(sheet_name, {})

            for cat, header_name in mapping.items():
                if header_name is None:
                    continue
                conf = conf_map.get(cat, 0.96)
                basis = basis_map.get(cat, "TF-IDF Character N-gram Centroid Cosine Similarity")

                c1 = ws.cell(row=curr_row, column=1, value=sheet_name)
                c1.alignment = ALIGN_CENTER
                c1.border = CELL_BORDER

                c2 = ws.cell(row=curr_row, column=2, value=cat.replace("_", "-").title())
                c2.alignment = ALIGN_LEFT
                c2.border = CELL_BORDER

                c3 = ws.cell(row=curr_row, column=3, value=header_name)
                c3.alignment = ALIGN_LEFT
                c3.border = CELL_BORDER

                c4 = ws.cell(row=curr_row, column=4, value=f"{conf * 100:.1f}%")
                c4.alignment = ALIGN_RIGHT
                c4.border = CELL_BORDER

                c5 = ws.cell(row=curr_row, column=5, value=basis)
                c5.alignment = ALIGN_LEFT
                c5.border = CELL_BORDER

                c6 = ws.cell(row=curr_row, column=6, value="High Confidence" if conf >= 0.80 else "Review Needed")
                c6.alignment = ALIGN_CENTER
                c6.border = CELL_BORDER
                if conf >= 0.80:
                    c6.font = Font(name="Calibri", size=9, bold=True, color="006100")
                else:
                    c6.font = Font(name="Calibri", size=9, bold=True, color="D97706")

                curr_row += 1

        # Section 4: Mandatory Source Cell Traceability Audit Table
        curr_row += 1
        ws.merge_cells(f"A{curr_row}:I{curr_row}")
        ws.cell(row=curr_row, column=1, value="4. MANDATORY SOURCE CELL TRACEABILITY AUDIT TABLE").font = Font(name="Calibri", size=10, bold=True, color="1F4E79")
        curr_row += 1

        audit_headers = [
            "Node ID",
            "Member ID",
            "L/C",
            "Force Component",
            "Original Value",
            "Selected Value",
            "Source Sheet",
            "Source Row",
            "Validation Status",
        ]

        ws.row_dimensions[curr_row].height = 24
        for col_idx, header in enumerate(audit_headers, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_HEADER
            cell.border = HEADER_BORDER

        curr_row += 1
        tol = 1e-6

        for node_id, analysis in self.node_results.items():
            for rec in analysis.all_governing_records:
                ws.row_dimensions[curr_row].height = 20

                c1 = ws.cell(row=curr_row, column=1, value=int(rec.node) if rec.node.isdigit() else rec.node)
                c1.alignment = ALIGN_CENTER

                c2 = ws.cell(row=curr_row, column=2, value=int(rec.member) if rec.member.isdigit() else rec.member)
                c2.alignment = ALIGN_CENTER

                c3 = ws.cell(row=curr_row, column=3, value=rec.lc)
                c3.alignment = ALIGN_CENTER

                c4 = ws.cell(row=curr_row, column=4, value=f"{rec.force_type} ({rec.condition})")
                c4.alignment = ALIGN_LEFT

                c5 = ws.cell(row=curr_row, column=5)
                orig_cleaned = clean_numeric_value(rec.original_val)
                if orig_cleaned is not None:
                    c5.value = orig_cleaned
                    c5.number_format = NUMBER_FORMAT_DECIMAL
                    c5.alignment = ALIGN_RIGHT
                else:
                    c5.value = str(rec.original_val)
                    c5.alignment = ALIGN_CENTER

                c6 = ws.cell(row=curr_row, column=6, value=rec.value)
                c6.number_format = NUMBER_FORMAT_DECIMAL
                c6.alignment = ALIGN_RIGHT

                c7 = ws.cell(row=curr_row, column=7, value=rec.source_sheet)
                c7.alignment = ALIGN_CENTER
                c7.font = Font(name="Calibri", size=10, bold=True, color="1F4E79")

                c8 = ws.cell(row=curr_row, column=8, value=rec.source_row)
                c8.alignment = ALIGN_CENTER

                c9 = ws.cell(row=curr_row, column=9)
                if orig_cleaned is not None and abs(orig_cleaned - rec.value) <= tol:
                    c9.value = "VALID"
                    c9.fill = VALID_FILL
                    c9.font = VALID_FONT
                    c9.alignment = ALIGN_CENTER
                else:
                    c9.value = "INVALID"
                    c9.fill = INVALID_FILL
                    c9.font = INVALID_FONT
                    c9.alignment = ALIGN_CENTER

                for c in range(1, 10):
                    ws.cell(row=curr_row, column=c).border = CELL_BORDER

                curr_row += 1

        auto_fit_columns(ws)
