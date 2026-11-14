"""
Robust ML-Assisted Excel Reader for STAAD.Pro Member Force Worksheets (R-L, R-R, JB-F, JB-B).
Features Scikit-learn TF-IDF semantic header classification, confidence scoring,
dynamic multi-level header merging, and exact source-cell row tracking.
"""

import re
from typing import Dict, List, Optional, Tuple, Any, Set
import pandas as pd
import openpyxl

from .config import (
    DEFAULT_TARGET_SHEETS,
    COLUMN_ALIASES,
    FORCE_MOMENT_Z,
    FORCE_SHEAR_Y,
    FORCE_AXIAL,
)
from .ml_engine import (
    MLColumnClassifier,
    MLSheetClassifier,
    CAT_NODE,
    CAT_MEMBER,
    CAT_LC,
    CAT_MOMENT_Z,
    CAT_SHEAR_Y,
    CAT_AXIAL,
    CAT_IGNORE,
)

# Global default constant to prevent any unbound scope errors
DEFAULT_MAX_HEADER_SCAN_ROWS: int = 30


def normalize_token(val: Any) -> str:
    """Normalize a string for robust matching."""
    if val is None:
        return ""
    s = str(val).strip().lower()
    return re.sub(r"\s+", " ", s)


def match_sheet_name(actual_sheet: str, target: str) -> bool:
    """Case-insensitive, whitespace-trimmed sheet name matcher."""
    a = normalize_token(actual_sheet)
    t = normalize_token(target)
    if a == t:
        return True
    a_clean = re.sub(r"[\s\-_]+", "", a)
    t_clean = re.sub(r"[\s\-_]+", "", t)
    return a_clean == t_clean


def clean_numeric_value(val: Any) -> Optional[float]:
    """
    Parse numeric cell value safely:
    Handles commas ('1,526.629'), Unicode minus signs ('−'),
    text whitespace, scientific notation, and empty markers ('-', 'N/A').
    """
    if val is None or pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)

    s = str(val).strip()
    if s == "" or s == "-" or s.lower() in ("nan", "null", "none", "n/a", "#n/a"):
        return None

    # Replace Unicode minus / en-dash / em-dash with ASCII minus
    s = s.replace("\u2212", "-").replace("\u2013", "-").replace("\u2014", "-")
    # Remove commas
    s = s.replace(",", "")

    try:
        return float(s)
    except ValueError:
        return None


def clean_id_value(val: Any) -> str:
    """Normalize node, member, and L/C identifiers without altering notation."""
    if val is None or pd.isna(val):
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    s = str(val).strip()
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    return s


def find_header_row_ml(
    sheet_data: List[List[Any]],
    classifier: Optional[MLColumnClassifier] = None,
    max_rows: int = DEFAULT_MAX_HEADER_SCAN_ROWS,
) -> Tuple[int, Dict[str, int], Dict[str, float], int]:
    """
    Identifies the header row using ML semantic classification, supporting both
    single-line and two-line merged headers.
    Returns: (header_row_index, mapping_dict, confidence_dict, header_height)
    """
    if classifier is None:
        classifier = MLColumnClassifier()

    best_row_idx = -1
    best_score = -1.0
    best_mapping: Dict[str, int] = {}
    best_confidences: Dict[str, float] = {}
    best_height = 1

    num_rows = min(max_rows, len(sheet_data))

    for r in range(num_rows):
        row1 = sheet_data[r]
        row2 = sheet_data[r + 1] if (r + 1 < len(sheet_data)) else None

        # 1. Single row r evaluation
        score1, map1, conf1 = _evaluate_row_ml(row1, classifier)
        if score1 > best_score:
            best_score = score1
            best_row_idx = r
            best_mapping = map1
            best_confidences = conf1
            best_height = 1

        # 2. Merged row (row r + row r+1) evaluation
        if row2 is not None:
            merged_row = []
            for c1, c2 in zip(row1, row2):
                p1 = str(c1).strip() if c1 is not None else ""
                p2 = str(c2).strip() if c2 is not None else ""
                combined = f"{p1} {p2}".strip()
                merged_row.append(combined)

            score2, map2, conf2 = _evaluate_row_ml(merged_row, classifier)
            if score2 > best_score:
                best_score = score2
                best_row_idx = r
                best_mapping = map2
                best_confidences = conf2
                best_height = 2

    # A valid header must have identified both 'node' and 'member' with high confidence
    if best_score >= 4.0 and "node" in best_mapping and "member" in best_mapping:
        return best_row_idx, best_mapping, best_confidences, best_height

    return -1, {}, {}, 1


def _evaluate_row_ml(
    row: List[Any],
    classifier: MLColumnClassifier,
) -> Tuple[float, Dict[str, int], Dict[str, float]]:
    """Evaluates a candidate header row using ML classifier predictions and confidence scores."""
    mapping: Dict[str, int] = {}
    confidences: Dict[str, float] = {}
    total_score = 0.0

    for col_idx, cell_val in enumerate(row):
        if cell_val is None or str(cell_val).strip() == "":
            continue

        cat, conf = classifier.predict_column(cell_val)

        if cat != CAT_IGNORE and conf >= 0.50:
            # If category not yet mapped or higher confidence than previous candidate
            if cat not in mapping or conf > confidences.get(cat, 0.0):
                mapping[cat] = col_idx
                confidences[cat] = conf

    # Weigh essential columns
    if "node" in mapping:
        total_score += 2.0 * confidences["node"]
    if "member" in mapping:
        total_score += 2.0 * confidences["member"]
    if "lc" in mapping:
        total_score += 1.5 * confidences["lc"]
    if "moment_z" in mapping:
        total_score += 1.5 * confidences["moment_z"]
    if "shear_y" in mapping:
        total_score += 1.5 * confidences["shear_y"]
    if "axial" in mapping:
        total_score += 1.5 * confidences["axial"]

    return total_score, mapping, confidences


class StaadReader:
    """
    ML-assisted reader for STAAD.Pro member force workbooks.
    Classifies sheets, maps columns via character n-gram TF-IDF,
    assigns confidence scores, and tracks exact cell coordinates for auditability.
    """

    def __init__(
        self,
        file_path: str,
        target_sheets: Optional[List[str]] = None,
        classifier: Optional[MLColumnClassifier] = None,
    ):
        self.file_path = file_path
        self.target_sheet_names = target_sheets or DEFAULT_TARGET_SHEETS
        self.wb = openpyxl.load_workbook(file_path, data_only=True)
        self.available_sheets = self.wb.sheetnames

        self.classifier = classifier or MLColumnClassifier()
        self.sheet_classifier = MLSheetClassifier()

        self.sheets_found: List[str] = []
        self.sheets_missing: List[str] = []
        self.per_sheet_dfs: Dict[str, pd.DataFrame] = {}
        self.per_sheet_col_mappings: Dict[str, Dict[str, str]] = {}
        self.per_sheet_confidences: Dict[str, Dict[str, float]] = {}
        self.per_sheet_bases: Dict[str, Dict[str, str]] = {}
        self.invalid_force_records: List[Dict[str, Any]] = []

        self.combined_df: Optional[pd.DataFrame] = None

    def load_worksheets(
        self,
        max_scan_rows: int = DEFAULT_MAX_HEADER_SCAN_ROWS,
    ) -> Tuple[pd.DataFrame, Dict[str, pd.DataFrame], List[Dict[str, Any]]]:
        """
        Loads worksheets using ML classification, parses numeric forces,
        and generates a combined DataFrame with exact original Excel row tracking.
        """
        matched_sheets: Dict[str, str] = {}
        for target in self.target_sheet_names:
            found = None
            for actual in self.available_sheets:
                if match_sheet_name(actual, target):
                    found = actual
                    break
            if found:
                matched_sheets[target] = found
            else:
                self.sheets_missing.append(target)

        sheets_to_process = list(matched_sheets.values())
        if not sheets_to_process:
            for s in self.available_sheets:
                ws = self.wb[s]
                sheet_data = list(ws.iter_rows(values_only=True))
                h_idx, col_map, _, _ = find_header_row_ml(
                    sheet_data, classifier=self.classifier, max_rows=max_scan_rows
                )
                if h_idx != -1 and "node" in col_map and "member" in col_map:
                    sheets_to_process.append(s)

        if not sheets_to_process:
            raise ValueError(
                f"No valid member force worksheets found in '{self.file_path}'. "
                f"Expected worksheets: {self.target_sheet_names}. "
                f"Available worksheets: {self.available_sheets}."
            )

        self.sheets_found = sheets_to_process
        all_dfs: List[pd.DataFrame] = []

        for sheet_name in sheets_to_process:
            df, col_mapping, confidences, invalids = self._parse_single_sheet(
                sheet_name, max_scan_rows=max_scan_rows
            )
            self.per_sheet_dfs[sheet_name] = df
            self.per_sheet_col_mappings[sheet_name] = col_mapping
            self.per_sheet_confidences[sheet_name] = confidences
            bases = {}
            for cat, hdr in col_mapping.items():
                if hdr:
                    _, _, b = self.classifier.predict_column_with_basis(hdr)
                    bases[cat] = b
            self.per_sheet_bases[sheet_name] = bases
            self.invalid_force_records.extend(invalids)
            all_dfs.append(df)

        self.combined_df = pd.concat(all_dfs, ignore_index=True)
        return self.combined_df, self.per_sheet_dfs, self.invalid_force_records

    def _parse_single_sheet(
        self,
        sheet_name: str,
        max_scan_rows: int = DEFAULT_MAX_HEADER_SCAN_ROWS,
    ) -> Tuple[pd.DataFrame, Dict[str, str], Dict[str, float], List[Dict[str, Any]]]:
        """Parses a single member force worksheet with ML column mapping."""
        ws = self.wb[sheet_name]
        sheet_data = list(ws.iter_rows(values_only=True))
        if not sheet_data:
            return pd.DataFrame(), {}, {}, []

        header_idx, col_map, confidences, header_height = find_header_row_ml(
            sheet_data, classifier=self.classifier, max_rows=max_scan_rows
        )
        if header_idx == -1 or "node" not in col_map or "member" not in col_map:
            raise ValueError(
                f"Worksheet '{sheet_name}' in '{self.file_path}' does not contain recognized "
                "Node and Member/Beam header columns."
            )

        # Construct raw headers
        if header_height == 1:
            header_row = sheet_data[header_idx]
            raw_headers = [str(c).strip() if c is not None else f"Col_{i}" for i, c in enumerate(header_row)]
        else:
            row1 = sheet_data[header_idx]
            row2 = sheet_data[header_idx + 1]
            raw_headers = []
            for i, (c1, c2) in enumerate(zip(row1, row2)):
                p1 = str(c1).strip() if c1 is not None else ""
                p2 = str(c2).strip() if c2 is not None else ""
                combined = f"{p1} {p2}".strip()
                raw_headers.append(combined if combined else f"Col_{i}")

        col_mapping = {
            "member": raw_headers[col_map["member"]],
            "node": raw_headers[col_map["node"]],
            "lc": raw_headers[col_map["lc"]] if "lc" in col_map else None,
            "axial": raw_headers[col_map["axial"]] if "axial" in col_map else None,
            "shear_y": raw_headers[col_map["shear_y"]] if "shear_y" in col_map else None,
            "moment_z": raw_headers[col_map["moment_z"]] if "moment_z" in col_map else None,
        }

        for req in ["member", "node", "lc"]:
            if col_mapping.get(req) is None:
                raise ValueError(f"Worksheet '{sheet_name}' is missing essential column: '{req}'.")

        data_start_idx = header_idx + header_height
        data_rows = sheet_data[data_start_idx:]
        df = pd.DataFrame(data_rows, columns=raw_headers)

        node_col = col_mapping["node"]
        member_col = col_mapping["member"]
        lc_col = col_mapping["lc"]

        # Drop rows where Node or Member is empty
        df = df[df[node_col].notna() & df[member_col].notna()].copy()
        df = df[df[node_col].astype(str).str.strip() != ""].copy()
        df = df[df[member_col].astype(str).str.strip() != ""].copy()

        df["_source_sheet"] = sheet_name
        df["_original_sheet_row"] = [data_start_idx + 1 + i for i in range(len(df))]

        df["_clean_node"] = df[node_col].apply(clean_id_value)
        df["_clean_member"] = df[member_col].apply(clean_id_value)
        df["_clean_lc"] = df[lc_col].apply(clean_id_value)

        invalids = []
        for force_key in ["moment_z", "shear_y", "axial"]:
            actual_col = col_mapping.get(force_key)
            if actual_col and actual_col in df.columns:
                clean_col = f"_clean_{force_key}"
                raw_store_col = f"_raw_val_{force_key}"

                df[raw_store_col] = df[actual_col]
                df[clean_col] = df[actual_col].apply(clean_numeric_value)

                null_mask = df[actual_col].isna() | (df[clean_col].isna() & df[actual_col].notna())
                if null_mask.any():
                    for _, bad_row in df[null_mask].iterrows():
                        invalids.append({
                            "source_sheet": sheet_name,
                            "excel_row": bad_row["_original_sheet_row"],
                            "node": bad_row["_clean_node"],
                            "member": bad_row["_clean_member"],
                            "lc": bad_row["_clean_lc"],
                            "force_type": force_key,
                            "raw_value": bad_row[actual_col],
                        })

        return df, col_mapping, confidences, invalids
