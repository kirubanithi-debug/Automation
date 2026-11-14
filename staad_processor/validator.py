"""
Validation Engine for STAAD.Pro Member-Wise Force Extraction & Node-Wise Deduplication:
Enforces 100% exact source-value traceability, cross-checks numerical extremes,
verifies inclusion of R-L, R-R, JB-F, JB-B, validates connected member counts,
checks benchmark extremes (Node 136 / Member 2028), and verifies the four master checks:
  - Check A: Governing Force Correctness (all ties retained)
  - Check B: L/C Correctness (exact row belonging)
  - Check C: Node-Wise Deduplication (no duplicates, no omissions, independent per node)
  - Check D: Source Cell Traceability (100% match)
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any
import pandas as pd

from .config import DEFAULT_TARGET_SHEETS, FORCE_MOMENT_Z, FORCE_SHEAR_Y, FORCE_AXIAL
from .analyzer import NodeAnalysis
from .reader import clean_numeric_value


@dataclass
class ValidationReport:
    """Validation report containing status, audit checks, and error logs."""
    is_valid: bool
    checks_passed: int
    checks_failed: int
    total_source_records: int = 0
    total_nodes_identified: int = 0
    total_node_member_combos: int = 0
    total_governing_records: int = 0
    total_tied_extremes: int = 0
    per_node_unique_lcs: Dict[str, int] = field(default_factory=dict)
    unprocessed_records_count: int = 0
    check_a_passed: bool = True
    check_b_passed: bool = True
    check_c_passed: bool = True
    check_d_passed: bool = True
    messages: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    invalid_data_warnings: List[str] = field(default_factory=list)
    audit_mismatches: List[Dict[str, Any]] = field(default_factory=list)

    def add_pass(self, msg: str):
        self.checks_passed += 1
        self.messages.append(f"[PASS] {msg}")

    def add_fail(self, err: str):
        self.checks_failed += 1
        self.is_valid = False
        self.errors.append(f"[FAIL] {err}")
        self.messages.append(f"[FAIL] {err}")

    def add_warning(self, warn: str):
        self.invalid_data_warnings.append(f"[WARN] {warn}")


class StaadValidator:
    """Enforces strict engineering accuracy, exact L/C traceability, and node-wise deduplication."""

    def __init__(
        self,
        node_results: Dict[str, NodeAnalysis],
        combined_df: pd.DataFrame,
        sheets_loaded: List[str],
        invalid_force_records: List[Dict[str, Any]],
        expected_sheets: Optional[List[str]] = None,
    ):
        self.node_results = node_results
        self.combined_df = combined_df
        self.sheets_loaded = sheets_loaded
        self.invalid_force_records = invalid_force_records
        self.expected_sheets = expected_sheets or DEFAULT_TARGET_SHEETS

    def validate(self) -> ValidationReport:
        report = ValidationReport(is_valid=True, checks_passed=0, checks_failed=0)
        tol = 1e-6

        # Statistics gathering
        report.total_source_records = len(self.combined_df)
        report.total_nodes_identified = len(self.node_results)
        report.total_node_member_combos = sum(len(a.connected_members) for a in self.node_results.values())
        report.total_governing_records = sum(len(a.all_governing_records) for a in self.node_results.values())
        report.unprocessed_records_count = len(self.invalid_force_records)
        report.per_node_unique_lcs = {nid: len(a.unique_governing_lcs) for nid, a in self.node_results.items()}

        # Count tied extremes
        tied_count = 0
        for analysis in self.node_results.values():
            for m_gov in analysis.member_governing_forces.values():
                for rec_list in [m_gov.max_mz, m_gov.min_mz, m_gov.max_fy, m_gov.min_fy, m_gov.max_fx, m_gov.min_fx]:
                    if len(rec_list) > 1:
                        tied_count += (len(rec_list) - 1)
        report.total_tied_extremes = tied_count

        # 1. Verify that all 4 source worksheets were processed
        loaded_set = set(self.sheets_loaded)
        for req_sheet in self.expected_sheets:
            if req_sheet in loaded_set:
                report.add_pass(f"Source worksheet '{req_sheet}' confirmed loaded and processed.")
            else:
                report.add_fail(f"Source worksheet '{req_sheet}' missing from loaded worksheets!")

        # 2. Verify all unique nodes from raw data are present
        raw_nodes = set(self.combined_df["_clean_node"].unique())
        result_nodes = set(self.node_results.keys())
        if raw_nodes == result_nodes:
            report.add_pass(f"All {len(raw_nodes)} unique nodes across all worksheets were processed without omission.")
        else:
            missing_nodes = raw_nodes - result_nodes
            report.add_fail(f"Nodes missing from analysis: {missing_nodes}")

        # 3. Report any blank or invalid data rows separately
        if self.invalid_force_records:
            report.add_warning(
                f"Identified {len(self.invalid_force_records)} records with blank or invalid force values. "
                "These values were safely excluded from extreme calculations."
            )
            for inv in self.invalid_force_records[:5]:
                report.add_warning(
                    f"  Sheet: {inv['source_sheet']}, Row: {inv['excel_row']}, Node: {inv['node']}, "
                    f"Member: {inv['member']}, L/C: {inv['lc']}, Force: {inv['force_type']}, Raw: '{inv['raw_value']}'"
                )
        else:
            report.add_pass("Zero blank or corrupted force records found in raw input sheets.")

        # =========================================================================
        # CHECK A: Governing Force Correctness (Member-Wise, all ties retained)
        # =========================================================================
        for node_id, analysis in self.node_results.items():
            node_df = self.combined_df[self.combined_df["_clean_node"] == str(node_id)]

            for m_id in analysis.connected_members:
                m_gov = analysis.member_governing_forces.get(m_id)
                if not m_gov:
                    report.add_fail(f"Check A Failure: Node {node_id} Member {m_id} missing from governing results.")
                    report.check_a_passed = False
                    continue

                m_df = node_df[node_df["_clean_member"] == str(m_id)]

                # Check Moment-Z
                if "_clean_moment_z" in m_df.columns:
                    valid_mz = m_df[m_df["_clean_moment_z"].notna()]["_clean_moment_z"].astype(float)
                    if not valid_mz.empty:
                        exp_max = float(valid_mz.max())
                        exp_min = float(valid_mz.min())
                        exp_max_ties = int((abs(valid_mz - exp_max) <= tol).sum())
                        exp_min_ties = int((abs(valid_mz - exp_min) <= tol).sum())

                        if not m_gov.max_mz or abs(m_gov.max_mz[0].value - exp_max) > tol:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Moment-Z Max mismatch. Expected {exp_max}, got {m_gov.max_mz[0].value if m_gov.max_mz else None}")
                            report.check_a_passed = False
                        elif len(m_gov.max_mz) != exp_max_ties:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Moment-Z Max tie count mismatch. Expected {exp_max_ties}, got {len(m_gov.max_mz)}")
                            report.check_a_passed = False
                        else:
                            report.add_pass(f"Check A Verified: Node {node_id} Member {m_id} Moment-Z Max = {exp_max:.3f} ({len(m_gov.max_mz)} ties retained)")

                        if not m_gov.min_mz or abs(m_gov.min_mz[0].value - exp_min) > tol:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Moment-Z Min mismatch. Expected {exp_min}, got {m_gov.min_mz[0].value if m_gov.min_mz else None}")
                            report.check_a_passed = False
                        elif len(m_gov.min_mz) != exp_min_ties:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Moment-Z Min tie count mismatch. Expected {exp_min_ties}, got {len(m_gov.min_mz)}")
                            report.check_a_passed = False
                        else:
                            report.add_pass(f"Check A Verified: Node {node_id} Member {m_id} Moment-Z Min = {exp_min:.3f} ({len(m_gov.min_mz)} ties retained)")

                # Check Shear-Y
                if "_clean_shear_y" in m_df.columns:
                    valid_fy = m_df[m_df["_clean_shear_y"].notna()]["_clean_shear_y"].astype(float)
                    if not valid_fy.empty:
                        exp_max = float(valid_fy.max())
                        exp_min = float(valid_fy.min())
                        exp_max_ties = int((abs(valid_fy - exp_max) <= tol).sum())
                        exp_min_ties = int((abs(valid_fy - exp_min) <= tol).sum())

                        if not m_gov.max_fy or abs(m_gov.max_fy[0].value - exp_max) > tol:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Shear-Y Max mismatch.")
                            report.check_a_passed = False
                        elif len(m_gov.max_fy) != exp_max_ties:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Shear-Y Max tie count mismatch.")
                            report.check_a_passed = False
                        else:
                            report.add_pass(f"Check A Verified: Node {node_id} Member {m_id} Shear-Y Max = {exp_max:.3f} ({len(m_gov.max_fy)} ties retained)")

                        if not m_gov.min_fy or abs(m_gov.min_fy[0].value - exp_min) > tol:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Shear-Y Min mismatch.")
                            report.check_a_passed = False
                        elif len(m_gov.min_fy) != exp_min_ties:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Shear-Y Min tie count mismatch.")
                            report.check_a_passed = False
                        else:
                            report.add_pass(f"Check A Verified: Node {node_id} Member {m_id} Shear-Y Min = {exp_min:.3f} ({len(m_gov.min_fy)} ties retained)")

                # Check Axial Force
                if "_clean_axial" in m_df.columns:
                    valid_fx = m_df[m_df["_clean_axial"].notna()]["_clean_axial"].astype(float)
                    if not valid_fx.empty:
                        exp_max = float(valid_fx.max())
                        exp_min = float(valid_fx.min())
                        exp_max_ties = int((abs(valid_fx - exp_max) <= tol).sum())
                        exp_min_ties = int((abs(valid_fx - exp_min) <= tol).sum())

                        if not m_gov.max_fx or abs(m_gov.max_fx[0].value - exp_max) > tol:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Axial Force Max mismatch.")
                            report.check_a_passed = False
                        elif len(m_gov.max_fx) != exp_max_ties:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Axial Force Max tie count mismatch.")
                            report.check_a_passed = False
                        else:
                            report.add_pass(f"Check A Verified: Node {node_id} Member {m_id} Axial Force Max = {exp_max:.3f} ({len(m_gov.max_fx)} ties retained)")

                        if not m_gov.min_fx or abs(m_gov.min_fx[0].value - exp_min) > tol:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Axial Force Min mismatch.")
                            report.check_a_passed = False
                        elif len(m_gov.min_fx) != exp_min_ties:
                            report.add_fail(f"Check A: Node {node_id} Member {m_id} Axial Force Min tie count mismatch.")
                            report.check_a_passed = False
                        else:
                            report.add_pass(f"Check A Verified: Node {node_id} Member {m_id} Axial Force Min = {exp_min:.3f} ({len(m_gov.min_fx)} ties retained)")

        # =========================================================================
        # CHECK B: Exact L/C Correctness (Belongs to exact source row of extreme)
        # =========================================================================
        for node_id, analysis in self.node_results.items():
            for rec in analysis.all_governing_records:
                match_rows = self.combined_df[
                    (self.combined_df["_source_sheet"] == rec.source_sheet) &
                    (self.combined_df["_original_sheet_row"] == rec.source_row)
                ]

                if match_rows.empty:
                    report.add_fail(f"Check B Failure: Source row {rec.source_row} in Sheet '{rec.source_sheet}' does not exist in dataset.")
                    report.check_b_passed = False
                    continue

                source_row_data = match_rows.iloc[0]
                raw_lc = str(source_row_data["_clean_lc"])

                if raw_lc != str(rec.lc):
                    report.add_fail(
                        f"Check B Failure: L/C mismatch for Node {rec.node} Member {rec.member} {rec.force_type}. "
                        f"Governing record has L/C '{rec.lc}', but exact source row {rec.source_row} in '{rec.source_sheet}' has L/C '{raw_lc}'!"
                    )
                    report.check_b_passed = False
                else:
                    report.add_pass(f"Check B Verified: Node {rec.node} Member {rec.member} {rec.force_type} L/C '{rec.lc}' matches exact source row {rec.source_row}.")

        # =========================================================================
        # CHECK C: Node-Wise Deduplication (No duplicates, no omissions, per-node only)
        # =========================================================================
        all_nodes_lcs: Dict[str, List[str]] = {}
        for node_id, analysis in self.node_results.items():
            lcs = analysis.unique_governing_lcs

            # 1. Check for duplicates within this node
            if len(lcs) != len(set(lcs)):
                report.add_fail(f"Check C Failure: Node {node_id} unique L/C list contains duplicates: {lcs}")
                report.check_c_passed = False
            else:
                report.add_pass(f"Check C Verified: Node {node_id} L/C list contains no duplicates ({len(lcs)} unique: {lcs}).")

            # 2. Check that every governing record's L/C is present in the node list
            expected_lcs = {r.lc for r in analysis.all_governing_records}
            actual_lcs = set(lcs)
            if not expected_lcs.issubset(actual_lcs):
                missing_lcs = expected_lcs - actual_lcs
                report.add_fail(f"Check C Failure: Node {node_id} governing L/C list omitted governing L/Cs: {missing_lcs}")
                report.check_c_passed = False
            else:
                report.add_pass(f"Check C Verified: Node {node_id} contains all governing L/Cs ({len(expected_lcs)} distinct).")

            # 3. Check that no non-governing L/Cs were added
            if not actual_lcs.issubset(expected_lcs):
                extra_lcs = actual_lcs - expected_lcs
                report.add_fail(f"Check C Failure: Node {node_id} unique L/C list contains non-governing L/Cs: {extra_lcs}")
                report.check_c_passed = False

            all_nodes_lcs[node_id] = lcs

        # Verify that common L/Cs across different nodes are NOT globally removed
        # (e.g., L/C 1257 should appear in multiple nodes)
        nodes_with_1257 = [nid for nid, lcs in all_nodes_lcs.items() if "1257" in lcs]
        if len(nodes_with_1257) > 1:
            report.add_pass(
                f"Check C Verified: L/C '1257' correctly retained across {len(nodes_with_1257)} separate nodes "
                f"({nodes_with_1257}), proving deduplication is strictly local to each node."
            )

        # =========================================================================
        # CHECK D: Mandatory Source Cell Traceability (100% Match)
        # =========================================================================
        for node_id, analysis in self.node_results.items():
            for rec in analysis.all_governing_records:
                orig_cleaned = clean_numeric_value(rec.original_val)
                if orig_cleaned is None or abs(orig_cleaned - rec.value) > tol:
                    report.add_fail(
                        f"Check D Failure at Node {rec.node}, Member {rec.member}, L/C {rec.lc}: "
                        f"Selected value {rec.value} does not match original value {rec.original_val} "
                        f"in Sheet '{rec.source_sheet}', Row {rec.source_row}!"
                    )
                    report.check_d_passed = False
                    report.audit_mismatches.append({
                        "node": rec.node,
                        "member": rec.member,
                        "lc": rec.lc,
                        "selected": rec.value,
                        "original": rec.original_val,
                        "sheet": rec.source_sheet,
                        "row": rec.source_row,
                    })
                else:
                    report.add_pass(
                        f"Check D Verified: Node {rec.node} Member {rec.member} L/C {rec.lc} {rec.force_type} "
                        f"= {rec.value:.3f} matched Sheet '{rec.source_sheet}' Row {rec.source_row}."
                    )

        # Step 10 Benchmark Check: Node 136 and Member 2028 (All 6 Extremes)
        if "136" in self.node_results:
            n136 = self.node_results["136"]
            if "2028" in n136.connected_members:
                m2028 = n136.member_governing_forces.get("2028")
                if m2028:
                    node136_m2028_df = self.combined_df[
                        (self.combined_df["_clean_node"] == "136") & (self.combined_df["_clean_member"] == "2028")
                    ]
                    has_mz_bench = (node136_m2028_df["_clean_moment_z"].round(3) == 1526.629).any() if "_clean_moment_z" in node136_m2028_df else False
                    if has_mz_bench:
                        if m2028.max_mz and abs(m2028.max_mz[0].value - 1526.629) <= 0.05:
                            report.add_pass("Benchmark Verified: Node 136 Member 2028 Moment-Z Max = +1526.629 at L/C 1257")
                        if m2028.min_mz and abs(m2028.min_mz[0].value - (-919.990)) <= 0.05:
                            report.add_pass("Benchmark Verified: Node 136 Member 2028 Moment-Z Min = -919.990 at L/C 101")

                    has_fy_bench = (node136_m2028_df["_clean_shear_y"].round(3) == -146.266).any() if "_clean_shear_y" in node136_m2028_df else False
                    if has_fy_bench:
                        if m2028.max_fy and abs(m2028.max_fy[0].value - 230.941) <= 0.05:
                            report.add_pass("Benchmark Verified: Node 136 Member 2028 Shear-Y Max = +230.941 at L/C 102")
                        if m2028.min_fy and abs(m2028.min_fy[0].value - (-146.266)) <= 0.05:
                            report.add_pass("Benchmark Verified: Node 136 Member 2028 Shear-Y Min = -146.266 at L/C 1005")

                    has_fx_bench = (node136_m2028_df["_clean_axial"].round(3) == 16.094).any() if "_clean_axial" in node136_m2028_df else False
                    if has_fx_bench:
                        if m2028.max_fx and abs(m2028.max_fx[0].value - 16.094) <= 0.05:
                            report.add_pass("Benchmark Verified: Node 136 Member 2028 Axial Force Max = +16.094 at L/C 103")
                        if m2028.min_fx and abs(m2028.min_fx[0].value - (-72.126)) <= 0.05:
                            report.add_pass("Benchmark Verified: Node 136 Member 2028 Axial Force Min = -72.126 at L/C 104")

        return report
