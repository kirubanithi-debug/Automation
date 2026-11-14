"""
Core Structural Force Analyzer and Node-Member Intelligence Engine.
Builds the node-member relationship matrix across worksheets (R-L, R-R, JB-F, JB-B),
calculates exact numerical governing forces (max & min Moment-Z, Shear-Y, Axial Force),
preserves signs, retains all ties, and tracks exact source cell coordinates.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Set, Tuple
import pandas as pd

from .config import FORCE_MOMENT_Z, FORCE_SHEAR_Y, FORCE_AXIAL


@dataclass
class ConnectedMemberDetail:
    """Detailed structural member framing information for a node."""
    node: str
    member: str
    source_sheet: str
    record_count: int
    available_lcs: List[str]


@dataclass
class GoverningRecord:
    """Represents a single governing force record (Max or Min) for a node."""
    node: str
    member: str
    lc: str
    force_type: str
    condition: str  # 'MAX' or 'MIN'
    value: float
    original_val: Any  # Raw unformatted value from Excel cell
    source_sheet: str
    source_row: int   # 1-based Excel row in source worksheet
    row_index: int


@dataclass
class MemberGoverningForces:
    """Holds the six governing extreme values for a single (Node, Member) combination."""
    node: str
    member: str
    source_sheet: str
    record_count: int
    available_lcs: List[str]

    max_mz: List[GoverningRecord] = field(default_factory=list)
    min_mz: List[GoverningRecord] = field(default_factory=list)

    max_fy: List[GoverningRecord] = field(default_factory=list)
    min_fy: List[GoverningRecord] = field(default_factory=list)

    max_fx: List[GoverningRecord] = field(default_factory=list)
    min_fx: List[GoverningRecord] = field(default_factory=list)

    all_governing_records: List[GoverningRecord] = field(default_factory=list)
    governing_lcs: List[str] = field(default_factory=list)


@dataclass
class NodeAnalysis:
    """Full analysis results for a structural joint (node)."""
    node: str
    connected_members: List[str]  # Unique member IDs connected to this node
    connected_member_details: List[ConnectedMemberDetail] = field(default_factory=list)
    member_governing_forces: Dict[str, MemberGoverningForces] = field(default_factory=dict)
    total_records: int = 0

    all_governing_records: List[GoverningRecord] = field(default_factory=list)
    selected_lcs: List[str] = field(default_factory=list)          # All collected L/Cs
    unique_governing_lcs: List[str] = field(default_factory=list)  # Deduplicated unique list

    @property
    def max_mz(self) -> List[GoverningRecord]:
        return [r for r in self.all_governing_records if r.force_type == FORCE_MOMENT_Z and r.condition == "MAX"]

    @property
    def min_mz(self) -> List[GoverningRecord]:
        return [r for r in self.all_governing_records if r.force_type == FORCE_MOMENT_Z and r.condition == "MIN"]

    @property
    def max_fy(self) -> List[GoverningRecord]:
        return [r for r in self.all_governing_records if r.force_type == FORCE_SHEAR_Y and r.condition == "MAX"]

    @property
    def min_fy(self) -> List[GoverningRecord]:
        return [r for r in self.all_governing_records if r.force_type == FORCE_SHEAR_Y and r.condition == "MIN"]

    @property
    def max_fx(self) -> List[GoverningRecord]:
        return [r for r in self.all_governing_records if r.force_type == FORCE_AXIAL and r.condition == "MAX"]

    @property
    def min_fx(self) -> List[GoverningRecord]:
        return [r for r in self.all_governing_records if r.force_type == FORCE_AXIAL and r.condition == "MIN"]


class StaadAnalyzer:
    """
    Analyzes combined STAAD.Pro member forces across all worksheets.
    Constructs the node-member relationship matrix and extracts deterministic governing forces
    independently for each (Node, Member) combination, followed by node-wise unique L/C extraction.
    """

    def __init__(self, combined_df: pd.DataFrame):
        self.combined_df = combined_df
        self.nodes: List[str] = []
        self.node_results: Dict[str, NodeAnalysis] = {}

        # Highlighting lookup: (source_sheet, source_row, force_key) -> 'MAX' | 'MIN' | 'BOTH'
        self.highlight_map: Dict[Tuple[str, int, str], str] = {}

    def run_analysis(self) -> Dict[str, NodeAnalysis]:
        """
        Executes member-wise governing force identification and node-wise unique L/C deduplication.
        """
        # 1. Identify all unique node numbers across the combined data
        self.nodes = list(dict.fromkeys(self.combined_df["_clean_node"].tolist()))

        for node_id in self.nodes:
            node_df = self.combined_df[self.combined_df["_clean_node"] == node_id]

            # 2. Identify every unique member connected to this node
            connected_members = sorted(
                list(dict.fromkeys(node_df["_clean_member"].tolist())),
                key=lambda x: (int(x) if x.isdigit() else float("inf"), x),
            )

            member_details: List[ConnectedMemberDetail] = []
            member_gov_map: Dict[str, MemberGoverningForces] = {}
            node_all_recs: List[GoverningRecord] = []
            node_selected_lcs: List[str] = []

            # 3. Process each connected member independently
            for m_id in connected_members:
                m_sub = node_df[node_df["_clean_member"] == m_id]
                source_sheet = str(m_sub["_source_sheet"].iloc[0])
                m_lcs = list(dict.fromkeys(m_sub["_clean_lc"].tolist()))

                member_details.append(
                    ConnectedMemberDetail(
                        node=node_id,
                        member=m_id,
                        source_sheet=source_sheet,
                        record_count=len(m_sub),
                        available_lcs=m_lcs,
                    )
                )

                m_gov = MemberGoverningForces(
                    node=node_id,
                    member=m_id,
                    source_sheet=source_sheet,
                    record_count=len(m_sub),
                    available_lcs=m_lcs,
                )

                # A. Moment-Z (kN-m)
                if "_clean_moment_z" in m_sub.columns:
                    m_max_mz, m_min_mz = self._find_extremes(
                        m_sub, "_clean_moment_z", "_raw_val_moment_z", FORCE_MOMENT_Z, "moment_z"
                    )
                    m_gov.max_mz = m_max_mz
                    m_gov.min_mz = m_min_mz

                # B. Shear-Y (kN)
                if "_clean_shear_y" in m_sub.columns:
                    m_max_fy, m_min_fy = self._find_extremes(
                        m_sub, "_clean_shear_y", "_raw_val_shear_y", FORCE_SHEAR_Y, "shear_y"
                    )
                    m_gov.max_fy = m_max_fy
                    m_gov.min_fy = m_min_fy

                # C. Axial Force (kN)
                if "_clean_axial" in m_sub.columns:
                    m_max_fx, m_min_fx = self._find_extremes(
                        m_sub, "_clean_axial", "_raw_val_axial", FORCE_AXIAL, "axial"
                    )
                    m_gov.max_fx = m_max_fx
                    m_gov.min_fx = m_min_fx

                m_recs: List[GoverningRecord] = []
                m_recs.extend(m_gov.max_mz)
                m_recs.extend(m_gov.min_mz)
                m_recs.extend(m_gov.max_fy)
                m_recs.extend(m_gov.min_fy)
                m_recs.extend(m_gov.max_fx)
                m_recs.extend(m_gov.min_fx)

                m_gov.all_governing_records = m_recs
                m_gov.governing_lcs = [r.lc for r in m_recs]

                member_gov_map[m_id] = m_gov
                node_all_recs.extend(m_recs)
                node_selected_lcs.extend(m_gov.governing_lcs)

            # 4. Node-wise Unique L/C Extraction
            unique_lcs: List[str] = []
            seen: Set[str] = set()
            for lc in node_selected_lcs:
                if lc not in seen:
                    seen.add(lc)
                    unique_lcs.append(lc)

            # Sort unique L/Cs numerically where possible
            unique_lcs_sorted = sorted(
                unique_lcs,
                key=lambda x: (int(x) if x.isdigit() else float("inf"), x),
            )

            node_analysis = NodeAnalysis(
                node=node_id,
                connected_members=connected_members,
                connected_member_details=member_details,
                member_governing_forces=member_gov_map,
                total_records=len(node_df),
                all_governing_records=node_all_recs,
                selected_lcs=node_selected_lcs,
                unique_governing_lcs=unique_lcs_sorted,
            )

            self.node_results[node_id] = node_analysis

        return self.node_results

    def _find_extremes(
        self,
        node_df: pd.DataFrame,
        clean_col: str,
        raw_col: str,
        force_type_label: str,
        force_key: str,
    ) -> Tuple[List[GoverningRecord], List[GoverningRecord]]:
        """
        Finds numerical maximum and minimum values on the real number line,
        retaining all tied records.
        """
        valid_df = node_df[node_df[clean_col].notna()]
        if valid_df.empty:
            return [], []

        series = valid_df[clean_col].astype(float)
        max_val = float(series.max())
        min_val = float(series.min())

        tol = 1e-6
        max_records: List[GoverningRecord] = []
        min_records: List[GoverningRecord] = []

        for idx, row in valid_df.iterrows():
            val = float(row[clean_col])
            raw_val = row[raw_col] if raw_col in row else val
            orig_sheet = str(row["_source_sheet"])
            orig_row = int(row["_original_sheet_row"])
            node_val = str(row["_clean_node"])
            member_val = str(row["_clean_member"])
            lc_val = str(row["_clean_lc"])

            is_max = abs(val - max_val) <= tol
            is_min = abs(val - min_val) <= tol

            if is_max:
                rec = GoverningRecord(
                    node=node_val,
                    member=member_val,
                    lc=lc_val,
                    force_type=force_type_label,
                    condition="MAX",
                    value=val,
                    original_val=raw_val,
                    source_sheet=orig_sheet,
                    source_row=orig_row,
                    row_index=idx,
                )
                max_records.append(rec)
                self._record_highlight(orig_sheet, orig_row, force_key, "MAX")

            if is_min:
                rec = GoverningRecord(
                    node=node_val,
                    member=member_val,
                    lc=lc_val,
                    force_type=force_type_label,
                    condition="MIN",
                    value=val,
                    original_val=raw_val,
                    source_sheet=orig_sheet,
                    source_row=orig_row,
                    row_index=idx,
                )
                min_records.append(rec)
                self._record_highlight(orig_sheet, orig_row, force_key, "MIN")

        return max_records, min_records

    def _record_highlight(self, sheet: str, row: int, force_key: str, condition: str):
        """Records highlight state for a specific cell in a specific source worksheet."""
        key = (sheet, row, force_key)
        existing = self.highlight_map.get(key)
        if existing:
            if existing != condition:
                self.highlight_map[key] = "BOTH"
        else:
            self.highlight_map[key] = condition
