"""
Core Structural Force Analyzer and Node-Member Intelligence Engine.
Builds the node-member relationship matrix across worksheets (R-L, R-R, JB-F, JB-B),
calculates exact numerical governing forces (max & min Moment-Z, Shear-Y, Axial Force),
preserves signs, retains all ties, and tracks exact source cell coordinates.
"""

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Set, Tuple
import numpy as np
import pandas as pd

from .config import (
    FORCE_MOMENT_Z,
    FORCE_SHEAR_Y,
    FORCE_AXIAL,
    BEAM_TYPE_RAFTER_LEFT,
    BEAM_TYPE_RAFTER_RIGHT,
    BEAM_TYPE_RAFTER,
    BEAM_TYPE_JACK_BEAM_FRONT,
    BEAM_TYPE_JACK_BEAM_BACK,
    BEAM_TYPE_JACK_BEAM,
    BEAM_TYPE_TIE_BEAM,
    BEAM_TYPE_COLUMN,
    BEAM_TYPE_BRACING,
    BEAM_TYPE_BEAM,
    CONN_TYPE_MOMENT,
    CONN_TYPE_PINNED,
    CONN_TYPE_AXIAL_STRUT,
)


@dataclass
class ConnectedMemberDetail:
    """Detailed structural member framing information for a node."""
    node: str
    member: str
    source_sheet: str
    record_count: int
    available_lcs: List[str]
    beam_type: str = BEAM_TYPE_BEAM
    connection_type: str = CONN_TYPE_MOMENT


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
    beam_type: str = ""
    connection_type: str = ""


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
    is_zero_moment: bool = False
    beam_type: str = BEAM_TYPE_BEAM
    connection_type: str = CONN_TYPE_MOMENT


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

                # Automatically identify Beam Type and Connection Type
                beam_type, conn_type, is_zero_moment = self._classify_beam_and_connection(source_sheet, m_sub)

                member_details.append(
                    ConnectedMemberDetail(
                        node=node_id,
                        member=m_id,
                        source_sheet=source_sheet,
                        record_count=len(m_sub),
                        available_lcs=m_lcs,
                        beam_type=beam_type,
                        connection_type=conn_type,
                    )
                )

                m_gov = MemberGoverningForces(
                    node=node_id,
                    member=m_id,
                    source_sheet=source_sheet,
                    record_count=len(m_sub),
                    available_lcs=m_lcs,
                    is_zero_moment=is_zero_moment,
                    beam_type=beam_type,
                    connection_type=conn_type,
                )

                # Branch based on connection type:
                # 1. Pinned / Shear Connection (Moment-Z is zero: |Mz| <= 1e-4)
                #    Bypasses Moment-Z, filters Shear-Y Max/Min candidates, then filters Axial Force Max/Min
                #    (Matches exact manual workflow from Screen Recording 2026-11-14 131115.mp4)
                # 2. Moment Connection (Moment-Z != 0)
                #    Evaluates all 3 force components independently (Mz, Fy, Fx)
                if is_zero_moment:
                    m_recs = self._find_zero_moment_extremes(
                        m_sub, m_gov, node_id, m_id, source_sheet,
                        beam_type=beam_type, connection_type=conn_type,
                    )
                else:
                    # A. Moment-Z (kN-m)
                    if "_clean_moment_z" in m_sub.columns:
                        m_max_mz, m_min_mz = self._find_extremes(
                            m_sub, "_clean_moment_z", "_raw_val_moment_z", FORCE_MOMENT_Z, "moment_z",
                            beam_type=beam_type, connection_type=conn_type,
                        )
                        m_gov.max_mz = m_max_mz
                        m_gov.min_mz = m_min_mz

                    # B. Shear-Y (kN)
                    if "_clean_shear_y" in m_sub.columns:
                        m_max_fy, m_min_fy = self._find_extremes(
                            m_sub, "_clean_shear_y", "_raw_val_shear_y", FORCE_SHEAR_Y, "shear_y",
                            beam_type=beam_type, connection_type=conn_type,
                        )
                        m_gov.max_fy = m_max_fy
                        m_gov.min_fy = m_min_fy

                    # C. Axial Force (kN)
                    if "_clean_axial" in m_sub.columns:
                        m_max_fx, m_min_fx = self._find_extremes(
                            m_sub, "_clean_axial", "_raw_val_axial", FORCE_AXIAL, "axial",
                            beam_type=beam_type, connection_type=conn_type,
                        )
                        m_gov.max_fx = m_max_fx
                        m_gov.min_fx = m_min_fx

                    m_recs = []
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

    def _classify_beam_and_connection(
        self, source_sheet: str, m_sub: pd.DataFrame
    ) -> Tuple[str, str, bool]:
        """
        Automatically classifies structural beam type and connection behavior:
        1. Layer 1: Semantic parsing of source worksheet and member properties.
        2. Layer 2: Physics-based force profile analysis (evaluating Moment-Z, Shear-Y, and Axial Force).
        Returns: (beam_type, connection_type, is_zero_moment)
        """
        s = source_sheet.strip().lower()
        s_clean = re.sub(r"[\s\-_]+", "", s)

        # Check whether Moment-Z is effectively zero (|Mz| <= 1e-4) across all load cases
        max_mz_abs = 0.0
        if "_clean_moment_z" in m_sub.columns:
            valid_mz = m_sub[m_sub["_clean_moment_z"].notna()]["_clean_moment_z"].astype(float)
            if not valid_mz.empty:
                max_mz_abs = float(valid_mz.abs().max())

        max_fy_abs = 0.0
        if "_clean_shear_y" in m_sub.columns:
            valid_fy = m_sub[m_sub["_clean_shear_y"].notna()]["_clean_shear_y"].astype(float)
            if not valid_fy.empty:
                max_fy_abs = float(valid_fy.abs().max())

        max_fx_abs = 0.0
        if "_clean_axial" in m_sub.columns:
            valid_fx = m_sub[m_sub["_clean_axial"].notna()]["_clean_axial"].astype(float)
            if not valid_fx.empty:
                max_fx_abs = float(valid_fx.abs().max())

        is_zero_moment = (max_mz_abs <= 1e-4)

        # Layer 1: Semantic Classification
        if s_clean in ("rl", "rafterleft", "rafterl") or "r-l" in s or "rafter-l" in s:
            beam_type = BEAM_TYPE_RAFTER_LEFT
            conn_type = CONN_TYPE_MOMENT if not is_zero_moment else CONN_TYPE_PINNED
        elif s_clean in ("rr", "rafterright", "rafterr") or "r-r" in s or "rafter-r" in s:
            beam_type = BEAM_TYPE_RAFTER_RIGHT
            conn_type = CONN_TYPE_MOMENT if not is_zero_moment else CONN_TYPE_PINNED
        elif s_clean in ("jbf", "jackfront") or "jb-f" in s or "jack-f" in s:
            beam_type = BEAM_TYPE_JACK_BEAM_FRONT
            conn_type = CONN_TYPE_PINNED if is_zero_moment else CONN_TYPE_MOMENT
        elif s_clean in ("jbb", "jackback") or "jb-b" in s or "jack-b" in s:
            beam_type = BEAM_TYPE_JACK_BEAM_BACK
            conn_type = CONN_TYPE_PINNED if is_zero_moment else CONN_TYPE_MOMENT
        elif s_clean in ("jb", "jack", "jackbeam") or "jack" in s:
            beam_type = BEAM_TYPE_JACK_BEAM
            conn_type = CONN_TYPE_PINNED if is_zero_moment else CONN_TYPE_MOMENT
        elif s_clean in ("tb", "tbf", "tbb", "tie", "tiebeam", "strut") or "tie" in s:
            beam_type = BEAM_TYPE_TIE_BEAM
            conn_type = CONN_TYPE_AXIAL_STRUT if is_zero_moment else CONN_TYPE_MOMENT
        elif "col" in s:
            beam_type = BEAM_TYPE_COLUMN
            conn_type = CONN_TYPE_MOMENT if not is_zero_moment else CONN_TYPE_PINNED
        elif "br" in s or "brace" in s:
            beam_type = BEAM_TYPE_BRACING
            conn_type = CONN_TYPE_AXIAL_STRUT if is_zero_moment else CONN_TYPE_MOMENT
        elif "rafter" in s or s_clean == "r":
            beam_type = BEAM_TYPE_RAFTER
            conn_type = CONN_TYPE_MOMENT if not is_zero_moment else CONN_TYPE_PINNED
        else:
            # Layer 2: Force Physics Signature (generic sheets like "Sheet1" or "Forces")
            if not is_zero_moment:
                beam_type = "Moment Beam / Rafter"
                conn_type = CONN_TYPE_MOMENT
            else:
                if max_fx_abs > 2.0 * (max_fy_abs + 1e-3):
                    beam_type = BEAM_TYPE_TIE_BEAM
                    conn_type = CONN_TYPE_AXIAL_STRUT
                else:
                    beam_type = BEAM_TYPE_JACK_BEAM
                    conn_type = CONN_TYPE_PINNED

        return beam_type, conn_type, is_zero_moment

    def _find_extremes(
        self,
        node_df: pd.DataFrame,
        clean_col: str,
        raw_col: str,
        force_type_label: str,
        force_key: str,
        beam_type: str = "",
        connection_type: str = "",
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
                    beam_type=beam_type,
                    connection_type=connection_type,
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
                    beam_type=beam_type,
                    connection_type=connection_type,
                )
                min_records.append(rec)
                self._record_highlight(orig_sheet, orig_row, force_key, "MIN")

        return max_records, min_records

    def _find_zero_moment_extremes(
        self,
        m_sub: pd.DataFrame,
        m_gov: MemberGoverningForces,
        node_id: str,
        m_id: str,
        source_sheet: str,
        beam_type: str = "",
        connection_type: str = "",
    ) -> List[GoverningRecord]:
        """
        Zero-Moment (Pinned Connection) Logic:
        Reproduces the manual Excel engineering workflow from Screen Recording 2026-11-14 131115.mp4:
        1. When Moment-Z is zero (e.g. pinned members in JB-F, JB-B), bypass Moment-Z.
        2. Filter Shear-Y (kN) to find its numerical Maximum and Minimum.
        3. Form candidate pool of rows with extreme Shear-Y (Max or Min).
        4. From this candidate pool, find the extreme Axial Force (Max and Min).
        5. Select the resulting governing rows (typically 2 rows) and retain exact ties.
        """
        tol = 1e-6
        if "_clean_shear_y" not in m_sub.columns or "_clean_axial" not in m_sub.columns:
            return []

        valid_df = m_sub[m_sub["_clean_shear_y"].notna() & m_sub["_clean_axial"].notna()].copy()
        if valid_df.empty:
            return []

        valid_df["_clean_shear_y"] = valid_df["_clean_shear_y"].astype(float)
        valid_df["_clean_axial"] = valid_df["_clean_axial"].astype(float)

        max_fy_val = float(valid_df["_clean_shear_y"].max())
        min_fy_val = float(valid_df["_clean_shear_y"].min())

        # Candidate pool: rows where Shear-Y is at its extreme (Max or Min)
        is_fy_max = np.isclose(valid_df["_clean_shear_y"], max_fy_val, atol=tol, rtol=tol)
        is_fy_min = np.isclose(valid_df["_clean_shear_y"], min_fy_val, atol=tol, rtol=tol)
        candidates_df = valid_df[is_fy_max | is_fy_min]

        if candidates_df.empty:
            return []

        # From candidate pool, find extreme Axial Force
        cand_max_fx = float(candidates_df["_clean_axial"].max())
        cand_min_fx = float(candidates_df["_clean_axial"].min())

        is_fx_max = np.isclose(candidates_df["_clean_axial"], cand_max_fx, atol=tol, rtol=tol)
        is_fx_min = np.isclose(candidates_df["_clean_axial"], cand_min_fx, atol=tol, rtol=tol)

        selected_indices = set(candidates_df[is_fx_max | is_fx_min].index)

        # Safeguard: ensure both Shear-Y Max and Min branches are represented
        max_branch_mask = np.isclose(candidates_df["_clean_shear_y"], max_fy_val, atol=tol, rtol=tol)
        min_branch_mask = np.isclose(candidates_df["_clean_shear_y"], min_fy_val, atol=tol, rtol=tol)
        max_branch_indices = set(candidates_df[max_branch_mask].index)
        min_branch_indices = set(candidates_df[min_branch_mask].index)

        if not (selected_indices & max_branch_indices):
            sub_max = candidates_df[max_branch_mask]
            if not sub_max.empty:
                # pick extreme axial from max_fy branch
                sub_fx_min = float(sub_max["_clean_axial"].min())
                match_idx = sub_max[np.isclose(sub_max["_clean_axial"], sub_fx_min, atol=tol, rtol=tol)].index
                selected_indices.update(match_idx)

        if not (selected_indices & min_branch_indices):
            sub_min = candidates_df[min_branch_mask]
            if not sub_min.empty:
                # pick extreme axial from min_fy branch
                sub_fx_max = float(sub_min["_clean_axial"].max())
                match_idx = sub_min[np.isclose(sub_min["_clean_axial"], sub_fx_max, atol=tol, rtol=tol)].index
                selected_indices.update(match_idx)

        # Now build governing records for all selected rows
        governing_records: List[GoverningRecord] = []
        for idx in sorted(selected_indices):
            row = valid_df.loc[idx]
            orig_sheet = str(row["_source_sheet"])
            orig_row = int(row["_original_sheet_row"])
            node_val = str(row["_clean_node"])
            member_val = str(row["_clean_member"])
            lc_val = str(row["_clean_lc"])
            fy_val = float(row["_clean_shear_y"])
            fx_val = float(row["_clean_axial"])
            raw_fy = row["_raw_val_shear_y"] if "_raw_val_shear_y" in row else fy_val
            raw_fx = row["_raw_val_axial"] if "_raw_val_axial" in row else fx_val

            # Record for Shear-Y
            if np.isclose(fy_val, max_fy_val, atol=tol, rtol=tol):
                rec_fy = GoverningRecord(
                    node=node_val, member=member_val, lc=lc_val,
                    force_type=FORCE_SHEAR_Y, condition="MAX", value=fy_val,
                    original_val=raw_fy, source_sheet=orig_sheet, source_row=orig_row, row_index=idx,
                    beam_type=beam_type, connection_type=connection_type,
                )
                m_gov.max_fy.append(rec_fy)
                governing_records.append(rec_fy)
                self._record_highlight(orig_sheet, orig_row, "shear_y", "MAX")

            if np.isclose(fy_val, min_fy_val, atol=tol, rtol=tol):
                rec_fy = GoverningRecord(
                    node=node_val, member=member_val, lc=lc_val,
                    force_type=FORCE_SHEAR_Y, condition="MIN", value=fy_val,
                    original_val=raw_fy, source_sheet=orig_sheet, source_row=orig_row, row_index=idx,
                    beam_type=beam_type, connection_type=connection_type,
                )
                m_gov.min_fy.append(rec_fy)
                governing_records.append(rec_fy)
                self._record_highlight(orig_sheet, orig_row, "shear_y", "MIN")

            # Record for Axial Force
            if np.isclose(fx_val, cand_max_fx, atol=tol, rtol=tol):
                rec_fx = GoverningRecord(
                    node=node_val, member=member_val, lc=lc_val,
                    force_type=FORCE_AXIAL, condition="MAX", value=fx_val,
                    original_val=raw_fx, source_sheet=orig_sheet, source_row=orig_row, row_index=idx,
                    beam_type=beam_type, connection_type=connection_type,
                )
                m_gov.max_fx.append(rec_fx)
                governing_records.append(rec_fx)
                self._record_highlight(orig_sheet, orig_row, "axial", "MAX")

            if np.isclose(fx_val, cand_min_fx, atol=tol, rtol=tol):
                rec_fx = GoverningRecord(
                    node=node_val, member=member_val, lc=lc_val,
                    force_type=FORCE_AXIAL, condition="MIN", value=fx_val,
                    original_val=raw_fx, source_sheet=orig_sheet, source_row=orig_row, row_index=idx,
                    beam_type=beam_type, connection_type=connection_type,
                )
                m_gov.min_fx.append(rec_fx)
                governing_records.append(rec_fx)
                self._record_highlight(orig_sheet, orig_row, "axial", "MIN")

        return governing_records

    def _record_highlight(self, sheet: str, row: int, force_key: str, condition: str):
        """Records highlight state for a specific cell in a specific source worksheet."""
        key = (sheet, row, force_key)
        existing = self.highlight_map.get(key)
        if existing:
            if existing != condition:
                self.highlight_map[key] = "BOTH"
        else:
            self.highlight_map[key] = condition
