"""
AI/ML-Powered STAAD.Pro Governing Force Extraction System.
Combines Scikit-learn TF-IDF semantic interpretation with deterministic engineering calculations.
"""

from typing import Optional, Callable, Dict, Any, Tuple, List
import os

from .config import APP_NAME, APP_VERSION, DEFAULT_TARGET_SHEETS
from .ml_engine import (
    MLColumnClassifier,
    MLLearningStore,
    MLDataQualityInspector,
    MLAnomalyDetector,
    CAT_NODE,
    CAT_MEMBER,
    CAT_LC,
    CAT_MOMENT_Z,
    CAT_SHEAR_Y,
    CAT_AXIAL,
    CAT_IGNORE,
)
from .reader import StaadReader, DEFAULT_MAX_HEADER_SCAN_ROWS
from .analyzer import StaadAnalyzer, NodeAnalysis, GoverningRecord, ConnectedMemberDetail, MemberGoverningForces
from .exporter import StaadExporter
from .validator import StaadValidator, ValidationReport


def process_staad_file(
    input_file: str,
    output_file: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    target_sheets: Optional[List[str]] = None,
    learning_store: Optional[MLLearningStore] = None,
    include_summary_sheets: bool = False,
) -> Tuple[str, Dict[str, NodeAnalysis], ValidationReport, List[str]]:
    """
    Complete AI/ML-Assisted Processing Pipeline:
      1. ML semantic header interpretation with Scikit-learn character n-gram TF-IDF.
      2. Reads target worksheets: R-L, R-R, JB-F, JB-B (and any other member force sheets).
      3. Constructs the node-member relationship matrix.
      4. Calculates deterministic governing forces (numerical Max & Min) for each (Node, Member) combination.
      5. Gathers governing L/Cs and eliminates duplicates per node.
      6. Runs Scikit-learn IsolationForest anomaly detection on numeric force distributions.
      7. Generates 5-sheet formatted Excel workbook:
         - Sheet 1: Member-Wise Governing Summary
         - Sheet 2: Unique LC by Node
         - Sheet 3: Governing Force Records
         - Sheet 4: Combined Original Data (with Green/Red highlights)
         - Sheet 5: ML Validation Report (Worksheets, Mappings, Anomalies & 100% Traceability Audit)
      8. Validates 100% mathematical accuracy and source cell traceability.
    """
    def report_progress(pct: float, message: str):
        if progress_callback:
            progress_callback(pct, message)

    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    if not output_file:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_Governing_Forces_STEP1.xlsx"

    report_progress(5, "Initializing AI/ML Column Classifier and Knowledge Store...")
    l_store = learning_store or MLLearningStore()
    classifier = MLColumnClassifier(learning_store=l_store)

    report_progress(15, "Reading input workbook and predicting column semantics with ML...")
    reader = StaadReader(input_file, target_sheets=target_sheets, classifier=classifier)
    combined_df, per_sheet_dfs, invalid_records = reader.load_worksheets()

    report_progress(
        30,
        f"ML classified {len(reader.sheets_found)} worksheets ({', '.join(reader.sheets_found)}) "
        f"with {len(combined_df)} records across {combined_df['_clean_node'].nunique()} unique nodes."
    )

    report_progress(45, "Running Scikit-learn IsolationForest anomaly detection and building Node-Member intelligence...")
    anomaly_detector = MLAnomalyDetector(contamination=0.01)
    anomalies = anomaly_detector.detect_anomalies(combined_df)

    analyzer = StaadAnalyzer(combined_df=combined_df)
    node_results = analyzer.run_analysis()

    report_progress(65, f"Identified governing forces for {len(node_results)} nodes across all connected members.")

    report_progress(70, "Executing automated validation (Checks A, B, C, D & Source Traceability)...")
    validator = StaadValidator(
        node_results=node_results,
        combined_df=combined_df,
        sheets_loaded=reader.sheets_found,
        invalid_force_records=invalid_records,
    )
    validation_report = validator.validate()

    report_progress(80, "Generating formatted Excel workbook...")
    exporter = StaadExporter(
        node_results=node_results,
        combined_df=combined_df,
        highlight_map=analyzer.highlight_map,
        sheets_loaded=reader.sheets_found,
        per_sheet_col_mappings=reader.per_sheet_col_mappings,
        per_sheet_confidences=reader.per_sheet_confidences,
        per_sheet_bases=reader.per_sheet_bases,
        anomalies=anomalies,
        validation_report=validation_report,
        input_file_path=input_file,
        include_summary_sheets=include_summary_sheets,
    )
    final_output_file = exporter.generate_workbook(output_file)
    if final_output_file != output_file:
        report_progress(90, f"Target file locked in Excel; auto-saved to: {os.path.basename(final_output_file)}")

    if not validation_report.is_valid:
        report_progress(95, f"Validation warning: {validation_report.checks_failed} checks failed!")
    else:
        report_progress(100, f"Extraction Complete! 100% verified ({validation_report.checks_passed} checks passed).")

    return final_output_file, node_results, validation_report, reader.sheets_found


__all__ = [
    "APP_NAME",
    "APP_VERSION",
    "DEFAULT_TARGET_SHEETS",
    "MLColumnClassifier",
    "MLSheetClassifier",
    "MLLearningStore",
    "MLDataQualityInspector",
    "StaadReader",
    "StaadAnalyzer",
    "StaadExporter",
    "StaadValidator",
    "NodeAnalysis",
    "GoverningRecord",
    "ConnectedMemberDetail",
    "ValidationReport",
    "process_staad_file",
]
