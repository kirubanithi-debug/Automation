"""
Comprehensive Test Suite for AI/ML-Powered STAAD.Pro Force Extraction System.
Tests ML column classification, learning store feedback, node-member intelligence,
exact benchmark validation (Node 136 / Member 2028), and 100% source-value audit reporting.
"""

import os
import unittest
import openpyxl
import pandas as pd

from staad_processor import (
    process_staad_file,
    StaadReader,
    StaadAnalyzer,
    StaadValidator,
    MLColumnClassifier,
    MLLearningStore,
)
from tests.generate_sample_staad import create_sample_staad_4sheets


class TestStaadProcessorAIML(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.sample_input = "d:\\Automation\\sample_data\\sample_staad_4sheets.xlsx"
        create_sample_staad_4sheets(cls.sample_input)
        cls.output_file = "d:\\Automation\\sample_data\\test_output_aiml.xlsx"

    def test_01_ml_column_classifier(self):
        """Test Scikit-learn TF-IDF column classifier on various STAAD column names."""
        classifier = MLColumnClassifier()

        test_cases = [
            ("Node", "node"),
            ("Joint ID", "node"),
            ("Node Number", "node"),
            ("Beam", "member"),
            ("Member ID", "member"),
            ("Element No", "member"),
            ("L/C", "lc"),
            ("Load Case", "lc"),
            ("Loading Combination", "lc"),
            ("Moment-Z (kN-m)", "moment_z"),
            ("Mz (kNm)", "moment_z"),
            ("Bending-Z", "moment_z"),
            ("Shear-Y (kN)", "shear_y"),
            ("Fy (kN)", "shear_y"),
            ("Axial Force (kN)", "axial"),
            ("Fx (kN)", "axial"),
            # Exclusions
            ("Moment-Y (kN-m)", "ignore"),
            ("My", "ignore"),
            ("Shear-Z (kN)", "ignore"),
            ("Fz", "ignore"),
            ("Torsion (kN-m)", "ignore"),
        ]

        for raw_header, expected_cat in test_cases:
            pred_cat, conf = classifier.predict_column(raw_header)
            self.assertEqual(
                pred_cat, expected_cat,
                f"Header '{raw_header}' expected '{expected_cat}', got '{pred_cat}' (conf: {conf:.2f})"
            )
            self.assertGreater(conf, 0.50, f"Confidence for '{raw_header}' should be >= 0.50")

    def test_02_ml_learning_store_feedback(self):
        """Test feedback mechanism: user-confirmed mappings are learned and boosted."""
        temp_store_path = "d:\\Automation\\sample_data\\test_learning_store.json"
        if os.path.exists(temp_store_path):
            os.remove(temp_store_path)

        store = MLLearningStore(temp_store_path)
        classifier = MLColumnClassifier(learning_store=store)

        # Custom non-standard header
        custom_header = "CustomJointCol_X9"
        pred1, conf1 = classifier.predict_column(custom_header)

        # User provides feedback / approves mapping as 'node'
        store.add_feedback(custom_header, "node")

        # Re-predict: should now return 'node' with confidence 1.0
        pred2, conf2 = classifier.predict_column(custom_header)
        self.assertEqual(pred2, "node")
        self.assertEqual(conf2, 1.0)

        if os.path.exists(temp_store_path):
            os.remove(temp_store_path)

    def test_03_read_all_four_worksheets_with_ml(self):
        """Test ML reader identifying columns across R-L, R-R, JB-F, JB-B."""
        reader = StaadReader(self.sample_input)
        combined_df, per_sheet_dfs, invalids = reader.load_worksheets()

        self.assertIsNotNone(combined_df)
        self.assertEqual(len(reader.sheets_found), 4)
        self.assertEqual(set(reader.sheets_found), {"R-L", "R-R", "JB-F", "JB-B"})

        for s_name in ["R-L", "R-R", "JB-F", "JB-B"]:
            mapping = reader.per_sheet_col_mappings[s_name]
            self.assertEqual(mapping["member"], "Beam")
            self.assertEqual(mapping["lc"], "L/C")
            self.assertEqual(mapping["node"], "Node")
            self.assertEqual(mapping["moment_z"], "Moment-Z (kN-m)")
            self.assertEqual(mapping["shear_y"], "Shear-Y (kN)")
            self.assertEqual(mapping["axial"], "Axial Force (kN)")

    def test_04_node_and_connected_member_details(self):
        """Test Node-Member relationship intelligence table."""
        reader = StaadReader(self.sample_input)
        combined_df, _, _ = reader.load_worksheets()

        analyzer = StaadAnalyzer(combined_df)
        results = analyzer.run_analysis()

        self.assertIn("136", results)
        n136 = results["136"]
        self.assertEqual(n136.connected_members, ["2018", "2028", "2077", "2502"])
        self.assertEqual(len(n136.connected_member_details), 4)

        # Check Member 2028 details
        m2028_detail = next(d for d in n136.connected_member_details if d.member == "2028")
        self.assertEqual(m2028_detail.source_sheet, "R-L")
        self.assertEqual(m2028_detail.record_count, 10)
        self.assertIn("1257", m2028_detail.available_lcs)
        self.assertIn("1005", m2028_detail.available_lcs)

    def test_05_node_136_member_2028_benchmark(self):
        """Verify explicit user benchmark on Node 136, Member 2028 across all 6 extremes."""
        reader = StaadReader(self.sample_input)
        combined_df, _, _ = reader.load_worksheets()

        analyzer = StaadAnalyzer(combined_df)
        results = analyzer.run_analysis()

        node136 = results["136"]
        self.assertIn("2028", node136.member_governing_forces)
        m2028 = node136.member_governing_forces["2028"]

        # 1. Moment-Z: Max +1526.629 (LC 1257), Min -919.990 (LC 1005) - from screen recording
        self.assertAlmostEqual(m2028.max_mz[0].value, 1526.629, places=3)
        self.assertEqual(m2028.max_mz[0].lc, "1257")
        self.assertEqual(m2028.max_mz[0].source_sheet, "R-L")

        self.assertAlmostEqual(m2028.min_mz[0].value, -919.990, places=3)
        self.assertEqual(m2028.min_mz[0].lc, "1005")
        self.assertEqual(m2028.min_mz[0].source_sheet, "R-L")

        # 2. Shear-Y: Max +230.941 (LC 1257), Min -146.266 (LC 1005) - from screen recording
        self.assertAlmostEqual(m2028.max_fy[0].value, 230.941, places=3)
        self.assertEqual(m2028.max_fy[0].lc, "1257")

        self.assertAlmostEqual(m2028.min_fy[0].value, -146.266, places=3)
        self.assertEqual(m2028.min_fy[0].lc, "1005")

        # 3. Axial Force: Max +16.094 (LC 1030), Min -72.126 (LC 1115) - from screen recording
        self.assertAlmostEqual(m2028.max_fx[0].value, 16.094, places=3)
        self.assertEqual(m2028.max_fx[0].lc, "1030")

        self.assertAlmostEqual(m2028.min_fx[0].value, -72.126, places=3)
        self.assertEqual(m2028.min_fx[0].lc, "1115")

        # 4. Unique governing L/Cs for Member 2028: Exactly 4 load cases (1005, 1030, 1115, 1257)
        m2028_lcs = {rec.lc for rec in m2028.all_governing_records}
        self.assertEqual(m2028_lcs, {"1005", "1030", "1115", "1257"})

    def test_06_governing_extremes_and_ties(self):
        """Test numerical max/min calculations and member-wise tie retention."""
        reader = StaadReader(self.sample_input)
        combined_df, _, _ = reader.load_worksheets()

        analyzer = StaadAnalyzer(combined_df)
        results = analyzer.run_analysis()

        node30 = results["30"]
        # Connected members at Node 30
        self.assertIn("103", node30.member_governing_forces)
        m103 = node30.member_governing_forces["103"]
        self.assertAlmostEqual(m103.max_mz[0].value, 650.000, places=3)
        self.assertEqual(m103.source_sheet, "R-L")

    def test_07_output_excel_all_five_sheets_and_audit(self):
        """Test generation of all 5 worksheets per Master Specification and verify 100% VALID audit traceability."""
        out_path, node_results, val_report, sheets_loaded = process_staad_file(
            input_file=self.sample_input,
            output_file=self.output_file,
        )

        self.assertTrue(os.path.exists(out_path))
        self.assertTrue(val_report.is_valid)
        self.assertEqual(val_report.checks_failed, 0)
        self.assertEqual(len(val_report.audit_mismatches), 0)

        wb = openpyxl.load_workbook(out_path)
        expected_sheets = [
            "Governing Summary",
            "Node-Wise Unique LC",
            "Governing Source Records",
            "Combined Source Data",
            "Validation Report",
        ]
        for s in expected_sheets:
            self.assertIn(s, wb.sheetnames)
        self.assertIn("Node 136", wb.sheetnames)

        # Verify Sheet 1 Headers (Governing Summary)
        ws1 = wb["Governing Summary"]
        headers1 = [ws1.cell(row=3, column=c).value for c in range(1, 9)]
        self.assertEqual(
            headers1,
            [
                "Node ID",
                "Member ID",
                "Force Component",
                "MAX / MIN",
                "Governing Force Value",
                "Corresponding L/C",
                "Source Sheet",
                "Source Row",
            ],
        )

        # Verify Sheet 2 Headers (Node-Wise Unique LC)
        ws2 = wb["Node-Wise Unique LC"]
        headers2 = [ws2.cell(row=4, column=c).value for c in range(1, 5)]
        self.assertEqual(
            headers2,
            [
                "Node ID",
                "Number of Unique L/Cs",
                "Unique L/C Values",
                "Connected Members",
            ],
        )

        # Verify Sheet 3 (Governing Source Records)
        ws3 = wb["Governing Source Records"]
        self.assertGreater(ws3.max_row, 10)

        # Verify Sheet 4 (Combined Source Data) exists and has rows
        ws4 = wb["Combined Source Data"]
        self.assertGreater(ws4.max_row, 10)

        # Verify Sheet 5 (Validation Report)
        ws5 = wb["Validation Report"]
        found_audit_header = False
        audit_start_row = -1
        for r in range(1, ws5.max_row + 1):
            if ws5.cell(row=r, column=1).value == "Node ID" and ws5.cell(row=r, column=2).value == "Member ID":
                found_audit_header = True
                audit_start_row = r + 1
                break
        self.assertTrue(found_audit_header, "Sheet 5 must contain Source Cell Audit Table.")

        # Verify all audit rows are VALID
        for r in range(audit_start_row, ws5.max_row + 1):
            val_status = ws5.cell(row=r, column=9).value
            self.assertEqual(val_status, "VALID", f"Row {r} in Sheet 5 must be VALID!")

    def test_08_gui_error_handling_scope(self):
        """Test that GUI worker exception handling does not trigger closure NameError."""
        from unittest.mock import patch
        from gui.main_window import StaadAppGUI

        with patch("tkinter.messagebox.showerror") as mock_err:
            app = StaadAppGUI()
            app.withdraw()

            # Verify direct error dispatch
            app._on_processing_error("Sample error")
            self.assertTrue(mock_err.called)

            # Verify simulated worker exception scheduled via self.after
            try:
                raise RuntimeError("Simulated background worker exception")
            except Exception as e:
                err_msg = str(e)
                app.after(0, lambda msg=err_msg: app._on_processing_error(msg))

            app.update_idletasks()
            app.update()
            app.destroy()

    def test_09_locked_file_auto_fallback(self):
        """Test that exporter automatically diverts to an incremented filename when destination is locked."""
        from staad_processor.exporter import StaadExporter
        from unittest.mock import MagicMock

        exporter = StaadExporter(
            node_results={},
            combined_df=pd.DataFrame(),
            highlight_map={},
            sheets_loaded=[],
        )

        test_dest = "d:\\Automation\\sample_data\\test_locked_sim.xlsx"
        # Mock save so that the first call on test_dest raises PermissionError, and subsequent call succeeds
        original_save = exporter.wb.save

        save_calls = []

        def mock_save(path):
            save_calls.append(path)
            if path == test_dest:
                raise PermissionError(13, "Permission denied")
            return original_save(path)

        exporter.wb.save = mock_save

        saved_path = exporter.generate_workbook(test_dest)
        self.assertNotEqual(saved_path, test_dest)
        self.assertTrue(saved_path.endswith("_1.xlsx"))
        self.assertTrue(os.path.exists(saved_path))
        os.remove(saved_path)


if __name__ == "__main__":
    unittest.main()
