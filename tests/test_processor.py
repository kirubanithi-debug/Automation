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
    StaadExporter,
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
            include_summary_sheets=True,
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

    def test_10_zero_moment_pinned_connection_logic(self):
        """Test pinned member logic: when Moment-Z is 0.0, govern by Shear-Y & Axial Force, mark yellow."""
        # Replicates Screen Recording 2026-11-14 131115.mp4:
        # Node 1457, Member 1072 in JB-F with Moment-Z = 0.0
        # Shear-Y values: 0.146 (Min), 0.195 (Mid), 0.243 (Max)
        # Chained Axial Force selection yields exactly Row 18 (L/C 1008) and Row 220 (L/C 1109)
        zero_mz_data = pd.DataFrame({
            "_clean_node": ["1457", "1457", "1457", "1457", "1457"],
            "_clean_member": ["1072", "1072", "1072", "1072", "1072"],
            "_clean_lc": ["1000", "1001", "1008", "1109", "1200"],
            "_clean_moment_z": [0.0, 0.0, 0.0, 0.0, 0.0],
            "_clean_shear_y": [0.243, 0.146, 0.146, 0.243, 0.195],  # Min: 0.146, Max: 0.243
            "_clean_axial": [-9.969, -41.524, 49.004, -43.702, 20.0],  # Cand Min: -43.702, Cand Max: 49.004
            "_raw_val_moment_z": [0.0, 0.0, 0.0, 0.0, 0.0],
            "_raw_val_shear_y": [0.243, 0.146, 0.146, 0.243, 0.195],
            "_raw_val_axial": [-9.969, -41.524, 49.004, -43.702, 20.0],
            "_source_sheet": ["JB-F", "JB-F", "JB-F", "JB-F", "JB-F"],
            "_original_sheet_row": [2, 4, 18, 220, 300],
        })

        analyzer = StaadAnalyzer(zero_mz_data)
        results = analyzer.run_analysis()
        self.assertIn("1457", results)

        node1457 = results["1457"]
        m1072 = node1457.member_governing_forces["1072"]
        self.assertTrue(m1072.is_zero_moment, "Member 1072 must be recognized as zero moment / pinned.")
        self.assertEqual(m1072.beam_type, "Jack Beam (Front)")
        self.assertEqual(m1072.connection_type, "Pinned Connection")

        # Moment-Z extremes must NOT capture any zero rows
        self.assertEqual(len(m1072.max_mz), 0)
        self.assertEqual(len(m1072.min_mz), 0)

        # Shear-Y extremes: Max = 0.243 (L/C 1109), Min = 0.146 (L/C 1008)
        self.assertEqual(len(m1072.max_fy), 1)
        self.assertEqual(m1072.max_fy[0].value, 0.243)
        self.assertEqual(m1072.max_fy[0].lc, "1109")
        self.assertEqual(len(m1072.min_fy), 1)
        self.assertEqual(m1072.min_fy[0].value, 0.146)
        self.assertEqual(m1072.min_fy[0].lc, "1008")

        # Axial Force extremes from candidate pool: Max = 49.004 (L/C 1008), Min = -43.702 (L/C 1109)
        self.assertEqual(len(m1072.max_fx), 1)
        self.assertEqual(m1072.max_fx[0].value, 49.004)
        self.assertEqual(m1072.max_fx[0].lc, "1008")
        self.assertEqual(len(m1072.min_fx), 1)
        self.assertEqual(m1072.min_fx[0].value, -43.702)
        self.assertEqual(m1072.min_fx[0].lc, "1109")

        # Governing rows in highlight_map must only be rows 18 and 220 (exactly 2 rows, matching video!)
        highlighted_rows = {k[1] for k in analyzer.highlight_map.keys()}
        self.assertEqual(highlighted_rows, {18, 220})
        self.assertNotIn(2, highlighted_rows)
        self.assertNotIn(4, highlighted_rows)
        self.assertNotIn(300, highlighted_rows)

        # Node unique L/Cs: exactly 1008 and 1109 (exactly 2 L/Cs)
        self.assertEqual(node1457.unique_governing_lcs, ["1008", "1109"])

        # Validate with StaadValidator
        validator = StaadValidator(
            node_results=results,
            combined_df=zero_mz_data,
            sheets_loaded=["JB-F"],
            invalid_force_records=[],
            expected_sheets=["JB-F"],
        )
        report = validator.validate()
        self.assertTrue(report.is_valid, f"Validation failed: {report.errors}")
        self.assertEqual(report.checks_failed, 0)

    def test_11_dynamic_sheet_discovery_jb_and_tb(self):
        """
        Test dynamic sheet discovery when an Excel workbook uses 'JB' and 'TB'
        instead of 'JB-F' and 'JB-B', along with non-data sheets like 'Notes'.
        """
        import tempfile
        import openpyxl

        wb = openpyxl.Workbook()
        # Sheet 1: R-L
        ws_rl = wb.active
        ws_rl.title = "R-L"
        ws_rl.append(["Beam", "L/C", "Node", "Axial Force (kN)", "Shear-Y (kN)", "Moment-Z (kN-m)"])
        ws_rl.append([101, 1, 50, 10.0, 50.0, 120.0])
        ws_rl.append([101, 2, 50, -20.0, -40.0, -110.0])

        # Sheet 2: R-R
        ws_rr = wb.create_sheet("R-R")
        ws_rr.append(["Beam", "L/C", "Node", "Axial Force (kN)", "Shear-Y (kN)", "Moment-Z (kN-m)"])
        ws_rr.append([102, 1, 50, 12.0, 45.0, 95.0])
        ws_rr.append([102, 2, 50, -18.0, -35.0, -90.0])

        # Sheet 3: JB (Jack Beam, Moment-Z = 0.0)
        ws_jb = wb.create_sheet("JB")
        ws_jb.append(["Beam", "L/C", "Node", "Axial Force (kN)", "Shear-Y (kN)", "Moment-Z (kN-m)"])
        ws_jb.append([103, 1, 50, 5.0, 0.250, 0.0])
        ws_jb.append([103, 2, 50, -30.0, 0.250, 0.0])
        ws_jb.append([103, 3, 50, 40.0, 0.150, 0.0])

        # Sheet 4: TB (Tie Beam, Moment-Z = 0.0)
        ws_tb = wb.create_sheet("TB")
        ws_tb.append(["Beam", "L/C", "Node", "Axial Force (kN)", "Shear-Y (kN)", "Moment-Z (kN-m)"])
        ws_tb.append([104, 1, 50, 100.0, 0.100, 0.0])
        ws_tb.append([104, 2, 50, -80.0, 0.200, 0.0])

        # Sheet 5: Notes (Non-data sheet without member tables)
        ws_notes = wb.create_sheet("Notes")
        ws_notes.append(["Project Name: Warehouse Phase 2"])
        ws_notes.append(["Engineer: Structural Team"])

        temp_dir = tempfile.mkdtemp()
        temp_input = os.path.join(temp_dir, "test_dynamic_jb_tb.xlsx")
        temp_output = os.path.join(temp_dir, "test_dynamic_jb_tb_out.xlsx")
        wb.save(temp_input)

        try:
            reader = StaadReader(temp_input)
            combined_df, per_sheet_dfs, invalid = reader.load_worksheets()

            # Dynamic discovery must find R-L, R-R, JB, and TB, safely ignoring Notes
            self.assertEqual(set(reader.sheets_found), {"R-L", "R-R", "JB", "TB"})
            self.assertNotIn("Notes", reader.sheets_found)

            analyzer = StaadAnalyzer(combined_df)
            node_results = analyzer.run_analysis()
            self.assertIn("50", node_results)

            n50 = node_results["50"]
            self.assertEqual(set(n50.connected_members), {"101", "102", "103", "104"})

            # Verify JB and TB recognized as zero-moment pinned connections
            m101 = n50.member_governing_forces["101"]
            self.assertFalse(m101.is_zero_moment)
            self.assertEqual(m101.beam_type, "Rafter (Left)")
            self.assertEqual(m101.connection_type, "Moment Connection")

            m102 = n50.member_governing_forces["102"]
            self.assertFalse(m102.is_zero_moment)
            self.assertEqual(m102.beam_type, "Rafter (Right)")
            self.assertEqual(m102.connection_type, "Moment Connection")

            m103 = n50.member_governing_forces["103"]
            self.assertTrue(m103.is_zero_moment)
            self.assertEqual(m103.beam_type, "Jack Beam")
            self.assertEqual(m103.connection_type, "Pinned Connection")

            m104 = n50.member_governing_forces["104"]
            self.assertTrue(m104.is_zero_moment)
            self.assertEqual(m104.beam_type, "Tie Beam")
            self.assertEqual(m104.connection_type, "Pinned / Axial Strut")

            # Validate with StaadValidator (dynamic sheets_loaded)
            validator = StaadValidator(
                node_results=node_results,
                combined_df=combined_df,
                sheets_loaded=reader.sheets_found,
                invalid_force_records=[],
            )
            report = validator.validate()
            self.assertTrue(report.is_valid, f"Validation failed: {report.errors}")
            self.assertEqual(report.checks_failed, 0)

            # Exporter: output must include source sheets with highlights and dedicated joint tab
            exporter = StaadExporter(
                node_results=node_results,
                combined_df=combined_df,
                highlight_map=analyzer.highlight_map,
                sheets_loaded=reader.sheets_found,
                per_sheet_col_mappings=reader.per_sheet_col_mappings,
                per_sheet_confidences=reader.per_sheet_confidences,
                per_sheet_bases=reader.per_sheet_bases,
                input_file_path=temp_input,
                include_summary_sheets=False,
            )
            saved_file = exporter.generate_workbook(temp_output)
            self.assertTrue(os.path.exists(saved_file))

            out_wb = openpyxl.load_workbook(saved_file)
            self.assertIn("R-L", out_wb.sheetnames)
            self.assertIn("R-R", out_wb.sheetnames)
            self.assertIn("JB", out_wb.sheetnames)
            self.assertIn("TB", out_wb.sheetnames)
            self.assertIn("Node 50", out_wb.sheetnames)

            # Verify yellow highlight exists in JB and TB
            ws_out_jb = out_wb["JB"]
            has_yellow_jb = any(
                str(ws_out_jb.cell(row=r, column=1).fill.start_color.rgb) in ("00FFFF00", "FFFF00")
                for r in range(2, ws_out_jb.max_row + 1)
            )
            self.assertTrue(has_yellow_jb, "Governing rows in 'JB' must be highlighted in yellow.")

            ws_out_tb = out_wb["TB"]
            has_yellow_tb = any(
                str(ws_out_tb.cell(row=r, column=1).fill.start_color.rgb) in ("00FFFF00", "FFFF00")
                for r in range(2, ws_out_tb.max_row + 1)
            )
            self.assertTrue(has_yellow_tb, "Governing rows in 'TB' must be highlighted in yellow.")

            # Verify Node 50 tab includes Beam Type and Connection columns
            ws_node50 = out_wb["Node 50"]
            node50_headers = [ws_node50.cell(row=4, column=c).value for c in range(1, ws_node50.max_column + 1)]
            self.assertIn("Beam Type", node50_headers)
            self.assertIn("Connection", node50_headers)
            self.assertIn("Governing Reasons", node50_headers)

            # Subtitle check
            sub_val = str(ws_node50.cell(row=2, column=1).value)
            self.assertIn("Rafter (Left)", sub_val)
            self.assertIn("Jack Beam", sub_val)
            self.assertIn("Tie Beam", sub_val)

        finally:
            if os.path.exists(temp_input):
                os.remove(temp_input)
            if os.path.exists(temp_output):
                os.remove(temp_output)
            if os.path.exists(temp_dir):
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
