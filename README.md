# Master Project: Hybrid AI + ML STAAD.Pro Governing Member Force Automation

A professional Python desktop and command-line application combining **Machine Learning (Scikit-learn TF-IDF semantic interpretation, IsolationForest anomaly detection, and continuous feedback learning)** with **deterministic structural engineering calculations** to process STAAD.Pro exported Excel workbooks (`R-L`, `R-R`, `JB-F`, `JB-B`), construct node-member relationship intelligence, calculate independent member-wise governing forces, eliminate duplicate load cases per node, and provide 100% source-cell audit traceability.

> **CRITICAL ENGINEERING PRINCIPLE**: Engineering calculations strictly use actual STAAD.Pro numbers. Machine learning is used exclusively for worksheet recognition, semantic header interpretation, confidence scoring, and anomaly detection. ML never predicts, fabricates, or modifies structural force values.

---

## 1. Five Output Worksheets (Master Specification)

The generated Excel workbook perfectly reproduces the user's manual workflow demonstrated in the screen recording:

### 1. Original Source Worksheets (`R-L`, `R-R`, `JB-F`, `JB-B`)
- All original sheets, columns, and numeric precision are preserved intact.
- **Visual Highlight**: Every row identified as a governing extreme for any connected member is highlighted in **YELLOW (`#FFFF00`)** directly in the sheet (e.g. Rows 12, 62, 232, 516 in `R-L`).

### 2. Dedicated Node Worksheets (`Node 136`, `Node 149`, etc.)
- A dedicated tab is generated for every unique joint node in the project.
- **Top Summary Block**: Lists all connected framing members, total governing force count, and deduplicated unique governing load cases.
- **Governing Records Table**: Contains all complete source rows for all framing members connected to that node, highlighted in **YELLOW (`#FFFF00`)**, with a `Governing Reasons` column detailing each extreme (e.g. `Moment-Z (MAX: 1526.629); Shear-Y (MAX: 230.941)`).

### 3. Master Summary: `Governing Summary`
- **Columns**: `| Node ID | Member ID | Force Component | MAX / MIN | Governing Force Value | Corresponding L/C | Source Sheet | Source Row |`
- Each `(Node, Member)` combination has independent calculations for $M_z$, $F_y$, $F_x$. Retains all ties.

### 4. Master Deduplication: `Node-Wise Unique LC`
- **Section 1**: Summary Table per Node (`| Node ID | Number of Unique L/Cs | Unique L/C Values | Connected Members |`).
- **Section 2**: Itemized Unique L/C Table (`| Node ID | Unique L/C |`).
- Deduplication is strictly isolated within each node. Common L/Cs across different nodes are preserved.

### 5. Audit & Assurance: `Validation Report`
- **Section 1**: Executive Summary & Data Metrics.
- **Section 2**: Four Mandatory Engineering Verification Checks (Check A, Check B, Check C, Check D).
- **Section 3**: AI Column Recognition Table with confidence scores.
- **Section 4**: Mandatory Source Cell Traceability Audit Table (asserts 100% `VALID` match within $10^{-6}$ tolerance).
- **Section 3**: AI Column Recognition table with ML confidence score and detection basis.
- **Section 4**: Mandatory Source Cell Traceability Audit Table (`Node ID | Member ID | L/C | Force Component | Original Value | Selected Value | Source Sheet | Source Row | Validation Status`).
  - Asserts 100% exact numerical match (`VALID`) between selected forces and original workbook cells.

---

## 2. Benchmark Verification – Node 136 & Member 2028

Verified on authentic STAAD export data:

- **Joint Node**: `136`
- **Framing Members**:
  - `R-L`: Member `2028`
  - `R-R`: Member `2502`
  - `JB-F`: Member `2077`
  - `JB-B`: Member `2018`
- **Member 2028 Governing Extremes (All 6 Verified from Screen Recording)**:
  - **Moment-Z**:
    - Maximum: **`+1526.629 kN-m`** at L/C **`1257`**
    - Minimum: **`-919.990 kN-m`** at L/C **`1005`**
  - **Shear-Y**:
    - Maximum: **`+230.941 kN`** at L/C **`1257`**
    - Minimum: **`-146.266 kN`** at L/C **`1005`**
  - **Axial Force**:
    - Maximum: **`+16.094 kN`** at L/C **`1030`**
    - Minimum: **`-72.126 kN`** at L/C **`1115`**
- **Unique Governing Load Cases for Member 2028**: Exactly 4 load cases (`1005`, `1030`, `1115`, `1257`), because L/C 1257 concurrently governs for Max $M_z$ & Max $F_y$, and L/C 1005 concurrently governs for Min $M_z$ & Min $F_y$.
- **Audit Verification**: 100% matched against source cells in `R-L` with zero deviation.

---

## 3. Machine Learning Architecture

1. **Scikit-learn TF-IDF Semantic Column Classifier** ([ml_engine.py](file:///d:/Automation/staad_processor/ml_engine.py)):
   - Character n-gram (`(2, 5)`) vectorization + cosine similarity against structural category centroids.
   - Built-in engineering safeguards preventing misclassification (e.g. $M_y$ vs $M_z$, $F_z$ vs $F_y$).
   - Confidence scoring (0% to 100%) with explanatory basis.

2. **Scikit-learn IsolationForest Anomaly Detector**:
   - Detects statistical outliers, irregular force distributions, and suspicious spikes in numeric force data.

3. **Continuous Active Feedback Learning Store** (`learning_store.json`):
   - Users can review and confirm detected mappings in the GUI.
   - Confirmed mappings are stored and recognized on future runs with boosted 1.0 (100%) confidence.

---

## 4. Desktop Application (Tkinter GUI)

Features:
- **Browse Input Workbook**: Select any STAAD.Pro Excel workbook.
- **🔍 Preview AI Mappings**: Inspect detected worksheets and column mappings with confidence scores before running.
- **✔ Confirm Mappings**: Confirm mappings and save to continuous learning store.
- **Proactive File Lock Detection**: Prompts if the output Excel file is currently open in Microsoft Excel and offers to auto-save to an incremented filename (`_1.xlsx`).
- **Live Progress & Execution Log**: Real-time progress bar and tagged console output.
- **One-Click Result Opening**: "📄 Open Generated Excel" and "📁 Open Folder".

### Launch GUI:
```powershell
python run_app.py
```
*(or `python main.py --gui`)*

---

## 5. Command-Line Interface (CLI)

Run directly from PowerShell:
```powershell
python main.py path\to\your_staad_file.xlsx -o path\to\output_file.xlsx
```

---

## 6. Automated Unit & Integration Tests

All 9 unit and integration tests pass:
```powershell
python -m unittest tests.test_processor -v
```
```text
test_01_ml_column_classifier ... ok
test_02_ml_learning_store_feedback ... ok
test_03_read_all_four_worksheets_with_ml ... ok
test_04_node_and_connected_member_details ... ok
test_05_node_136_member_2028_benchmark ... ok
test_06_governing_extremes_and_ties ... ok
test_07_output_excel_all_five_sheets_and_audit ... ok
test_08_gui_error_handling_scope ... ok
test_09_locked_file_auto_fallback ... ok

Ran 9 tests in 4.712s
OK
```
