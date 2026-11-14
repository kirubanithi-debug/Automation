# Master Project: Hybrid AI + ML STAAD.Pro Governing Force Workflow & Architecture

## 1. Overview & Objective
In connection design and member capacity checks, structural engineers receive STAAD.Pro export workbooks containing member end forces segregated into structural member groups. Depending on the model, sheet names vary:
1. **R-L** (Rafter - Left)
2. **R-R** (Rafter - Right)
3. **JB-F** (Jack Beam - Front) or **JB** (Jack Beam)
4. **JB-B** (Jack Beam - Back) or **TB** (Tie Beam)

The pipeline dynamically discovers and processes all valid structural member sheets present in the workbook (`R-L`, `R-R`, `JB-F`, `JB-B`, `JB`, `TB`, etc.).

At each structural joint (Node), members from different groups converge (e.g. Rafters and Jack Beams meeting at the same node).

### The Master Workflow:
1. **Node and Member Identification**:
   - Identify all unique nodes across all worksheets.
   - For every node, identify all connected members and their source worksheets.
   - Group records by **(Node ID, Member ID)**.
   - Each member is processed independently for governing forces.
2. **Member-Wise Governing Force Calculation**:
   - For each Node + Member combination, calculate the six governing values:
     - Numerical Max and Min **Moment-Z ($M_z$)** and corresponding L/Cs
     - Numerical Max and Min **Shear-Y ($F_y$)** and corresponding L/Cs
     - Numerical Max and Min **Axial Force ($F_x$)** and corresponding L/Cs
   - Preserve algebraic signs, preserve float precision, and retain all ties.
3. **Node-Wise Unique L/C Extraction**:
   - Collect the governing L/C numbers from all connected members of that node.
   - Deduplicate the load case numbers within each node.
4. **Automated ML Validation & Traceability Audit**:
   - Verify 100% of selected governing forces match the source Excel cell.
   - Run Scikit-learn `IsolationForest` anomaly detection for statistical quality assurance.

---

## 2. Selection Rules & Mathematical Definitions
- **Numerical Extremes**:
  - Maximum means the highest algebraic number on the real line: $\max(x)$.
  - Minimum means the lowest algebraic number on the real line: $\min(x)$.
  - Absolute values ($|x|$) are strictly **not** used.
  - Sign conventions ($+$ for tension, $-$ for compression, or standard STAAD orientations) are preserved without alteration.
- **Member Count Definition**:
  - A member having records across multiple load cases is counted as **one member**.
  - Connected member count reflects the number of distinct Member IDs framing into that node.
- **Tie Handling**:
  - If two or more load cases produce the exact same governing value, **all tied records are retained**.
- **Source Worksheet Traceability**:
  - Every governing record tracks whether it originated from `R-L`, `R-R`, `JB-F`, or `JB-B`, down to the exact 1-based source row number.

---

## 3. Five-Worksheet Output Specifications

### 1. Original Source Worksheets (`R-L`, `R-R`, `JB-F`, `JB-B`, `JB`, `TB`)
- All original sheets, columns, and numeric precision are preserved intact. Dynamically discovers varying sheet patterns (`JB-F`/`JB-B` or `JB`/`TB`).
- **Visual Highlight**: Every row identified as a governing extreme is highlighted in **YELLOW (`#FFFF00`)** directly in the sheet (reproducing the manual workflow from the screen recording).

### 2. Dedicated Node Worksheets (`Node 136`, `Node 149`, etc.)
- A dedicated tab is generated for each unique joint node.
- **Header**: Summarizes connected framing members with automatically identified beam types (e.g. `101 (Rafter (Left)), 103 (Jack Beam), 104 (Tie Beam)`) and unique governing load cases.
- **Table**: Gathers all governing rows for all members connected to that node, highlighted in **YELLOW (`#FFFF00`)**, with columns for original forces, **`Beam Type`**, **`Connection`**, and a **`Governing Reasons`** column (e.g. `Moment-Z (MAX: 1526.629); Shear-Y (MAX: 230.941)`).

### 3. Master Summary: `Governing Summary`
`| Node ID | Member ID | Force Component | MAX / MIN | Governing Force Value | Corresponding L/C | Source Sheet | Source Row |`
- Each Node + Member combination has independent calculations for $M_z$, $F_y$, and $F_x$.
- Retains all tied extremes.

### 4. Master Deduplication: `Node-Wise Unique LC`
- **Section 1**: Summary Table per Node (`| Node ID | Number of Unique L/Cs | Unique L/C Values | Connected Members |`).
- **Section 2**: Itemized Unique L/C Table (`| Node ID | Unique L/C |`).
- Deduplicated governing load cases per node across all connected members. Deduplication is strictly isolated to each node.

### Sheet 5 – Validation Report
- **Section 1**: Executive Summary & Data Metrics.
- **Section 2**: Four Mandatory Engineering Verification Checks (Check A, Check B, Check C, Check D).
- **Section 3**: AI Column Recognition Table with ML confidence score and detection basis.
- **Section 4**: Mandatory Source Cell Traceability Audit Table (`Node ID | Member ID | L/C | Force Component | Original Value | Selected Value | Source Sheet | Source Row | Validation Status`).
  - Asserts 100% exact numerical match (`VALID`) between selected forces and original workbook cells.

---

## 4. Benchmark Reference – Node 136, Member 2028

- **Node 136 Connected Framing Members**:
  - `R-L`: Member `2028`
  - `R-R`: Member `2502`
  - `JB-F`: Member `2077`
  - `JB-B`: Member `2018`
- **Member 2028 Verified Extremes (from Screen Recording Recording 2026-11-13 162544.mp4)**:
  - **Moment-Z**: Maximum = `+1526.629 kN-m` (L/C 1257), Minimum = `-919.990 kN-m` (L/C 1005)
  - **Shear-Y**: Maximum = `+230.941 kN` (L/C 1257), Minimum = `-146.266 kN` (L/C 1005)
  - **Axial Force**: Maximum = `+16.094 kN` (L/C 1030), Minimum = `-72.126 kN` (L/C 1115)
  - **Unique Governing L/Cs for Member 2028**: `[1005, 1030, 1115, 1257]` (4 unique L/Cs). L/C 1257 governs for both Max $M_z$ and Max $F_y$; L/C 1005 governs for both Min $M_z$ and Min $F_y$.

---

## 5. Benchmark Reference – Node 1457, Member 1072 (Pinned Connection Mz = 0.0)
*From Screen Recording 2026-11-14 131115.mp4*

- **Worksheet**: `JB-F`
- **Node**: `1457`
- **Member**: `1072`
- **Moment-Z**: All values are `0.000 kN-m` (pinned / moment-released joint).
- **Chained Selection Logic**:
  1. Filter `Shear-Y (kN)` extremes: Min = `0.146 kN`, Max = `0.243 kN` (165 candidate records).
  2. Filter `Axial Force (kN)` extremes from the candidate records: Min = `-43.702 kN` (Row 220), Max = `49.004 kN` (Row 18).
- **Selected Governing Records (Exactly 2 Rows)**:
  - **Row 18 (L/C 1008)**: Shear-Y = `0.146 kN` (Min), Axial Force = `+49.004 kN` (Max)
  - **Row 220 (L/C 1109)**: Shear-Y = `0.243 kN` (Max), Axial Force = `-43.702 kN` (Min)
- **Visual Highlight**: Rows 18 and 220 in worksheet `JB-F` are highlighted in **YELLOW (`#FFFF00`)**.
- **Unique Governing L/Cs**: `["1008", "1109"]`.
