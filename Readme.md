# EIS-DRT-Excel-Analyzer
**Empowering Engineers: Advanced DRT Analysis via Excel VBA—No Installation, No Cloud, No Compromises.**

---

### 1. Introduction: Our Philosophy
In Electrochemical Impedance Spectroscopy (EIS), Distribution of Relaxation Times (DRT) analysis is an incredibly powerful method. It resolves overlapping physical processes—such as charge transfer, diffusion, and contact resistance—into distinct peaks based on their time constants, providing clarity that a standard Nyquist plot cannot.



#### Why "Excel VBA"?
While excellent open-source tools like MATLAB's "DRTtools" or Python libraries exist, engineers in many industrial and research sectors face strict **IT Security Constraints**:

* **No Data Uploads**: Confidential measurement data cannot be uploaded to convenient cloud-based web apps due to trade secret protection.
* **No External Software**: Installing Python environments or executing unknown `.exe` files is often prohibited by corporate policy or requires months of administrative approval.
* **Budget Constraints**: Purchasing expensive commercial analysis software for every team member is not always feasible.

This project is built on three pillars: **No Installation required (Excel is enough)**, **Fully Local Processing (Zero data leak risk)**, and **Transparent VBA Source Code**. It allows engineers to perform high-level analysis immediately on the PC they already have, while staying compliant with strict security policies.

#### Co-creation with Generative AI
The source code for this project was constructed through an extensive dialogue with **Google Gemini**. By combining human domain expertise in electrochemistry with AI's ability to optimize numerical algorithms and ensure cross-platform (Mac/Windows) compatibility, we have implemented sophisticated mathematical modeling within the familiar environment of Excel VBA. This project stands as a testament to how AI can be a powerful "tool" for modern engineering.

---

### 2. Global Settings & Parameters
At the beginning of `DRT_core.bas`, independent variables are defined to control the analysis. Adjust these based on your specific system (Batteries, Fuel Cells, Sensors, etc.).

- **`KK_THRESHOLD`**: Evaluates data consistency based on Kramers-Kronig relations. Automatically excludes outliers exceeding this threshold (%) to prevent artifacts. (Default: 3%)
- **`LAMBDA_SCAN`**: Controls the "smoothness" of the spectrum. The tool automatically scans a wide range ($10^{0}$ to $10^{-10}$) to find the mathematically optimal balance.
- **`CUT_LOW_FREQ`**: Removes artificial spikes that often appear at the low-frequency edge of numerical inversions for a cleaner visualization.

---

### 3. The Internal Processing Workflow
The macro **`ActiveSheetDRT_all`** executes the following scientific steps:

1. **KK-Filter**: Checks if the measurement data is physically sound (linear, stable, and causal). This prevents "false" results derived from noisy measurements.
2. **Tikhonov Regularization**: Calculates the DRT spectrum by solving an "ill-posed" inverse problem while suppressing noise amplification.
3. **L-Curve Method**: Evaluates the trade-off between "model residual" and "solution smoothness" to identify the "elbow point," which represents the most balanced $\lambda$ (Regularization parameter).



4. **Model Reconstruction**: Re-calculates the impedance from the resulting DRT. By overlaying this on the original data, you can visually verify the fit accuracy.

---

### 4. How to Use (Usage)

#### A. Core Analysis (Single Sheet Analysis)
1.  Prepare your data in the active sheet: **Col A: Freq (Hz)**, **Col B: Z'**, **Col C: Z''**.
    *(Note: Z'' should be the positive value typically used in EIS, i.e., -Z'')*
2.  Run the macro **`ActiveSheetDRT_all`**.
3.  A Nyquist comparison plot and the resolved DRT spectrum will be generated near cells A7 and G7.

#### B. Optional Batch Processing (Multiple Files)
1.  In a starting sheet (e.g., `Top`), run **`SelectFiles`** to select multiple **`.z` files** (text format).
2.  Run **`InsertTextCsvFiles`**. Each file will be imported, formatted into the A-C column structure, and placed in its own sheet.
3.  Run **`ProcessAllExtSheets`**. The tool will cycle through all sheets, perform the analysis, and aggregate all results into a single **`Summary_Plots`** sheet.

---

### 5. Requirements & Constraints
* **Confirmed Environment**: Microsoft Excel for macOS.
* **Windows Environment**: The code includes logic for Windows/Mac compatibility (e.g., path separators), but detailed testing on Windows hardware has not been performed yet.
* **File Format**: The batch import function is optimized for **`.z` extension** text files. For other formats, manually paste data into Excel and use the Core Analysis macro.

---

### 6. License
This project is released under the **MIT License**.

* You are free to use, modify, and redistribute the code for commercial or private use.
* The author assumes no liability for any damages arising from the use of this software.
* Corporate use and local modifications are highly encouraged, provided the copyright and license notice are retained.

---
**Development Support**: Developed in collaboration with **Google Gemini**.  
**Author**: [Yuki Kusachi](https://edandc.com) (GitHub: [@ykusachi](https://github.com/ykusachi/Excel-DRT/))
## Detailed Documentation
For a more detailed technical explanation of the DRT algorithm, L-curve optimization, and usage instructions, please visit our dedicated documentation page:

? **[EIS-DRT-Excel-Analyzer Technical Documentation](https://www.edandc.com/Excel-DRT/)**