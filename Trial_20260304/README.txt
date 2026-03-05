
# REACH Visit Compliance Analysis
Author: Zhuoqun Tian
Institution: University of Cincinnati / Cincinnati Children's Hospital Medical Center
Date: October 12, 2025

-----------------------------------------
Purpose
-----------------------------------------
This program (main.py) automates the computation of REACH Visit Compliance for both the Original Cohort
and the New Cohort, in accordance with Dr. Teresa’s specifications.

It reads all visit and cohort data from the Excel workbook “REACH Visit Compliance.xlsx”,
constructs expected visit windows for each subject, determines whether each visit occurred
in-window or out-of-window, applies the COVID-19 adjustment rules,
and outputs summarized compliance tables back into the same Excel file.

-----------------------------------------
Input Excel Structure
-----------------------------------------
The script expects the following sheets to exist:

- 4_Original Cohort:  ID, Site, Screening date, HU initiation date, Off study date
- 4_ New Cohort:      ID, Site, Consent date, HU initiation date, Off study date
- 5_Visit Dates:      Visit records including record_id, site, visit_date, and month columns

-----------------------------------------
Output Sheets
-----------------------------------------
Two new summary sheets will be created in the same Excel file:

- 2_Orig Cohort Sample Output: Summary for the Original cohort
- 3_New Cohort Sample Output:  Summary for the New cohort

Each output table contains:
month | Visits Expected | Visits Completed | Completed % | Visits Completed In Window | In Window %

-----------------------------------------
Calculation Logic Overview
-----------------------------------------
1. Input Processing
Reads all data from the Excel sheets listed above.
Cleans missing or invalid initiation/off-study dates.
Maps all valid visits to their respective participant IDs.

2. Visit Window Generation
2.1 Original Cohort:
Months 0–24: 28-day cycles (±7-day window).
Months 25–48: 31-day cycles (±14-day window).
After Month 48: switches to quarterly visits (every 3 months) using ±14-day windows.

2.2 New Cohort:
Months 0–6: monthly 30-day cycles (±7-day window).
After Month 6: switches to quarterly visits (every 3 months, ±14-day window).

3. Visit Classification
Each visit is classified as one of:
In Window: visit date falls within its expected window.
Out of Window: visit date outside any defined window (counted for nearest prior window).
COVID Period: if window fully within 3/1/2020 – 1/1/2024, automatically marked compliant.

4. Summary Generation
For the monthly section, output is generated for each month.
For the quarterly section, results are aggregated and displayed every 3 months (e.g., 51, 54, 57…).
Each month or quarter is reported with expected/actual visit counts and completion percentages.

5. COVID Adjustment
All windows fully inside March 1, 2020 – January 1, 2024 default to 100% compliance
(both “Completed” and “In Window”).

-----------------------------------------
Run Instructions
-----------------------------------------
1. Open main.py and ensure the address variable points to your Excel file location:
   address = "/Users/.../REACH Visit Compliance tl 2025-1013.xlsx"
2. Run the script (in PyCharm, VSCode, or Terminal):
   pip install pandas openpyxl
   python main.py
3. After execution, check the two output sheets in your Excel file.

-----------------------------------------
Expected Output Example
-----------------------------------------
month | Visits Expected | Visits Completed | Completed % | Visits Completed In Window | In Window %
0     | 603             | 603              | 100         | 603                        | 100
1     | 600             | 595              | 99.1667     | 584                        | 97.333

-----------------------------------------
Notes
-----------------------------------------
--Month 0 is calculated because it exists in the specification sheet.
To exclude it, modify the filter in count_output() before writing results.

--COVID-period windows automatically set all compliance rates to 100%.

--Quarterly aggregation starts at:
Month 48 for Original Cohort
Month 6 for New Cohort

--The script overwrites the two output sheets each time it runs.

-----------------------------------------
Python Setup (optional)
-----------------------------------------
1. Install Python (version 3.13 or later) from https://www.python.org/downloads/
2. Open a terminal and install required packages:
   pip install pandas openpyxl
3. Navigate to the folder containing main.py:
   cd "path/to/your/project"
4. Run the analysis:
   python main.py
