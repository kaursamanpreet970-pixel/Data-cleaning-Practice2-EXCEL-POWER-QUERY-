# Employee Attendance Data Cleaning with Excel Power Query

## Brief One Line Summary

Converted 1,050 pipe-delimited, inconsistently formatted employee attendance records into a clean, correctly typed table using Excel Power Query.

---

## Overview

This project demonstrates a text-parsing and standardisation workflow in Microsoft Excel, using Power Query as the primary tool. The raw data is an attendance log exported as a single text column, with nine fields separated by a pipe (`|`) character. Beyond splitting the fields, the data also needed its text casing standardised and its dates and times converted into true date and time values.

The workbook contains two sheets:

| Sheet | Purpose |
|---|---|
| `raw data` | The original single-column log (`Raw Data`), plus a small `Final Outcome` example showing the target layout |
| `cleaned dataset` | The final table produced by Power Query: 1,050 rows, 9 columns |

---

## Problem Statement

Attendance logs exported from HR or biometric systems are rarely analysis-ready. Here, each record was stored as one long text string inside a single cell, so it could not be filtered by department, grouped by date, or used to calculate working hours.

The specific problems in the raw data were:

1. **Everything in one column.** Employee ID, date, title, first name, last name, department, designation, in time, and out time were all combined in one cell, separated by `|`.
2. **Headers stored as data.** The column names appeared as the first row of the data (`Emp ID | Date | Title | ...`) rather than as real headers.
3. **Inconsistent text casing.** The same value appeared in several forms. For example, the department `finance` appears as `finance`, `Finance`, and `FINANCE`, and the title `mr` appears as `mr` and `MR`. Across the dataset, 16 spellings of department collapse to just 6 real departments, and 100 spellings of designation collapse to 45.
4. **Names in mixed case.** Roughly 440 first-name and 450 last-name entries were not in a uniform case (for example `Anita VERMA`, `Rahul sharma`).
5. **Dates stored as text.** Dates such as `1-12-25` and `10-12-25` were text strings in day-month-year format with no leading zeros, so they could not be sorted or grouped chronologically.
6. **Times stored as text.** In and out times such as `9:05` and `18:10` were text, which prevents calculations such as hours worked.
7. **Spacing around delimiters.** Spaces around the `|` separators meant that fields could not be matched or compared reliably until they were trimmed.

---

## Dataset

**Source file:** `data_cleaning_practice_2.xlsx`
**Size:** 1,050 records, 9 fields

### Raw data structure

Each raw row follows this pattern:

```
Emp ID | Date | Title | First Name | Last Name | Department | Designation | In Time | Out Time
```

Examples of raw rows:

```
E002 | 2-12-25 | MS | Anita | VERMA | HR | HR Manager | 9:00 | 17:50
E001 | 4-12-25 | mr | rahul | sharma | SALES | sales executive | 9:20 | 18:15
E001 | 10-12-25 | mr | rahul | sharma | sales | sales executive | 9:00 | 18:00
```

### Fields

| Column | Description | Final type |
|---|---|---|
| `Emp ID` | Employee identifier (E001 to E050) | Text |
| `Date` | Attendance date | Date |
| `Title` | Mr / Ms | Text |
| `First Name` | Employee first name | Text |
| `Last Name` | Employee last name | Text |
| `Department` | Department name | Text |
| `Designation` | Job role | Text |
| `In Time` | Clock-in time | Time |
| `Out Time` | Clock-out time | Time |

### Dataset profile

| Measure | Value |
|---|---|
| Total records | 1,050 |
| Distinct employees | 50 |
| Distinct dates | 19 (1 December 2025 to 21 December 2025) |
| Departments (after standardisation) | 6 |
| Distinct designations (after standardisation) | 45 |
| Records with missing values | 0 |

---

## Tools and Technologies

- **Microsoft Excel** (Microsoft 365 / Excel 2016 or later)
- **Power Query** (Get & Transform Data) for all transformations
- **Power Query M language** (steps generated through the editor interface)

---

## Methods

All cleaning was done in a single Power Query, with every step recorded in the Applied Steps pane so the process is transparent and repeatable.

| Step | Power Query action | Purpose |
|---|---|---|
| 1 | Load the `raw data` table and set the column type to Text | Establish a consistent starting point |
| 2 | **Trim** | Remove leading and trailing spaces from each raw string |
| 3 | **Split Column by Delimiter** (`|`) into 9 columns | Separate the single text string into individual fields |
| 4 | Change all 9 columns to Text | Prepare the columns before headers are promoted |
| 5 | **Use First Row as Headers** | Promote the embedded header row (`Emp ID`, `Date`, ...) to real column names |
| 6 | **Change Type** per column | `Date` to Date, `In Time` and `Out Time` to Time, all other columns to Text |
| 7 | **Capitalize Each Word** on `Title` | Standardise `mr` / `MR` to `Mr`, and `ms` / `MS` to `Ms` |
| 8 | **Lowercase** on `First Name`, `Last Name`, `Department`, `Designation` | Collapse the casing variants into one consistent form |

**Outcome of the casing steps:** departments reduced from 16 variants to 6; designations from 100 variants to 45; titles from 5 variants to 2.

---

## Key Insights

**What the cleaned data shows**

- The log covers 50 employees across 6 departments, with records for 19 distinct dates between 1 and 21 December 2025.
- Department headcount in the log is led by Finance (252 records), Sales (231), HR (210), and IT (210), followed by Operations (105) and Logistics (42).
- Titles split 588 Mr to 462 Ms (about 56% to 44%).
- Every in time and out time converted successfully to a valid time value, and no record has an out time earlier than its in time. Hours worked now average about 8.9 per day, ranging from 1 to 9.5 hours; the shortest shifts are worth confirming with the data owner.

**Data quality observations and next steps**

These are points I identified on review and would address in a further iteration:

- **Repeated records.** There are 98 exact duplicate rows and 100 repeated Employee ID and Date combinations. A small number of these differ in other fields, so they should be reviewed before deciding which entry to keep. A `Remove Duplicates` step would handle the exact repeats.
- **Leftover spaces in the output.** The Trim step was applied before the split, so individual fields (and the column headers) still carry leading or trailing spaces, for example `E001 ` and ` finance `. Applying Trim to every column after the split would fix this and make lookups and grouping more reliable.
- **Names are lowercase rather than proper case.** First and last names were lowercased to remove inconsistency. Converting them to Capitalize Each Word instead would give a more presentable result for reports and dashboards.
- **Duplicate designation spellings.** Designations are now consistent in casing, but shortened or similar labels (for example `key account exec`) may still need a mapping table if they should match other titles.

---

## Dashboard/Model/Output

### Before and after

| Raw (`raw data`) | Cleaned (`cleaned dataset`) |
|---|---|
| `E002 \| 2-12-25 \| MS \| Anita \| VERMA \| HR \| HR Manager \| 9:00 \| 17:50` | E002 / 02-12-2025 / Ms / anita / verma / hr / hr manager / 09:00 / 17:50 |
| `E001 \| 4-12-25 \| mr \| rahul \| sharma \| SALES \| sales executive \| 9:20 \| 18:15` | E001 / 04-12-2025 / Mr / rahul / sharma / sales / sales executive / 09:20 / 18:15 |
| `E003 \| 3-12-25 \| MR \| Rohit \| GUPTA \| IT \| data analyst \| 9:10 \| 18:20` | E003 / 03-12-2025 / Mr / rohit / gupta / it / data analyst / 09:10 / 18:20 |

### Output structure

| Emp ID | Date | Title | First Name | Last Name | Department | Designation | In Time | Out Time |
|---|---|---|---|---|---|---|---|---|
| E001 | 01-12-2025 | Mr | rahul | sharma | sales | sales executive | 09:05 | 18:10 |
| E002 | 01-12-2025 | Ms | anita | verma | hr | hr manager | 08:55 | 17:45 |
| E003 | 01-12-2025 | Mr | rohit | gupta | it | data analyst | 09:00 | 11:00 |
| E004 | 01-12-2025 | Mr | sumit | kumar | finance | accountant | 09:10 | 18:00 |
| E005 | 01-12-2025 | Ms | pooja | sharma | sales | sales coordinator | 09:15 | 18:00 |

The cleaned table is ready for a PivotTable, or for loading into Power BI or Tableau to build attendance and punctuality dashboards.

---

## Results & Conclusion

- **Outcome:** 1,050 unstructured text records were converted into a nine-column table with correct data types: text, date, and time.
- **Standardisation:** inconsistent casing was resolved across title, name, department, and designation fields, reducing 16 department variants to 6 and 100 designation variants to 45.
- **Approach:** the work was completed entirely with Power Query's built-in transformations (trim, split by delimiter, promote headers, change type, and text case transformations), with no formulas or manual editing. The query can be refreshed whenever new data is added.
- **Skills demonstrated:** delimiter-based parsing, header promotion, data type conversion for dates and times, text standardisation, and documenting data quality issues.
- **Next steps:** remove repeated records, trim every column after splitting, switch names to proper case, and connect the cleaned table to a Power BI or Tableau attendance dashboard.

---

## Connect with Me

**SAMANPREET KAUR**
BA Student | Aspiring Data Analyst, Business Analyst, and BI Developer

- LinkedIn: [https://www.linkedin.com/in/kaursamanpreet/](https://www.linkedin.com/in/kaursamanpreet/)
- Email: [kaursamanpreet970@outlook.com](mailto:kaursamanpreet970@outlook.com)
- GitHub: [https://github.com/kaursamanpreet970-pixel](https://github.com/kaursamanpreet970-pixel)
