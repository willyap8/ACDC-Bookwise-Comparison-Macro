# ACDC Data Migration — Bookwise Daily Comparison Macro V3.1

> **BookwiseReconcileDailyReportV3** · VBA Excel Add-In · Clinical scheduling environment

A production-ready Excel Add-In that automatically compares a MASTER booking schedule against a daily Bookwise report, flags differences, and organises results into dedicated review sheets — without modifying any source data.

> **Version note:** The current source of truth is **V3.1** (`BookwiseComparisonMacro.bas`). Earlier releases are preserved as Git tags — for example, check out `v1` to retrieve the original version. See [Version History](#12-version-history).

---

## Table of Contents

1. [What Does This Macro Do?](#1-what-does-this-macro-do)
2. [The Files Involved](#2-the-files-involved)
3. [How to Run the Macro](#3-how-to-run-the-macro)
4. [Built-In Safeguards](#4-built-in-safeguards)
5. [The Reconciliation Logic](#5-the-reconciliation-logic)
6. [The Review Sheets in Detail](#6-the-review-sheets-in-detail)
7. [The Reconcile Audit Log](#7-the-reconcile-audit-log)
8. [End-of-Run Summary Message](#8-end-of-run-summary-message)
9. [Installation as an Excel Add-In](#9-installation-as-an-excel-add-in)
10. [What to Do If the Macro Stops](#10-what-to-do-if-the-macro-stops)
11. [Important Reminders](#11-important-reminders)
12. [Version History](#12-version-history)

---

## 1. What Does This Macro Do?

The **ACDC Data Migration — Bookwise Daily Comparison Macro V3.1** is an automated comparison tool built into the MASTER scheduling workbook. Its purpose is to compare the MASTER booking schedule against a freshly generated Bookwise Daily Report, and clearly flag any differences so that staff can review and act on changes quickly and with confidence.

The macro does **not** change any booking data in either the MASTER sheet or the Daily Report. It only applies colour highlights to draw attention to changes, and copies relevant rows into dedicated review sheets for easy actioning.

> **In plain terms:** Think of this macro as an automated checker that asks "What has changed between the master schedule and today's daily report?" It then organises those changes into three categories — Modified, Cancelled, and New — so nothing is missed and every change is traceable.

---

## 2. The Files Involved

The macro works with two Excel files and produces all output within the MASTER workbook itself.

| File / Sheet | Purpose |
|---|---|
| **MASTER Workbook** | The permanent scheduling record. The active sheet must be the tab named exactly `MASTER` (in capitals) when the macro is run. |
| **Daily Report** | A snapshot report generated each day from the Bookwise scheduling system. Selected via file browser each run. Opened read-only — it can never be changed by the macro. |
| **Modified Bookings to Review** | Created inside the MASTER workbook each run. Lists bookings that exist in both reports but where something has changed. |
| **Cancelled Bookings to Review** | Created inside the MASTER workbook each run. Lists bookings present on the MASTER that have disappeared from the Daily Report. |
| **New Bookings to Review** | Created inside the MASTER workbook each run. Lists bookings in the Daily Report that are not yet on the MASTER sheet. |
| **Reconcile Audit Log** | A protected, visible sheet inside the MASTER workbook. Keeps a permanent running record of every macro run, who ran it, and what it found. |

---

## 3. How to Run the Macro

The macro is installed as an Excel Add-In (`.xlam`), making it available whenever Excel is open.

| Step | Action |
|---|---|
| **1** | Ensure the **MASTER** sheet tab is selected before doing anything else. The macro stops immediately if run from any other sheet. |
| **2** | Ensure the MASTER sheet is **not** protected/locked. The macro stops with a clear message if the sheet is protected. |
| **3** | Click the **Reconcile Daily Report** button on the Quick Access Toolbar, or press `Alt + F8`, select `BookwiseReconcileDailyReportV3`, and click **Run**. |
| **4** | A file browser opens. Navigate to and select the Bookwise Daily Report file for today, then click **Open**. |
| **5** | A progress indicator appears in the status bar at the bottom of Excel. Do not interact with the workbook while the macro runs. |
| **6** | A summary message appears on screen showing totals for Modified, Cancelled, and New bookings. |
| **7** | Check the three review sheets for bookings that need attention. Each sheet is rebuilt fresh on every run. |

> **Procedure name:** In V3.1 the macro procedure is named `BookwiseReconcileDailyReportV3`. If you are upgrading from V1 (`ReconcileDailyReport`), your existing Quick Access Toolbar button will point to the old name and must be re-pointed to the new macro (see [Installation](#9-installation-as-an-excel-add-in), Step 5).

---

## 4. Built-In Safeguards

Given the clinical nature of this data, the macro includes a comprehensive set of validation checks before any comparison or output is produced. If any check fails, the macro stops immediately with a clear message. No highlights are applied and no review sheets are written to until all checks pass.

### 4.1 Pre-Run Checks (run in this order)

| # | Safeguard Check | What It Prevents |
|---|---|---|
| 1 | **Correct sheet must be active** | The very first action is to check the active sheet tab is named exactly `MASTER`. If run from any other sheet, the macro exits immediately with no changes made. |
| 2 | **MASTER sheet must be unlocked** | If the MASTER sheet is protected (locked), the macro stops **before** opening any file, so nothing is changed. This fail-fast check avoids partial runs against a sheet that cannot be highlighted. |
| 3 | **Header row 11 must be present** | The macro confirms it can detect column headers in row 11 of the active sheet, guarding against being run on the wrong sheet layout. |
| 4 | **User must select a Daily Report file** | If the file browser is cancelled, the macro exits cleanly. This check runs before any changes are made to MASTER — cancelling leaves everything exactly as it was. |
| 5 | **Search parameters must match** | Rows 2–8 of both reports contain search filter labels (column A) and values (column C), such as the date range covered. These must match exactly. A mismatch suggests the wrong Daily Report has been selected. |
| 6 | **Headers must match exactly (A–T)** | Every header in columns A to T (row 11) must be identical between MASTER and Daily — same spelling, case, and order. The macro stops and names the **first** mismatched column. Only A:T is compared, so only A:T headers need to match. |
| 7 | **MASTER data range formats as a table** | The MASTER range is converted or resized to a structured table covering `A11` down to the last populated column. If merged report-header cells or another table block this, the macro stops with guidance rather than producing an unreliable range. |
| 8 | **Duplicate Booking Numbers in MASTER** | Every Booking Number in MASTER must appear only once. If duplicates are found, the macro stops and lists all offending row numbers. Duplicate entries would make comparison results unreliable. |
| 9 | **No missing Booking Numbers on relevant rows** | Every row containing any booking data must have a Booking Number. Genuinely blank rows are silently skipped. Offending row numbers are listed before stopping. |

### 4.2 Data Integrity Protections

These protections are always active and built into the design of the macro:

- **No booking data is ever changed.** The macro reads cell values for comparison only. It never writes to, overwrites, or deletes any booking information in either file.
- **Only columns A to T are ever compared.** Any data used in the comparison is capped at Column T — nothing beyond T is read, compared, highlighted, or copied. (Comment columns after T may be included in the MASTER *table range* for correct filtering, but are never part of the comparison.)
- **The Daily Report is always opened read-only.** It is physically impossible for the macro to save changes back to the Daily Report file.
- **All highlights are cleared and reapplied fresh on every run.** Highlights from a previous run never carry over incorrectly.
- **Review sheets are fully cleared before each run.** All three review sheets are wiped and rebuilt from scratch every time.
- **Dates are copied as verbatim text.** `Date` and `Date of Birth` values are written into the review sheets as text, preventing Excel from silently reinterpreting a `dd/mm/yyyy` value as a US `mm/dd/yyyy` date on copy.
- **The Audit Log is protected against editing.** It can be read freely but not manually altered, preserving an accurate governance record.

---

## 5. The Reconciliation Logic

Once all safeguards have passed, the macro compares the two reports and categorises every booking into one of four outcomes.

### 5.1 Colour Highlighting on the MASTER Sheet

| Highlight | Meaning |
|---|---|
| 🟡 **Yellow cell** | One or more individual cells in this booking have a different value in the Daily Report. Only the specific changed cells are highlighted — not the whole row. |
| 🔴 **Red row** | The entire booking row (A:T) is highlighted red because the Booking Number no longer appears anywhere in the Daily Report. May indicate cancellation or removal. |
| ⬜ **No highlight** | This booking is unchanged. It appears in the Daily Report with identical values in every compared column. |

---

### 5.2 Modified Bookings

**What triggers this?**
A booking is classified as Modified when its Booking Number exists in both the MASTER and the Daily Report, but at least one piece of information has changed — subject to the special rules below.

When a modified booking is detected, the macro:
- Highlights the specific changed cells **yellow** on the MASTER sheet row
- Copies the entire booking row from the Daily Report (the updated version) into **Modified Bookings to Review**
- Highlights the same changed cells yellow within the review sheet
- Records the exact run date and time in the **Dt/Tm Added by Macro** column

#### Special Rule: Three Columns Are Excluded From Comparison

> Changes to the following three columns are **completely ignored** when deciding whether a booking is modified:
> - **Medicare Number**
> - **Consultant**
> - **Postcode**
>
> If the only differences between the MASTER and Daily Report rows are in these three columns, the booking will **not** be highlighted, will **not** appear in Modified Bookings to Review, and will **not** be counted as modified.

#### Special Rule: Patient Name Capitalisation

If the only difference in the `Patient` column is capitalisation (e.g. `SMITH John` vs `Smith John`), this is **not** treated as a modification. Any other change to the patient name — such as a different name entirely — will still be flagged normally.

#### Special Rule: Start Time Normalisation

Start times are normalised to `hh:mm` **before** comparison. A genuine Excel time value and a text time (e.g. `"10:45"`) that represent the same time are treated as equal, so differing underlying formats do not raise a false "Modified" flag. Start times copied into the review sheet are formatted as a readable time (e.g. `09:30`).

#### Special Rule: Date / Date of Birth Normalisation *(V3.1)*

`Date` and `Date of Birth` values are normalised to a canonical `dd/mm/yyyy` string **before** comparison. This means a genuine Excel date value on one side and a text date (e.g. `"01/09/2026"`) on the other are treated as equal, preventing false "Modified" flags caused only by differing underlying date types or formats. Text dates are parsed explicitly as **day/month/year** (Australian order); any value that cannot be safely and unambiguously parsed falls back to a literal text comparison, so no incorrect date assumptions are ever made.

---

### 5.3 Cancelled Bookings

**What triggers this?**
A booking is classified as Cancelled when its Booking Number is present on the MASTER sheet but cannot be found anywhere in the Daily Report.

When a cancelled booking is detected, the macro:
- Highlights the entire MASTER row (A:T) **red**
- Copies the MASTER version of the booking row into **Cancelled Bookings to Review**
- Records the run date and time in the **Dt/Tm Added by Macro** column

> **Important:** A booking appearing here does not automatically confirm cancellation. It means the Booking Number is absent from today's Daily Report. This could be a date filter issue, a data export issue, or a genuine cancellation. Each entry should be verified manually before any action is taken.

---

### 5.4 New Bookings

**What triggers this?**
A booking is classified as New when its Booking Number appears in the Daily Report but cannot be found anywhere on the MASTER sheet.

When a new booking is detected, the macro:
- Copies the booking row from the Daily Report into **New Bookings to Review**
- Records the run date and time in the **Dt/Tm Added by Macro** column
- Makes no changes to the MASTER sheet, as the booking does not yet exist there

> **Important:** New bookings will need to be manually added to the MASTER schedule by the appropriate staff member if they are to be included in ongoing tracking.

---

### 5.5 Unchanged Bookings

If a booking exists in both reports with identical values in every compared column (taking into account all exclusions, the Patient case rule, and the Start Time / Date / Date of Birth normalisation rules), no action is required. The MASTER row will have no highlight, and the booking will not appear in any of the three review sheets.

---

## 6. The Review Sheets in Detail

Each of the three review sheets follows the same consistent layout:

| Element | Detail |
|---|---|
| **Row 1 — Column headers** | Copied from the MASTER header row for columns A to T, with `Dt/Tm Added by Macro` appended as the final column. |
| **Row 2 onwards — Booking data** | One row per booking requiring review, starting immediately below the headers with no blank rows. |
| **Dt/Tm Added by Macro** | Shows the exact date and time the macro run identified and copied that booking. |
| **Yellow highlights** | Modified Bookings to Review only. Cells that differ from the MASTER are highlighted yellow, matching those on the MASTER sheet. |
| **Date columns as text** | `Date` and `Date of Birth` are written as verbatim text to prevent day/month reinterpretation. |
| **Data source** | Modified: values from the Daily Report. Cancelled: values from the MASTER. New: values from the Daily Report. |

> ⚠️ **These sheets are rebuilt from scratch on every run.** Record or action their contents before running the macro again — previous results are not retained.

---

## 7. The Reconcile Audit Log

Every successful run appends one row to the **Reconcile Audit Log** sheet. This log is permanent and accumulates over time, providing a full traceable history for governance purposes.

| Column | What It Records |
|---|---|
| Run Date/Time | The exact date and time the macro was run. |
| Run By (Username) | The Windows login name of the person who ran the macro. |
| Daily File Path | The full file location of the Daily Report selected for that run. |
| Total Processed | The total number of MASTER booking rows compared. |
| Modified | The count of bookings identified as modified. |
| Cancelled | The count of bookings identified as cancelled. |
| New Bookings | The count of bookings identified as new. |

The Audit Log sheet is visible as a normal sheet tab and can be read freely. Its contents are locked and cannot be manually altered. The macro temporarily unlocks the sheet to write each new row, then immediately re-locks it.

---

## 8. End-of-Run Summary Message

At the end of every successful run, a summary message is displayed showing:

| Summary Line | What It Means |
|---|---|
| Total bookings processed | The number of MASTER booking rows compared against the Daily Report. |
| Modified bookings to review | Bookings where one or more compared fields differ (excluding exempt columns and Patient case-only differences). |
| Cancelled bookings to review | MASTER bookings whose Booking Number could not be found in the Daily Report. |
| New bookings to review | Daily Report bookings whose Booking Number does not exist on the MASTER sheet. |

---

## 9. Installation as an Excel Add-In

### Step 1 — Prepare a blank workbook

Open a new blank Excel workbook. This will become the add-in file and should contain nothing except the macro code.

### Step 2 — Import the macro code

1. Press `Alt + F11` to open the VBA Editor
2. Right-click the project → **Import File…**
3. Select `BookwiseComparisonMacro.bas` from your local copy of this repository
4. Close the VBA Editor

> **Tip:** Importing the `.bas` file brings the whole module in as one unit and avoids the copy/paste errors that can occur when pasting code out of a chat window. To update to a future version later, remove the existing module first (right-click → **Remove**) and import the new `.bas`.

### Step 3 — Save as an Add-In

1. Press `F12` to open Save As
2. In the **Save as type** dropdown, select **Excel Add-In (*.xlam)**
3. Leave the save location as-is (Excel's default add-ins folder)
4. Name the file `ReconcileDailyReport` and click **Save**
5. Close the file

### Step 4 — Install the Add-In

1. Go to **File → Options → Add-ins**
2. Ensure **Manage: Excel Add-ins** is selected and click **Go**
3. Click **Browse**, locate the `.xlam` file, and select it
4. Tick the checkbox next to the add-in name and click **OK**

### Step 5 — Add a Quick Access Toolbar button (recommended)

1. Go to **File → Options → Quick Access Toolbar**
2. In the **Choose commands from** dropdown, select **Macros**
3. Find `BookwiseReconcileDailyReportV3`, click **Add >>**
4. Click **Modify** to set a display name and icon
5. Click **OK**

> **Upgrading from V1?** The procedure name changed from `ReconcileDailyReport` to `BookwiseReconcileDailyReportV3`. Any existing Quick Access Toolbar button bound to the old name will no longer run — remove it and add the new macro as above.

> **Note:** When running from the Quick Access Toolbar, the macro uses `ActiveSheet.Parent` (not `ThisWorkbook`) to correctly reference the MASTER workbook rather than the add-in file itself.

---

## 10. What to Do If the Macro Stops

| Message Received | Recommended Action |
|---|---|
| *This macro is only designed to run on the master sheet...* | Click the **MASTER** sheet tab so it is the active sheet, then run the macro again. |
| *The MASTER sheet is currently protected (locked)...* | Go to the **Review** tab → **Unprotect Sheet**, then run the macro again. |
| *Could not detect column headers in row 11 of this sheet...* | Ensure you are running the macro while the MASTER data sheet is active and that the header row is present in row 11. |
| *No file selected. Macro cancelled.* | No action required. Run the macro again when ready and select a file. |
| *The report search parameters appear to be different...* | Check you selected the correct Daily Report. The date range or filter settings in rows 2–8 do not match the MASTER sheet. |
| *The column headers do not match between the MASTER and Daily report... [column named]* | Open the named column in both files and make the header text in row 11 identical (spelling, case, and order) within columns A to T. |
| *Unable to format the data range as a table covering columns A to [letter]...* | Usually caused by merged cells overlapping the table area, or another table on the sheet. Check the MASTER sheet layout and try again. |
| *No data rows found below the header in the MASTER sheet* / *No booking data found in the MASTER 'Book No.' column* | Confirm the MASTER sheet actually contains booking rows below row 11 and that the `Book No.` column is populated. |
| *Duplicate 'Book No.' values detected... [rows listed]* | Open MASTER and resolve the listed duplicate Booking Numbers. Each must appear only once before the macro can run. |
| *Column 'Book No.' was not found in MASTER / DAILY header row 11 within columns A to T* | Check the column header is spelled exactly as `Book No.` (including the full stop) and falls within columns A to T. |
| *Column 'Patient' was not found in MASTER header row 11 within columns A to T* | Check the `Patient` column header is present and correctly spelled within columns A to T. |
| *Blank 'Book No.' values found in MASTER rows: [rows listed]* | Open MASTER and investigate the listed rows. Add the missing Booking Number or remove the incomplete row. All relevant booking rows must have a Booking Number. |
| *An unexpected error occurred during reconciliation. Error [n]...* | The Daily file has not been saved or modified. Note the error number/description and retry; if it persists, report it with the message shown. |

---

## 11. Important Reminders

> **Always run the macro from the MASTER sheet, and make sure it is unlocked.**
> The macro will stop and alert you if the wrong sheet is active or if the sheet is protected. Check the sheet tab and protection state before clicking Run.

> **Action review sheets before running the macro again.**
> All three review sheets are wiped and rebuilt on every run. Copy or save the content before running again if you need to retain a record of a previous reconciliation.

> **The macro never changes booking data.**
> Only colour highlights are applied to the MASTER sheet. All booking data values remain exactly as they were. The Daily Report is opened read-only and cannot be changed by this macro under any circumstances.

> **Only columns A to T are compared.**
> Any data in columns U onwards is ignored by the comparison, regardless of how many columns exist in either file.

> **Medicare Number, Consultant, and Postcode changes are not flagged.**
> Changes to these three columns are intentionally excluded from the comparison. A booking that differs only in these fields will not appear as modified. This is by design.

> **Date, Date of Birth, and Start Time formatting differences are not flagged.**
> These fields are normalised before comparison so that equivalent values stored in different formats/types are treated as equal. Only genuine differences in the underlying date or time are flagged.

---

## 12. Version History

Versions are tracked in Git. Each named release is retrievable via its tag (`git checkout <tag>`), and the latest version is always the top of the branch.

| Version | Procedure Name | Highlights |
|---|---|---|
| **V3.1** *(current)* | `BookwiseReconcileDailyReportV3` | Date / Date of Birth **comparison** normalised to canonical `dd/mm/yyyy` so a real date value and a text date are treated as equal. Procedure name retained from V3 so the Quick Access Toolbar binding is not broken by the point release. |
| **V3** | `BookwiseReconcileDailyReportV3` | Date / Date of Birth copied to review sheets as verbatim **text** (prevents `dd/mm` ↔ `mm/dd` corruption on copy). Locked-sheet detection, full A:T header-match validation, and dynamic MASTER table conversion/resize. |
| **v1.2** | `BookwiseReconcileDailyReportV3` | Start Time comparison normalised to `hh:mm`. |
| **v1** | `ReconcileDailyReport` | Original release. Highlighting, three review sheets, audit log, core safeguards. |

> To retrieve the original version: `git checkout v1`. To return to the latest: `git checkout <default-branch>`.

---

## Module Structure

```
BookwiseComparisonMacro.bas
│
├── BookwiseReconcileDailyReportV3()   Main sub — orchestrates all steps
│
├── ColLetter()                        Converts a column number to its Excel
│                                       letter(s), used in error messages
│
├── StandardiseFont()                  Sets fonts to Calibri across the A:T
│                                       compare range (neutralises Wingdings)
│
├── SafeText()                          Safe variant-to-string conversion;
│                                       returns "" for Null, Empty, or errors
│
├── NormaliseTime()                     Normalises a time to "hh:mm" for
│                                       comparison (value vs text time)
│
├── NormaliseDate()                     Normalises a date to "dd/mm/yyyy" for
│                                       comparison; parses text as d/m/y
│                                       (Australian), never guesses
│
├── CopyDateAsText()                    Writes a date to a review cell as
│                                       verbatim text (no dd/mm ↔ mm/dd coercion)
│
├── FindHeaderColumn()                  Searches row 11 up to Column T for a
│                                       header name (case-insensitive)
│
├── SetupReviewSheet()                  Creates or fully clears a review sheet,
│                                       writes headers, adds Dt/Tm Added by Macro
│
├── CopyRowWithTimestamp()              Copies a row as values only with timestamp;
│                                       used for Cancelled and New Bookings sheets
│
├── CopyModifiedRow()                   Copies the DAILY version of a modified row,
│                                       applies yellow highlights for changed cells,
│                                       respects excluded columns
│
└── GetOrCreateAuditSheet()             Returns the visible, protected audit log
                                        sheet, creating and formatting it on first run
```

---

*For internal use only. Clinical scheduling environment.*
