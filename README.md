# Advanced Excel — Complete Master Notes (Hinglish & Detailed Reference)
**Prepared for:** Lokesh Ahire (Lucky)  
**Topics Covered:** Basics, Cell References ($A$1/$A1/A$1), Logical (IF/IFS/AND/OR/SWITCH), Lookups (VLOOKUP, HLOOKUP, XLOOKUP, INDEX+MATCH), Text Functions, Date & Time, Math & Stats, Data Validation & Cascading Dropdowns, Conditional Formatting, PivotTables & Slicers, Visual Charts & Executive Dashboards, What-If Analysis & Solver, Dynamic Array Spill Formulas (Excel 365), Excel Tables (Ctrl+T), Named Ranges, Power Query ETL, Power Pivot & DAX, Macros & VBA Automation, Analysis ToolPak Add-in, Error Cheat Sheet, Job Best Practices & 20 Real-World Practical Scenarios.

---

## 1. Excel Basics & Shortcuts Quick Recap
- **Workbook:** Complete Excel file (`.xlsx`)
- **Worksheet:** Individual tabs/sheets inside a workbook
- **Cell:** Intersection of row & column (e.g., A1, B2)
- **Range:** Group of multiple cells (e.g., A1:B10)

### Key Shortcuts Table
| Shortcut | Action / Kaam |
|---|---|
| `Ctrl + C / V` | Copy / Paste |
| `Ctrl + Z / Y` | Undo / Redo |
| `Ctrl + Arrow Keys` | Jump to edge of data region |
| `Ctrl + Shift + Arrow` | Select region to edge |
| `Ctrl + Shift + L` | Toggle AutoFilter On/Off |
| `Ctrl + T` | Convert range to Excel Table |
| `Alt + =` | Instant AutoSum |
| `Ctrl + 1` | Format Cells dialog |
| `F4` | Repeat last action / Toggle absolute reference ($) |
| `F2` | Edit active cell |
| `Ctrl + E` | Flash Fill (auto pattern extraction) |
| `Ctrl + Shift + Enter` | Legacy array formula (CSE) |

---

## 2. Cell References (Relative, Absolute, Mixed)
Excel me 3 type ke references hote hain — advance formulas ke liye ye samajhna critical hai:
1. **Relative Reference (`A1`):** Formula copy/drag karne par cell address automatically change hota hai.
2. **Absolute Reference (`$A$1`):** Dollar sign ($) lagane se cell location completely FIX (Lock) ho jati hai.
3. **Mixed Reference (`$A1` or `A$1`):** Sirf row (`A$1`) ya sirf column (`$A1`) fix hota hai.

> **Practical Example:** Fixed tax rate cell (`$B$1`) ko sabhi product rows se multiply karte waqt tax rate cell ko `$B$1` bana do taaki drag karne par wo fix rahe. Press **F4** to cycle reference types!

---

## 3. Logical Functions
- `IF(condition, value_if_true, value_if_false)`: Basic decision making
- `IFS(cond1, val1, cond2, val2...)`: Multiple conditions ek sath (Nested IF ka replacement)
- `AND(cond1, cond2...)`: Sab conditions true hone chahiye
- `OR(cond1, cond2...)`: Kam se kam ek condition true
- `NOT(condition)`: Condition reverse करता है
- `IFERROR(formula, value_if_error)`: Custom message on any error
- `IFNA(formula, value_if_na)`: Specifically handles #N/A lookup errors
- `SWITCH(value, case1, res1, default)`: Direct value match list

---

## 4. Lookup & Reference Functions
- **VLOOKUP:** `=VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)` (Searches 1st column, looks right only. Always use 0 for exact match).
- **HLOOKUP:** Horizontal row lookup.
- **XLOOKUP (Excel 365/2021+):** `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` (Can look LEFT or RIGHT, exact match by default, built-in fallback error text).
- **INDEX + MATCH Combo:** `=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))` (Fastest lookup for large datasets, multi-directional).
- **OFFSET / INDIRECT / CHOOSE:** Dynamic ranges, text to reference conversion, index selection.

---

## 5. Text Functions & Data Cleaning
- `LEFT(text, n)` / `RIGHT(text, n)` / `MID(text, start, n)`: Character extraction
- `LEN(text)`: Text character count
- `TRIM(text)`: Removes extra spaces between words
- `UPPER / LOWER / PROPER`: Case converter
- `TEXTJOIN(delimiter, ignore_empty, range)`: Join text with custom delimiter (e.g. `=TEXTJOIN(", ", TRUE, A2:A10)`)
- `SUBSTITUTE(text, old, new)`: Specific replacement (e.g. removing hyphens from phone numbers)
- `TEXT(value, format)`: Number/date formatting to text (e.g. `=TEXT(A2,"dd-mm-yyyy")`)
- `TEXTSPLIT(text, delimiter)`: Dynamic array text splitting across cells

---

## 6. Date & Time Functions
- `TODAY()` / `NOW()`: Current date / datetime
- `DATE(year, month, day)`: Build date object
- `DATEDIF(start, end, unit)`: Difference between 2 dates (`"Y"`, `"M"`, `"D"`) (e.g. `=DATEDIF(B2, TODAY(), "Y") & " Years"`)
- `EDATE(date, months)` / `EOMONTH(date, months)`: Month additions & end of month
- `NETWORKDAYS(start, end)`: Net working days excluding weekends & holidays
- `WEEKDAY(date)`: Day number (1=Sunday or 1=Monday)

---

## 7. Math & Statistical Functions
- `SUMIFS(sum_range, crit_range1, crit1, ...)`: Multi-criteria sum
- `COUNTIFS / AVERAGEIFS`: Multi-criteria counting & averaging
- `SUMPRODUCT(range1, range2)`: Multiplies arrays & sums results (powerful multi-condition sum)
- `SUBTOTAL(function_num, range)`: Aggregate that ignores filtered/hidden rows (`9` or `109` code)
- `ROUND / ROUNDUP / ROUNDDOWN`: Controlled decimal precision
- `RANK(number, range, order)`: Value ranking

---

## 8. Data Validation & Cascading Dropdowns
- **Validation Types:** List dropdown, whole number boundaries, date restrictions, custom formula validation (`=COUNTIF($A$2:$A$100, A2)=1` for duplicate prevention).
- **Dependent (Cascading) Dropdowns:** Use `=INDIRECT(CategoryCell)` as the List Source after creating named ranges for subcategories!

---

## 9. Conditional Formatting
- Highlight Cells Rules, Data Bars, Color Scales, Icon Sets.
- **Whole Row Highlight Formula:** `=$C2="Pending"` (Lock column `$C` so the full row gets formatted).

---

## 10. PivotTables & Analytics
- Rows/Columns, Values (Sum, Count, Avg, % of Total, Running Total), Filters, Slicers, Timelines.
- Grouping dates by Month/Quarter/Year.
- Calculated Fields for custom metrics (`=Profit/Revenue`).
- Refresh All shortcut: `Ctrl + Alt + F5`.

---

## 11. Charts, Visualization & Executive Dashboards
- Column/Bar, Line, Pie, Combo Charts (Bar + Secondary Line Axis), PivotCharts, Sparklines.
- **Executive Dashboard:** Combine Slicers + PivotCharts + Conditional Formatting.

---

## 12. What-If Analysis & Solver
- **Goal Seek:** Reverse-calculates required input for a target output profit/units.
- **Data Tables:** 1-variable and 2-variable scenario modeling.
- **Scenario Manager:** Best Case, Worst Case, Expected Case comparison.
- **Solver Add-in:** Optimization engine with constraint rules.

---

## 13. Dynamic Arrays (Excel 365 / 2021+)
- `UNIQUE(range)`: Unique list generator
- `FILTER(range, condition)`: Dynamic data extraction
- `SORT(range, col, order)` / `SORTBY(range, by_range)`: Auto-sorting
- `SEQUENCE(rows, cols, start, step)`: Serial number generator

---

## 14. Excel Tables (Ctrl+T)
- Structured referencing (`Table1[Sales]`).
- Auto-expanding formulas, formatting, total row & Pivot source compatibility.

---

## 15. Named Ranges
- Define names in Name Manager for readable formulas: `=B2 * TaxRate`.

---

## 16. Power Query (Data Cleaning & ETL)
- Data -> Get & Transform
- **Merge Queries:** VLOOKUP replacement for complex relational JOINs.
- **Append Queries:** Stacking multiple tables with identical structure (UNION).
- **Applied Steps:** Recorded automated cleaning workflow, refreshable anytime!

---

## 17. Power Pivot & Data Model (DAX)
- Handles 1 Million+ rows data models across multiple related tables.
- **DAX Formulas:** `CALCULATE(SUM(Sales[Amount]), Sales[Region]="West")`, `SUMX()`, `RELATED()`, `DISTINCTCOUNT()`.

---

## 18. Macros & VBA Basics
- Macro Recorder, VBA Editor (`Alt + F11`), Developer tab.
- **VBA Code Example:**
```vba
Sub HighlightHighSales()
    Dim cell As Range
    For Each cell In Range("B2:B100")
        If cell.Value > 50000 Then
            cell.Interior.Color = RGB(144, 238, 144)
        End If
    Next cell
End Sub
```
- Always save workbooks containing macros as **.xlsm** format!

---

## 19. Data Analysis Tools (Analysis ToolPak)
- Enable via File -> Options -> Add-ins -> Analysis ToolPak.
- Perform Descriptive Statistics, Regression Analysis, Correlation matrices, Moving Averages, Histograms.

---

## 20. Error Handling Cheat Sheet
| Error | Meaning / Reason | Solution |
|---|---|---|
| **#N/A** | Value not found (Lookup failed) | Check key or use `IFERROR()` / `IFNA()` |
| **#REF!** | Invalid cell reference (Row deleted) | Fix formula broken link |
| **#DIV/0!** | Division by zero | Wrap in `IF(B2=0, 0, A2/B2)` |
| **#VALUE!** | Wrong data type used | Convert data type with `VALUE()` |
| **#NAME?** | Function name typo | Fix formula spelling |
| **#####** | Column width too narrow | Double click column border |

---

## 21. Job Best Practices & 20 Practical Scenarios

### 7 Golden Rules for Jobs & Interviews
1. Avoid hardcoded numbers inside formulas — use named cells/ranges.
2. Convert all raw datasets into Excel Tables (`Ctrl+T`).
3. Master absolute references (`$A$1`) in SUMIFS and Lookups.
4. Prefer XLOOKUP over VLOOKUP when using Excel 365.
5. Avoid heavy SUMPRODUCT on huge datasets — use INDEX-MATCH or Power Query.
6. Manually refresh PivotTables (`Ctrl+Alt+F5`) after data updates.
7. Save macro files as `.xlsm`.

### 20 Real-World Practical Scenarios
1. **Combine First & Last Name:** `=A2 & " " & B2` or `=TEXTJOIN(" ", TRUE, A2, B2)`
2. **Extract Email Domain:** `=MID(A2, FIND("@", A2)+1, LEN(A2))`
3. **Find Duplicates:** `=COUNTIF($A$2:$A$100, A2)>1` or Conditional Formatting
4. **Compare 2 Lists for Missing Items:** `=IF(ISNA(MATCH(A2, List2, 0)), "Missing", "Found")`
5. **Running Total:** `=SUM($C$2:C2)` copied down
6. **Highlight Overdue Deadlines:** `=$D2 < TODAY()`
7. **Dynamic Dashboard Dropdown:** Data Validation List + `INDEX/MATCH` or `CHOOSE`
8. **Consolidate Monthly Sheets:** Power Query Append or 3-D SUM `=SUM('Jan:Dec'!B5)`
9. **Clean Messy Text:** `=TRIM(PROPER(A2))` or Power Query
10. **Calculate Age:** `=DATEDIF(B2, TODAY(), "Y") & " Years"`
11. **Count Employee Attendance:** `=COUNTIF(B2:AF2, "P")`
12. **Calculate Loan EMI:** `=PMT(Rate/12, Months, -LoanAmount)`
13. **Filter Sales > 50k (West Region):** `=FILTER(A2:D100, (B2:B100="West")*(C2:C100>50000))`
14. **Get Unique Customer List:** `=UNIQUE(Sales[Customer])`
15. **Find 2nd Highest Salary:** `=LARGE(SalaryRange, 2)`
16. **Safe Division:** `=IFERROR(A2/B2, 0)`
17. **Check Text Contains Word:** `=ISNUMBER(SEARCH("Approved", A2))`
18. **Calculate Required Target for Profit:** Data Tab ➔ What-If Analysis ➔ Goal Seek
19. **Dynamic Subcategory Dropdown:** Data Validation Source `=INDIRECT(CategoryCell)`
20. **Auto-Highlight High Sales:** VBA Macro Procedure `HighlightHighSales()`
