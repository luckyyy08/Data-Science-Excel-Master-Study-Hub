# Advanced Excel Master Textbook — Complete Handwritten Study Hub

**Prepared for:** Lokesh Ahire (Lucky)  
**Topics Covered:** 24 Modules covering Excel Interface, Locking ($), Logical, Lookups (VLOOKUP, XLOOKUP, INDEX+MATCH), Text Cleaning, Date/Time, Math & Stats, Financial (PMT/IRR), Dynamic Arrays, Data Validation, Conditional Formatting, PivotTables, Advanced Charts, What-If Solver, Power Query ETL, Power Pivot DAX, VBA Macros, Security, 100+ Shortcuts, Error Troubleshooting & 100 Real-World Case Studies.

---

## Module 1: Excel Basics, Interface & Ribbon Complete Guide
**Category:** Fundamentals | **Description:** Workbook vs Worksheet, Cell Grid Structure, Formula Bar, Name Box, Ribbon Tabs & Quick Access Toolbar


      ### 1.1 Excel Architecture & Terminology

      **Workbook (.xlsx / .xlsm):** Pura Excel file. Ek workbook ke andar multiple sheets/tabs ho sakte hain.


      **Worksheet (Sheet):** Individual grid tab jahan data enter aur analyze hota hai. Standard Excel me 1,048,576 rows aur 16,384 columns hote hain (Col A to XFD).


      **Cell:** Row aur Column ka intersection. Example: `A1` (Column A, Row 1), `B25`, `XFD1048576`.


      **Range:** Multiple cells ka rectangle group. Syntax: `TopLeftCell:BottomRightCell` (e.g., `A1:B10`, `C5:G50`).


      
      ### 1.2 Ribbon Tabs & Interface Elements

      <ul>
        <li>**Name Box:** Formula bar ke left side me rehta hai. Isme active cell address dikhta hai. Isme address likhkar Enter dabane se direct us cell par jump kar sakte ho. Named Ranges bhi Name Box me dikhte hain.</li>
        <li>**Formula Bar:** Active cell ka exact formula ya raw value display karta hai. Single line ya expandable multiline view support karta hai.</li>
        <li>**Ribbon Tabs:**
          <ul>
            <li>**Home:** Formatting, Clipboard, Alignment, Number Format, Conditional Formatting, Cells, Editing.</li>
            <li>**Insert:** PivotTables, Tables, Pictures, Charts, Sparklines, Slicers, Timelines, Text Box, Symbols.</li>
            <li>**Page Layout:** Themes, Page Setup, Margins, Orientation, Print Area, Gridlines view/print toggle.</li>
            <li>**Formulas:** Insert Function, Function Library (Logical, Text, Date/Time, Lookup), Name Manager, Formula Auditing (Trace Precedents/Dependents, Evaluate Formula).</li>
            <li>**Data:** Get & Transform Data (Power Query), Queries & Connections, Sort & Filter, Data Tools (Text to Columns, Flash Fill, Remove Duplicates, Data Validation, Consolidate), What-If Analysis, Forecast, Group/Ungroup.</li>
            <li>**Review:** Spelling, Thesaurus, Translate, Comments, Protect Sheet, Protect Workbook, Allow Edit Ranges.</li>
            <li>**View:** Workbook Views (Normal, Page Break Preview, Page Layout), Show (Gridlines, Formula Bar, Headings), Zoom, Freeze Panes, Split Window, Macros.</li>
            <li>**Developer:** Visual Basic Editor (Alt+F11), Macros, Record Macro, Form Controls (Button, Combo Box), ActiveX Controls, XML. (File → Options → Customize Ribbon se enable karein).</li>
          </ul>
        </li>
      </ul>

      <div class="note-box">
        **💡 Pro Tip:** Agar Gridlines gayab ho jayen, to `View Tab → Show → Check Gridlines` par click karo!
      </div>
    

---

## Module 2: Cell References & Address Locking ($) Master Class
**Category:** Core Formula Rules | **Description:** Relative (A1), Absolute ($A$1), Mixed ($A1 / A$1), 3-D References across Sheets & Structured Table References


      ### 2.1 The 3 Types of Cell References

      <table class="notes-table">
        <thead>
          <tr>
            <th>Reference Type</th>
            <th>Syntax</th>
            <th>Behavior when Formula is Copied / Dragged</th>
            <th>Real World Use Case Example</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**Relative Reference**</td>
            <td>`A1`</td>
            <td>Row aur Column dono change hote hain target direction ke mutabiq.</td>
            <td>Calculating Line Total: `=B2*C2` (Qty * Price) in row 2, dragged down to row 100.</td>
          </tr>
          <tr>
            <td>**Absolute Reference**</td>
            <td>`$A$1`</td>
            <td>Row aur Column dono COMPLETE LOCK (Fix) rehte hain. Kabhi change nahi hote.</td>
            <td>Fixed Tax Rate cell `$B$1` ko 1,000 different product rows me multiply karna: `=C2*$B$1`.</td>
          </tr>
          <tr>
            <td>**Mixed (Col Lock)**</td>
            <td>`$A1`</td>
            <td>Column A fix rehta hai, lekin Row (1, 2, 3...) change hoti hai.</td>
            <td>Multiplication matrix table ya horizontal VLOOKUP range lookup column lock.</td>
          </tr>
          <tr>
            <td>**Mixed (Row Lock)**</td>
            <td>`A$1`</td>
            <td>Row 1 fix rehti hai, lekin Column (A, B, C...) change hota hai.</td>
            <td>Column Header rate ya discount calculation jab formulas right side drag ho rahe hon.</td>
          </tr>
        </tbody>
      </table>

      ### 2.2 The Magic Shortcut: F4

      Formula editing mode me cell reference select karke **F4** key press karne se toggle hota hai:


      `A1` → (F4) → `$A$1` → (F4) → `A$1` → (F4) → `$A1` → (F4) → `A1`



      ### 2.3 3-D References Across Multiple Worksheets

      Jab tumhe multiple identical sheets (e.g., Jan, Feb, Mar) ke same cell (e.g., B10) ka sum nikalna ho:


      <div class="code-snippet">=SUM(Jan:Mar!B10)</div>
      Ye formula Jan, Feb aur Mar teeno sheets ke B10 cell ko auto-sum kar dega!



      ### 2.4 External Workbook References

      Dusre workbook file se value link karna:


      <div class="code-snippet">='[Annual_Report_2026.xlsx]Sales_Data'!$C$15</div>

      ### 2.5 Structured References in Excel Tables (Ctrl+T)

      Jab range ko Excel Table bana lete ho (Ctrl+T), to cell coordinates (B2:B100) ke bajaye column names use hote hain:


      <div class="code-snippet">=SUM(SalesTable[Sales_Amount])</div>
      <div class="code-snippet">=[@Qty] * [@Price]  <span class="comment">-- Current row calculation</span></div>
    

---

## Module 3: Logical Functions & Decision Trees
**Category:** Formulas & Logic | **Description:** IF, Nested IF, IFS, AND, OR, NOT, XOR, IFERROR, IFNA, SWITCH & IS-Information Functions


      ### 3.1 Complete Logical Function Reference

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Description & Logic</th>
            <th>Example Formula</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**IF**</td>
            <td>`=IF(cond, val_if_true, val_if_false)`</td>
            <td>Ek single condition test karta hai. True hone pe first value, false hone pe second.</td>
            <td>`=IF(B2>=40, "Pass", "Fail")`</td>
          </tr>
          <tr>
            <td>**Nested IF**</td>
            <td>`=IF(c1, v1, IF(c2, v2, v3))`</td>
            <td>Multiple conditions check karne ke liye IF ke andar IF (Old method).</td>
            <td>`=IF(A2>=90,"A",IF(A2>=75,"B","C"))`</td>
          </tr>
          <tr>
            <td>**IFS**</td>
            <td>`=IFS(c1, v1, c2, v2, ...)`</td>
            <td>Nested IF ka modern replacement (Excel 2019/365). Multiple conditions clean likho.</td>
            <td>`=IFS(A2>=90,"A",A2>=75,"B",A2>=50,"C",TRUE,"Fail")`</td>
          </tr>
          <tr>
            <td>**AND**</td>
            <td>`=AND(cond1, cond2, ...)`</td>
            <td>Returns TRUE jab SABHI conditions TRUE hon. Ek bhi false to output FALSE.</td>
            <td>`=IF(AND(Age>=18, Marks>=60), "Eligible", "Not Eligible")`</td>
          </tr>
          <tr>
            <td>**OR**</td>
            <td>`=OR(cond1, cond2, ...)`</td>
            <td>Returns TRUE jab KAM SE KAM EK condition TRUE ho.</td>
            <td>`=IF(OR(Dept="Sales", Dept="Mktg"), "Bonus 10%", "Bonus 5%")`</td>
          </tr>
          <tr>
            <td>**NOT**</td>
            <td>`=NOT(condition)`</td>
            <td>Condition ka output reverse (flip) karta hai (TRUE → FALSE, FALSE → TRUE).</td>
            <td>`=IF(NOT(Status="Completed"), "Pending Followup", "OK")`</td>
          </tr>
          <tr>
            <td>**IFERROR**</td>
            <td>`=IFERROR(formula, value_if_error)`</td>
            <td>Formula me koi bhi error (#N/A, #VALUE!, #DIV/0!, #REF!) aane pe custom value dikhata hai.</td>
            <td>`=IFERROR(VLOOKUP(A2, Data, 2, 0), "Data Missing")`</td>
          </tr>
          <tr>
            <td>**IFNA**</td>
            <td>`=IFNA(formula, value_if_na)`</td>
            <td>Sirf #N/A error handle karta hai. Baaki errors ko pass hone deta hai debugging ke liye.</td>
            <td>`=IFNA(XLOOKUP(A2, B:B, C:C), "Not Found")`</td>
          </tr>
          <tr>
            <td>**SWITCH**</td>
            <td>`=SWITCH(val, case1, res1, case2, res2, default)`</td>
            <td>Value ko exact match list se compare karta hai (Modern alternative to IFS).</td>
            <td>`=SWITCH(A2, 1, "Mon", 2, "Tue", 3, "Wed", "Invalid Day")`</td>
          </tr>
        </tbody>
      </table>

      ### 3.2 IS-Information Functions (Data Auditing)

      <ul>
        <li>`=ISBLANK(cell)`: Cell empty hai to TRUE.</li>
        <li>`=ISNUMBER(cell)`: Cell me valid number hai to TRUE.</li>
        <li>`=ISTEXT(cell)`: Cell me text value hai to TRUE.</li>
        <li>`=ISFORMULA(cell)`: Cell me formula written hai to TRUE.</li>
        <li>`=ISERROR(cell)`: Koi bhi error ho to TRUE.</li>
      </ul>
    

---

## Module 4: Lookup & Reference Functions Master Class
**Category:** Most Important Section | **Description:** VLOOKUP, HLOOKUP, XLOOKUP, INDEX+MATCH, XMATCH, OFFSET, INDIRECT, CHOOSE, LOOKUP, ROW, COLUMN, ADDRESS, TRANSPOSE


      ### 4.1 VLOOKUP (Vertical Lookup)

      <div class="code-snippet">=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])</div>
      <ul>
        <li>**lookup_value:** Wo value jise search karna hai (e.g., EmpID A2).</li>
        <li>**table_array:** Data table range jahan search karna hai. Rule: Search column hamesha Table Array ka 1st column hona chahiye!</li>
        <li>**col_index_num:** Result column ka number (1, 2, 3...).</li>
        <li>**range_lookup:** Hamesha `0` (ya `FALSE`) lagao Exact Match ke liye. `1` (`TRUE`) approximate match ke liye.</li>
      </ul>
      **VLOOKUP Limitations:** Sirf right-side ke columns lookup karta hai (Left Lookup nahi kar sakta). Target table me new column insert karne se `col_index_num` break ho jata hai.



      ### 4.2 XLOOKUP (Excel 365 / 2021+ Modern Lookup)

      <div class="code-snippet">=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])</div>
      <ul>
        <li>**lookup_array:** Sirf wo single column jisme search karna hai (e.g., `EmpDB[EmpID]`).</li>
        <li>**return_array:** Sirf wo single column ya multiple columns jo return karne hain (e.g., `EmpDB[Salary]`).</li>
        <li>**if_not_found:** Built-in error message string if value is missing (e.g., `"Employee Not Found"`). No need for IFERROR!</li>
      </ul>
      **Why XLOOKUP is Best:** Left and Right dono directions me kaam karta hai. Column insert hone par bhi nahi toot-ta. Default exact match hota hai.



      ### 4.3 INDEX + MATCH Combo (Professional Gold Standard)

      <div class="code-snippet">=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))</div>
      <ul>
        <li>**MATCH(lookup_val, lookup_range, 0):** Returns row position (number) of the item.</li>
        <li>**INDEX(return_range, row_num):** Returns value at that row position from return range.</li>
      </ul>
      **Example:** EmpID se Employee Name (jo left column A me hai) fetch karna:


      <div class="code-snippet">=INDEX(A2:A100, MATCH(E2, B2:B100, 0))</div>

      ### 4.4 Advanced Lookup Functions Table

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Real World Usage Description</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**INDIRECT**</td>
            <td>`=INDIRECT(ref_text)`</td>
            <td>Text string ko actual cell reference me convert karta hai. Dynamic sheet lookup aur dependent dropdowns ke liye best. Example: `=SUM(INDIRECT(A1 & "!B2:B10"))`</td>
          </tr>
          <tr>
            <td>**OFFSET**</td>
            <td>`=OFFSET(ref, rows, cols, [height], [width])`</td>
            <td>Starting cell se X rows down aur Y cols right shift karke dynamic range banata hai. Dynamic Named Ranges sathi use hota hai.</td>
          </tr>
          <tr>
            <td>**CHOOSE**</td>
            <td>`=CHOOSE(index_num, val1, val2, ...)`</td>
            <td>Index number ke aadhar par value pick karta hai. Financial modeling me scenario switching ke liye use hota hai. Example: `=CHOOSE(B1, "Base", "Best", "Worst")`</td>
          </tr>
          <tr>
            <td>**XMATCH**</td>
            <td>`=XMATCH(val, range)`</td>
            <td>Modern version of MATCH function. Reverse search (bottom-to-top) aur wildcard match easily karta hai.</td>
          </tr>
          <tr>
            <td>**TRANSPOSE**</td>
            <td>`=TRANSPOSE(range)`</td>
            <td>Rows ko Columns me aur Columns ko Rows me flip (rotate) karta hai dynamically.</td>
          </tr>
        </tbody>
      </table>
    

---

## Module 5: Text Functions & Data Cleaning Complete Reference
**Category:** Data Cleaning | **Description:** LEFT, RIGHT, MID, LEN, TRIM, CLEAN, UPPER, LOWER, PROPER, CONCAT, TEXTJOIN, SUBSTITUTE, REPLACE, FIND, SEARCH, TEXT, VALUE, TEXTSPLIT, TEXTBEFORE, TEXTAFTER


      ### 5.1 Text Manipulation & Cleaning Functions

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Description & Example</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**TRIM**</td>
            <td>`=TRIM(text)`</td>
            <td>Words ke beech single space chhodkar aage, peeche aur beech ke sabhi EXTRA SPACES remove karta hai. Mandatory for VLOOKUP cleanup!</td>
          </tr>
          <tr>
            <td>**CLEAN**</td>
            <td>`=CLEAN(text)`</td>
            <td>Database import se aaye non-printable ASCII line-break characters (CHAR 0-31) delete karta hai.</td>
          </tr>
          <tr>
            <td>**LEFT / RIGHT**</td>
            <td>`=LEFT(text, n)` / `=RIGHT(text, n)`</td>
            <td>Start (Left) ya End (Right) se `n` characters extract karta hai. `=LEFT("IN-9081", 2) → "IN"`</td>
          </tr>
          <tr>
            <td>**MID**</td>
            <td>`=MID(text, start_num, num_chars)`</td>
            <td>Text ke beech se specific position se characters nikalta hai. `=MID("EMP-4509-P", 5, 4) → "4509"`</td>
          </tr>
          <tr>
            <td>**LEN**</td>
            <td>`=LEN(text)`</td>
            <td>Text ki Total Character Length count karke batata hai (spaces include karke).</td>
          </tr>
          <tr>
            <td>**UPPER / LOWER / PROPER**</td>
            <td>`=PROPER(text)`</td>
            <td>Text Case convert karta hai. PROPER: Every word ka 1st letter Capital karta hai (Names formatting ke liye best).</td>
          </tr>
          <tr>
            <td>**TEXTJOIN**</td>
            <td>`=TEXTJOIN(delimiter, ignore_empty, range)`</td>
            <td>Multiple cells ko specific delimiter (e.g. comma ", ") se join karta hai. Blank cells ignore karne ka option hota hai. `=TEXTJOIN(", ", TRUE, A2:A10)`</td>
          </tr>
          <tr>
            <td>**SUBSTITUTE**</td>
            <td>`=SUBSTITUTE(text, old_text, new_text, [instance])`</td>
            <td>Text me jahan bhi `old_text` mile use `new_text` se replace karta hai. Example: `=SUBSTITUTE(A2, "-", "/")`</td>
          </tr>
          <tr>
            <td>**FIND vs SEARCH**</td>
            <td>`=FIND(find_text, within_text)`</td>
            <td>Character ki exact numerical position batata hai. **FIND is Case-Sensitive** ("A" != "a"). **SEARCH is Case-Insensitive**.</td>
          </tr>
          <tr>
            <td>**TEXT**</td>
            <td>`=TEXT(value, format_text)`</td>
            <td>Number ya Date ko custom formatted string me convert karta hai. Example: `=TEXT(TODAY(), "dd-mmm-yyyy (dddd)") → "12-Sep-2026 (Saturday)"`</td>
          </tr>
          <tr>
            <td>**TEXTSPLIT (365)**</td>
            <td>`=TEXTSPLIT(text, col_delimiter)`</td>
            <td>Delimited text ko spill karke multiple cells me divide karta hai (Text-to-Columns ka formula version).</td>
          </tr>
          <tr>
            <td>**TEXTBEFORE / TEXTAFTER (365)**</td>
            <td>`=TEXTBEFORE(text, delimiter)`</td>
            <td>Delimiter ke pehle ya baad ka saara text extract kar leta hai. Example: `=TEXTBEFORE("lokesh@gmail.com", "@") → "lokesh"`</td>
          </tr>
        </tbody>
      </table>

      <div class="note-box">
        **💡 Real World Formula (Extract Domain from Email):**<br/>
        `=MID(A2, FIND("@", A2)+1, LEN(A2))`
      </div>
    

---

## Module 6: Date & Time Functions Master Class
**Category:** Time Series Analysis | **Description:** TODAY, NOW, DATE, TIME, DAY, MONTH, YEAR, HOUR, MINUTE, SECOND, EDATE, EOMONTH, DATEDIF, NETWORKDAYS, WORKDAY, WEEKDAY, WEEKNUM, YEARFRAC


      ### 6.1 Excel Date System Fundamental

      Excel me **Date ek Serial Number hoti hai** (1 = 1st January 1900). Time ek decimal fraction hota hai (0.5 = 12:00 PM Noon). Isliye dates ko directly subtract (subtract `End_Date - Start_Date`) karke days nikal sakte ho!



      ### 6.2 Date & Time Functions Reference

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Description & Utility</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**TODAY() / NOW()**</td>
            <td>`=TODAY()` / `=NOW()`</td>
            <td>Current System Date / Current System Date & Time. Daily updating reports ke liye auto-dynamic reference.</td>
          </tr>
          <tr>
            <td>**DATE**</td>
            <td>`=DATE(year, month, day)`</td>
            <td>Year, Month aur Day numbers se valid Excel Date object construct karta hai. `=DATE(2026, 9, 12)`</td>
          </tr>
          <tr>
            <td>**DATEDIF**</td>
            <td>`=DATEDIF(start, end, unit)`</td>
            <td>Do dates ke beech exact difference calculate karta hai. Units: `"Y"` (Years), `"M"` (Months), `"D"` (Days), `"YM"` (Remaining months).</td>
          </tr>
          <tr>
            <td>**EDATE**</td>
            <td>`=EDATE(start_date, months)`</td>
            <td>Start date me exact `X` months add ya subtract (negative value) karke date batata hai. Maturity date calculation ke liye.</td>
          </tr>
          <tr>
            <td>**EOMONTH**</td>
            <td>`=EOMONTH(start_date, months)`</td>
            <td>Given month ke exact last day (End of Month) ki date nikalta hai. `=EOMONTH(TODAY(), 0) → Last day of current month`.</td>
          </tr>
          <tr>
            <td>**NETWORKDAYS**</td>
            <td>`=NETWORKDAYS(start, end, [holidays])`</td>
            <td>Do dates ke beech ke **Working Days count** karta hai. Weekends (Saturday & Sunday) aur custom holidays list ko exclude karke!</td>
          </tr>
          <tr>
            <td>**NETWORKDAYS.INTL**</td>
            <td>`=NETWORKDAYS.INTL(start, end, weekend, [holidays])`</td>
            <td>Custom weekend configuration (e.g., sirf Sunday weekend, ya Friday/Saturday) ke sath working days calculate karta hai.</td>
          </tr>
          <tr>
            <td>**WORKDAY**</td>
            <td>`=WORKDAY(start_date, days, [holidays])`</td>
            <td>Start date se exact `X` working days baad ki target completion date nikalta hai. Project SLA delivery tracking sathi best!</td>
          </tr>
          <tr>
            <td>**WEEKDAY**</td>
            <td>`=WEEKDAY(date, [return_type])`</td>
            <td>Date ke din ka index number return karta hai (1=Sunday, 2=Monday...7=Saturday).</td>
          </tr>
          <tr>
            <td>**YEARFRAC**</td>
            <td>`=YEARFRAC(start_date, end_date)`</td>
            <td>Do dates ke beech ka fraction of year calculate karta hai (e.g., 1.5 years). Financial interest calculations sathi useful.</td>
          </tr>
        </tbody>
      </table>

      <div class="note-box">
        **🎂 Real World Age Calculation Formula:**<br/>
        `=DATEDIF(DOB_Cell, TODAY(), "Y") & " Years, " & DATEDIF(DOB_Cell, TODAY(), "YM") & " Months"`
      </div>
    

---

## Module 7: Math & Trig Functions Deep Dive
**Category:** Core Math Engine | **Description:** SUM, SUMIF, SUMIFS, SUMPRODUCT, SUBTOTAL, AGGREGATE, ABS, MOD, ROUND, ROUNDUP, ROUNDDOWN, INT, TRUNC, CEILING, FLOOR, MROUND, RAND, RANDBETWEEN


      ### 7.1 Aggregation & Multi-Criteria Math Functions

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Description & Example Formula</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**SUMIF**</td>
            <td>`=SUMIF(range, criteria, [sum_range])`</td>
            <td>Single condition ke basis par numbers sum karta hai. `=SUMIF(Region, "West", Sales)`</td>
          </tr>
          <tr>
            <td>**SUMIFS**</td>
            <td>`=SUMIFS(sum_range, crit_range1, crit1, crit_range2, crit2...)`</td>
            <td>**Multiple conditions (AND logic)** ke basis par sum karta hai. Note: Sum Range sabse PEHLE likha jata hai! `=SUMIFS(Sales, Region, "West", Status, "Completed")`</td>
          </tr>
          <tr>
            <td>**SUMPRODUCT**</td>
            <td>`=SUMPRODUCT(array1, array2...)`</td>
            <td>Corresponding array elements ko multiply karke sum return karta hai. Helper columns ke bina multi-condition sum & weighted average sathi best! `=SUMPRODUCT(Qty_Range, Price_Range)`</td>
          </tr>
          <tr>
            <td>**SUBTOTAL**</td>
            <td>`=SUBTOTAL(function_num, range)`</td>
            <td>AutoFilter lagane par **Hidden / Filtered-out rows ko IGNORE karke** sum/count calculate karta hai. Code `9` or `109` = SUM.</td>
          </tr>
          <tr>
            <td>**AGGREGATE**</td>
            <td>`=AGGREGATE(function_num, options, array)`</td>
            <td>SUBTOTAL se bhi powerful! Hidden rows ke sath-sath **Error values (#N/A) ko bhi ignore** karke final calculation karta hai. Option `6` = Ignore Error Values.</td>
          </tr>
          <tr>
            <td>**MOD**</td>
            <td>`=MOD(number, divisor)`</td>
            <td>Division ka Remainder return karta hai. Alternate row zebra striping (Conditional Formatting `MOD(ROW(),2)=0`) me use hota hai.</td>
          </tr>
          <tr>
            <td>**ROUND / ROUNDUP / ROUNDDOWN**</td>
            <td>`=ROUND(number, num_digits)`</td>
            <td>Decimal precision control karta hai. ROUNDUP: Hamesha upar round karta hai. ROUNDDOWN: Hamesha niche round karta hai. `=ROUND(123.456, 2) → 123.46`</td>
          </tr>
          <tr>
            <td>**MROUND**</td>
            <td>`=MROUND(number, multiple)`</td>
            <td>Number ko specified multiple me round karta hai. Example: Price ko nearest 5 ya 10 ke multiple me round karna: `=MROUND(47, 5) → 45`.</td>
          </tr>
          <tr>
            <td>**RAND / RANDBETWEEN**</td>
            <td>`=RANDBETWEEN(bottom, top)`</td>
            <td>Specified Range me Random Integer numbers generate karta hai (Mock Data set tayar karne sathi useful). `=RANDBETWEEN(1000, 9999)`</td>
          </tr>
        </tbody>
      </table>
    

---

## Module 8: Statistical Functions & Analytics
**Category:** Data Science & MIS | **Description:** AVERAGE, AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, MAX, MAXIFS, MIN, MINIFS, MEDIAN, MODE, LARGE, SMALL, RANK.EQ, FREQUENCY


      ### 8.1 Complete Statistical Analytics Functions

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Description & Example Usage</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**COUNT vs COUNTA vs COUNTBLANK**</td>
            <td>`=COUNT(range)` / `=COUNTA(range)`</td>
            <td>COUNT: Sirf NUMERIC cells count karta hai. COUNTA: Non-empty (Text + Numbers) cells count karta hai. COUNTBLANK: Khali cells count karta hai.</td>
          </tr>
          <tr>
            <td>**COUNTIF / COUNTIFS**</td>
            <td>`=COUNTIFS(range1, crit1, range2, crit2...)`</td>
            <td>Specified criteria ke basis par rows count karta hai. Example: Duplicate prevention rule `=COUNTIF($A$2:$A$100, A2)>1`.</td>
          </tr>
          <tr>
            <td>**AVERAGEIF / AVERAGEIFS**</td>
            <td>`=AVERAGEIFS(avg_range, crit_range1, crit1...)`</td>
            <td>Specified conditions match karne wali rows ka Average calculate karta hai.</td>
          </tr>
          <tr>
            <td>**MAXIFS / MINIFS**</td>
            <td>`=MAXIFS(max_range, crit_range1, crit1...)`</td>
            <td>Condition match karne wale group ka Maximum ya Minimum value nikalta hai. `=MAXIFS(Salary, Dept, "IT")`</td>
          </tr>
          <tr>
            <td>**LARGE / SMALL**</td>
            <td>`=LARGE(array, k)` / `=SMALL(array, k)`</td>
            <td>Dataset me se **N-th Largest (Top 2nd, Top 3rd)** ya **N-th Smallest** value return karta hai. `=LARGE(Sales, 2) → 2nd Highest Sale`.</td>
          </tr>
          <tr>
            <td>**RANK.EQ**</td>
            <td>`=RANK.EQ(number, ref, [order])`</td>
            <td>List me specific number ki Rank (1st, 2nd, 3rd position) calculate karta hai. Order `0` = Descending (Highest number = Rank 1).</td>
          </tr>
          <tr>
            <td>**MEDIAN / MODE.SNGL**</td>
            <td>`=MEDIAN(range)`</td>
            <td>MEDIAN: Middlemost value (outliers bias handle karne sathi best). MODE: Sabse frequent repeated value.</td>
          </tr>
          <tr>
            <td>**STDEV.S / VAR.S**</td>
            <td>`=STDEV.S(range)`</td>
            <td>Sample Standard Deviation & Variance calculate karta hai (Risk & Volatility analysis in Finance & Quality Control).</td>
          </tr>
        </tbody>
      </table>
    

---

## Module 9: Financial Functions (Banking, EMI & Investment)
**Category:** Finance & Accounting | **Description:** PMT, IPMT, PPMT, RATE, NPER, PV, FV, NPV, IRR, XNPV, XIRR, CUMIPMT, CUMPRINC


      ### 9.1 Loan EMI & Investment Functions

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Real World Financial Use Case</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**PMT**</td>
            <td>`=PMT(rate, nper, pv, [fv])`</td>
            <td>**Loan EMI (Monthly Payment)** calculate karta hai. Note: Annual Interest Rate ko `12` se divide karo (`Rate/12`) aur Years ko 12 se multiply karo (`Years*12`)!</td>
          </tr>
          <tr>
            <td>**IPMT**</td>
            <td>`=IPMT(rate, per, nper, pv)`</td>
            <td>EMI ke andar specific period/month ka **Interest Component (Vyaj)** kitna hai wo batata hai.</td>
          </tr>
          <tr>
            <td>**PPMT**</td>
            <td>`=PPMT(rate, per, nper, pv)`</td>
            <td>EMI ke andar specific month ka **Principal Component (Mool)** kitna hai wo batata hai. (Note: `PMT = IPMT + PPMT`).</td>
          </tr>
          <tr>
            <td>**FV**</td>
            <td>`=FV(rate, nper, pmt, [pv])`</td>
            <td>SIP / Recurring Investment ki **Future Value (Maturity Amount)** calculate karta hai.</td>
          </tr>
          <tr>
            <td>**NPV vs XNPV**</td>
            <td>`=XNPV(rate, values, dates)`</td>
            <td>Net Present Value. NPV standard equal annual intervals ke liye. **XNPV irregular real-life dates** ke cashflows ke liye valuation karta hai.</td>
          </tr>
          <tr>
            <td>**IRR vs XIRR**</td>
            <td>`=XIRR(values, dates, [guess])`</td>
            <td>Internal Rate of Return (Annualized Return % of Project/SIP). XIRR is mandatory in Wealth Management & Private Equity!</td>
          </tr>
        </tbody>
      </table>

      <div class="note-box">
        **💰 Example Loan EMI Formula (₹5,000,000 Home Loan at 8.5% p.a. for 20 Years):**<br/>
        `=PMT(8.5%/12, 20*12, -5000000) → ₹43,391 per month`
      </div>
    

---

## Module 10: Database Functions (DSUM, DCOUNT, DAVERAGE...)
**Category:** Advanced Querying | **Description:** DSUM, DCOUNT, DCOUNTA, DAVERAGE, DMAX, DMIN, DGET for criteria-table driven database calculations


      ### 10.1 Database Functions Architecture

      Database Functions Excel ke internal mini-SQL engines hain. Ye pure table me se criteria-block range ke mutabiq dynamic aggregate karte hain. Sabhi Database functions `D` se start hote hain.


      
      <div class="code-snippet">=DSUM(database_range, field_name, criteria_range)</div>
      
      ### 10.2 Database Function Reference

      <ul>
        <li>**DSUM:** Criteria table ke mutabiq total sum calculate karta hai.</li>
        <li>**DAVERAGE:** Specified criteria matching rows ka average nikalta hai.</li>
        <li>**DCOUNT / DCOUNTA:** Criteria match karne wale numerical ya text rows count karta hai.</li>
        <li>**DMAX / DMIN:** Filtered Database criteria matching group ka Max/Min fetch karta hai.</li>
        <li>**DGET:** Criteria matching exact single record extract karta hai (Lookup alternative).</li>
      </ul>
    

---

## Module 11: Dynamic Arrays & Modern Excel 365 Functions
**Category:** Modern Excel 365 | **Description:** UNIQUE, SORT, SORTBY, FILTER, SEQUENCE, RANDARRAY, Spill Operator, LET, LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL


      ### 11.1 The Spill Behavior & Spill Operator (#)

      Modern Excel 365 me formulas **Array (Multiple Cells) return karte hain**. Jab tum ek cell me formula likhte ho, to results niche aur right me auto-spill hote hain. Spilled range ka reference lene ke liye **# (Spill Range Operator)** use hota hai (e.g., `A2#`).


      **#SPILL! Error:** Agar spill path me koi data cell block kar raha ho, to Excel #SPILL! error throw karta hai. Path clear karte hi error gayab ho jayega.



      ### 11.2 Modern Dynamic Array Functions Table

      <table class="notes-table">
        <thead>
          <tr>
            <th>Function</th>
            <th>Syntax</th>
            <th>Real World Example Formula</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**UNIQUE**</td>
            <td>`=UNIQUE(array)`</td>
            <td>List me se Duplicates remove karke exact Unique items extract karta hai. `=UNIQUE(A2:A100)`</td>
          </tr>
          <tr>
            <td>**FILTER**</td>
            <td>`=FILTER(array, include, [if_empty])`</td>
            <td>Condition ke basis par poore raw data table me se matching rows dynamically extract karke spill karta hai. `=FILTER(A2:D100, B2:B100="West", "No Data")`</td>
          </tr>
          <tr>
            <td>**SORT**</td>
            <td>`=SORT(array, [sort_index], [sort_order])`</td>
            <td>Data ko specific column index ke basis pe Sort karta hai (1 = Ascending, -1 = Descending). `=SORT(A2:C100, 3, -1)`</td>
          </tr>
          <tr>
            <td>**SORTBY**</td>
            <td>`=SORTBY(array, by_array1, [order1]...)`</td>
            <td>Ek range ko kisi dusre range/column ke basis pe sort karta hai. `=SORTBY(EmpNames, Salaries, -1)`</td>
          </tr>
          <tr>
            <td>**SEQUENCE**</td>
            <td>`=SEQUENCE(rows, [cols], [start], [step])`</td>
            <td>Numbers ki custom series (1, 2, 3... 100) auto-generate karke array fill karta hai. `=SEQUENCE(10, 1, 100, 10)`</td>
          </tr>
          <tr>
            <td>**LET**</td>
            <td>`=LET(var1, val1, var2, val2, calculation)`</td>
            <td>Formula ke andar temporary variables define karta hai. Complex nested formulas ki speed 10x badha deta hai!</td>
          </tr>
          <tr>
            <td>**LAMBDA**</td>
            <td>`=LAMBDA(param1, param2, calculation)`</td>
            <td>Bina VBA ke apna Custom Reusable Excel Function banata hai! Name Manager me save karke use kar sakte ho.</td>
          </tr>
        </tbody>
      </table>

      <div class="note-box">
        **🔥 Master Combo Formula (Filter + Sort + Unique):**<br/>
        `=SORT(UNIQUE(FILTER(SalesDB[Customer], SalesDB[Amount]>50000)))`
      </div>
    

---

## Module 12: Data Cleaning, Formatting & Transformation
**Category:** Data Wrangling | **Description:** Text to Columns, Flash Fill (Ctrl+E), Remove Duplicates, Consolidate, Custom Number Formatting rules


      ### 12.1 Text to Columns Wizard

      Data Tab → Text to Columns: Single column data ko delimiter (Comma, Tab, Space, Pipe |) ya Fixed Width ke basis par multiple columns me divide karna.



      ### 12.2 Flash Fill (Ctrl + E)

      Excel ka AI Pattern Detector! First row me desired output manual type karo, phir agle cell me **Ctrl + E** press karo — Excel poore column me pattern pehchan kar auto-fill kar dega! (First Name, Last Name extraction, Phone formatting sathi best).



      ### 12.3 Custom Number Formatting Rules Architecture

      Format Cells (Ctrl + 1) → Number → Custom. Custom format code ke 4 sections hote hain (Semicolon ; se separated):


      <div class="code-snippet">POSITIVE_FORMAT ; NEGATIVE_FORMAT ; ZERO_FORMAT ; TEXT_FORMAT</div>
      
      #### Examples of Custom Format Codes:

      <ul>
        <li>`#,##0.00;[Red](#,##0.00);"-" ;"Check"@` → Positive green/normal, Negative red with brackets, Zero hyphen, Text custom.</li>
        <li>`[>=10000000]0.0,," Cr";[>=100000]0.0, " Lakh";0.0` → Automatic Lakhs & Crores formatting!</li>
        <li>`dd-mmm-yyyy (dddd)` → 12-Sep-2026 (Saturday).</li>
        <li>`"INR "#,##0` → Adds currency prefix dynamically without breaking math formulas!</li>
      </ul>
    

---

## Module 13: Data Validation & Cascading Dropdowns
**Category:** Data Quality Control | **Description:** Dropdown Lists, Whole Numbers, Dates, Text Length, Custom Formula Rules & Dependent Cascading Dropdowns (=INDIRECT)


      ### 13.1 Data Validation Setup & Types

      **Path:** Data Tab → Data Tools → Data Validation (Shortcut: `Alt + A + V + V`).


      <ul>
        <li>**List:** Dropdown List bananta hai. Source me comma-separated values (`Pending, Approved, Rejected`) ya cell range (`=$F$1:$F$5`) do.</li>
        <li>**Whole Number / Decimal:** Input boundary restrict karta hai (e.g., Age 18 to 60 only).</li>
        <li>**Date:** Restricted date window (e.g., Only future dates `>=TODAY()`).</li>
        <li>**Text Length:** Restrict exact characters count (e.g., Mobile Number exactly 10 digits `LEN(A2)=10`).</li>
        <li>**Custom Formula Rules:** Formula evaluation boolean TRUE hone par hi value accept hoti hai.</li>
      </ul>

      ### 13.2 Formula Rules for Data Integrity

      <ul>
        <li>**Prevent Duplicate Entries in Column A:**<br/>`=COUNTIF($A$2:$A$100, A2)=1`</li>
        <li>**Restrict Entry to Uppercase Letters Only:**<br/>`=EXACT(A2, UPPER(A2))`</li>
        <li>**Allow Entry Only If Total Budget Is Not Exceeded:**<br/>`=SUM($B$2:$B$50)<=1000000`</li>
      </ul>

      ### 13.3 Dependent (Cascading) Dropdown Lists

      Jab Category select hone par Sub-Category Dropdown dynamically badalna chahiye (e.g., Country → City):


      <ol>
        <li>Sub-categories ke cell ranges select karke Named Ranges banao (e.g., Named Range `India` = Delhi, Mumbai, Pune; Named Range `USA` = NY, LA, Chicago). Note: Named range ka naam exact Category value jaisa hona chahiye!</li>
        <li>1st Dropdown (Category) in Cell A2: Normal List of Countries.</li>
        <li>2nd Dropdown (City) in Cell B2: Select List, Source text box me ye formula dalo:<br/>
        <div class="code-snippet">=INDIRECT(A2)</div></li>
      </ol>
    

---

## Module 14: Conditional Formatting Master Class
**Category:** Visual Formatting | **Description:** Highlight Rules, Top/Bottom, Data Bars, Color Scales, Icon Sets, Rule Hierarchy & Custom Formula Rules


      ### 14.1 Custom Formula Rules in Conditional Formatting

      Home Tab → Conditional Formatting → New Rule → **"Use a formula to determine which cells to format"**.


      
      <table class="notes-table">
        <thead>
          <tr>
            <th>Desired Visualization Goal</th>
            <th>Custom Formula Logic</th>
            <th>Applies To Range</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Highlight entire row if Status is "Pending"</td>
            <td>`=$C2="Pending"`</td>
            <td>`=$A$2:$G$100`</td>
          </tr>
          <tr>
            <td>Highlight rows with Overdue Deadlines</td>
            <td>`=AND($D2<TODAY(), $E2<>"Completed")`</td>
            <td>`=$A$2:$E$100`</td>
          </tr>
          <tr>
            <td>Highlight Duplicate rows as they are typed</td>
            <td>`=COUNTIF($A$2:$A2, $A2)>1`</td>
            <td>`=$A$2:$A$500`</td>
          </tr>
          <tr>
            <td>Alternate Row Zebra Striping (Light Grey)</td>
            <td>`=MOD(ROW(), 2)=0`</td>
            <td>`=$A$2:$Z$1000`</td>
          </tr>
          <tr>
            <td>Highlight Top 10% Salaries in Dept</td>
            <td>`=$B2>=PERCENTILE($B$2:$B$100, 0.9)`</td>
            <td>`=$A$2:$B$100`</td>
          </tr>
        </tbody>
      </table>

      <div class="note-box">
        **⚠️ Critical Rule for Whole-Row Formatting:** Formula me Column reference lock karna mandatory hai (`$C2`), par Row reference un-lock (`2`) rakhna chahiye!
      </div>
    

---

## Module 15: PivotTables & PivotCharts Master Class
**Category:** Data Summarization | **Description:** Creating PivotTables, Value Field Settings (% of Total, Running Total), Date Grouping, Calculated Fields, Slicers & Timelines


      ### 15.1 PivotTable Fundamentals & Creation

      **Insert Tab → PivotTable** (Shortcut: `Alt + N + V + T`). Raw data Table (Ctrl+T) format me hona chahiye taaki future records auto-include ho sakenge.



      ### 15.2 Value Field Settings Capabilities

      <ul>
        <li>**Summarize Values By:** Sum, Count, Average, Max, Min, Product, StdDev.</li>
        <li>**Show Values As:**
          <ul>
            <li>`% of Grand Total` → Each category's percentage share in overall revenue.</li>
            <li>`% of Column / Row Total` → Contribution within sub-categories.</li>
            <li>`Running Total In` → Cumulative cumulative total month-over-month.</li>
            <li>`Difference From / % Difference From` → Variance against base target/previous month.</li>
            <li>`Rank Smallest to Largest / Largest to Smallest` → Dynamic ranking inside Pivot.</li>
          </ul>
        </li>
      </ul>

      ### 15.3 Date Grouping & Calculated Fields

      <ul>
        <li>**Grouping Dates:** PivotTable me Date field pe Right Click → **Group** → Select `Years`, `Quarters`, `Months`.</li>
        <li>**Calculated Field:** PivotTable Analyze Tab → Fields, Items & Sets → Calculated Field. Formula: `Profit_Margin = (Sales - Cost) / Sales`.</li>
        <li>**Slicers & Timelines:** Interactive clickable filter buttons. Single Slicer ko multiple PivotTables se connect karne ke liye Slicer par Right Click → **Report Connections** → Check all PivotTables!</li>
      </ul>
    

---

## Module 16: Advanced Visual Charts & Dynamic Dashboards
**Category:** Data Visualization | **Description:** Column, Bar, Line, Pie, Combo Charts with Secondary Axis, Waterfall, Funnel, Sparklines & Interactive Form Controls


      ### 16.1 Chart Selection Guide for Executives

      <table class="notes-table">
        <thead>
          <tr>
            <th>Business Scenario</th>
            <th>Recommended Chart Type</th>
            <th>Why it Works Best</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Comparing Revenue across 5 Regions</td>
            <td>**Clustered Column / Bar Chart**</td>
            <td>Clear visual height comparison between discrete categories.</td>
          </tr>
          <tr>
            <td>Monthly Revenue Trend over 3 Years</td>
            <td>**Line Chart with Markers**</td>
            <td>Displays continuous continuous trend and seasonality clearly.</td>
          </tr>
          <tr>
            <td>Revenue (Bar) vs Profit Margin % (Line)</td>
            <td>**Combo Chart (Secondary Axis)**</td>
            <td>Two different scales (Currency & %) combined on single plot area.</td>
          </tr>
          <tr>
            <td>Financial P&L Breakdown (Gross → Net)</td>
            <td>**Waterfall Chart**</td>
            <td>Shows step-by-step additions and deductions impact on final total.</td>
          </tr>
          <tr>
            <td>Sales Funnel Pipeline Stages</td>
            <td>**Funnel Chart**</td>
            <td>Visualizes drop-offs at each conversion stage.</td>
          </tr>
          <tr>
            <td>In-cell Mini Trend Display</td>
            <td>**Sparklines (Line / Column)**</td>
            <td>Insert → Sparklines. Tiny charts inside table cells!</td>
          </tr>
        </tbody>
      </table>
    

---

## Module 17: What-If Analysis & Optimization Solver
**Category:** Decision Modeling | **Description:** Goal Seek, One-Variable & Two-Variable Data Tables, Scenario Manager & Optimization Solver Add-in


      ### 17.1 What-If Analysis Tools Blueprint

      <ul>
        <li>**Goal Seek (Data → What-If Analysis → Goal Seek):** Jab tumhe desired RESULT pata ho aur input variable reverse-calculate karna ho. (e.g., Net Profit ₹1,00,000 lane ke liye sales volume kitna chahiye?).</li>
        <li>**Data Tables (1-Variable & 2-Variable):** Sensitivity analysis matrix. Ek ya do variables badalne par final profit kitna badlega uski complete summary table banata hai.</li>
        <li>**Scenario Manager:** Best Case, Base Case aur Worst Case scenarios save karke comparative Summary Report generate karna.</li>
        <li>**Solver Add-in (File → Options → Add-ins → Solver):** Linear Programming Optimization tool! Target cell ko Maximize/Minimize karna under specific business constraints.</li>
      </ul>
    

---

## Module 18: Power Query ETL (Data Cleaning & Transformation)
**Category:** Data Engineering | **Description:** Importing from Files/Folders/Web, Applied Steps, Split, Merge Queries (Joins), Append Queries (Union), Unpivot & M Language


      ### 18.1 What is Power Query ETL?

      Power Query Excel ka internal ETL (Extract, Transform, Load) engine hai. Data → Get Data. Automated multi-source data cleaning tool jo bina formulas ke large messy datasets ko format karta hai. **One-time setup, then just press REFRESH!**



      ### 18.2 Core Power Query Operations

      <ul>
        <li>**Merge Queries (Joins):** VLOOKUP ka powerful replacement! Two tables ko primary key (e.g. EmpID) ke basis pe join karna (Left Outer, Right Outer, Full Outer, Inner, Left Anti, Right Anti).</li>
        <li>**Append Queries (Union):** Multiple monthly files/tables ko single master table me vertical stack (Combine) karna.</li>
        <li>**Unpivot Columns:** Wide cross-tab tables ko clean normalized long database format me transform karna.</li>
        <li>**Applied Steps Pane:** Har cleaning step record hota hai. Undo, re-order, edit ya M-Code view karna via Advanced Editor.</li>
      </ul>
    

---

## Module 19: Power Pivot & Data Model (DAX Master Class)
**Category:** Big Data & Modeling | **Description:** Data Model Star Schema, Relationships, Measures vs Calculated Columns, DAX Formulas (CALCULATE, SUMX, RELATED, Time Intelligence)


      ### 19.1 Enabling Power Pivot & Building Data Model

      File → Options → Add-ins → COM Add-ins → Check **Microsoft Power Pivot for Excel**. Single workbook me **Millions of Rows (10 Million+)** without Excel slowing down handle kar sakta hai!



      ### 19.2 DAX (Data Analysis Expressions) Reference Table

      <table class="notes-table">
        <thead>
          <tr>
            <th>DAX Function</th>
            <th>Type</th>
            <th>Description & DAX Code Formula</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**CALCULATE**</td>
            <td>Measure</td>
            <td>**DAX ka sabse powerful function!** Calculation ka filter context alter/override karta hai. `Total_West_Sales := CALCULATE(SUM(Sales[Amount]), Sales[Region]="West")`</td>
          </tr>
          <tr>
            <td>**SUMX**</td>
            <td>Iterator Measure</td>
            <td>Row-by-row calculation karke final sum return karta hai. `Total_Revenue := SUMX(Sales, Sales[Qty] * Sales[UnitPrice])`</td>
          </tr>
          <tr>
            <td>**RELATED**</td>
            <td>Calculated Column</td>
            <td>Related dimension table se matching value pull karta hai (VLOOKUP equivalent in DAX). `Category := RELATED(Products[CategoryName])`</td>
          </tr>
          <tr>
            <td>**DISTINCTCOUNT**</td>
            <td>Measure</td>
            <td>Unique values count karta hai. `Unique_Customers := DISTINCTCOUNT(Sales[CustomerID])`</td>
          </tr>
          <tr>
            <td>**SAMEPERIODLASTYEAR**</td>
            <td>Time Intelligence</td>
            <td>Prior year same period ka revenue calculate karke Year-over-Year (YoY) Growth batata hai.</td>
          </tr>
        </tbody>
      </table>
    

---

## Module 20: Macros & VBA Automation Fundamentals
**Category:** Automation Engine | **Description:** Developer Tab, VBE Editor (Alt+F11), Recording Macros, Sub/Function Procedures, Variables, Loops & Conditional VBA


      ### 20.1 Recording & Running Your First Macro

      Developer Tab → Record Macro → Perform Repetitive Formatting → Stop Recording. Assign macro to a custom button or shortcut key (e.g., Ctrl+Shift+M). Workbook ko **Excel Macro-Enabled Workbook (.xlsm)** format me save karna compulsory hai!



      ### 20.2 VBA Editor (Alt + F11) Code Structure

      <div class="code-snippet">
<span class="comment">' VBA Macro to Highlight High Sales Records in Red</span>
Sub HighlightHighSales()
    Dim cell As Range
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    For Each cell In Range("B2:B" & lastRow)
        If cell.Value > 50000 Then
            cell.Interior.Color = RGB(144, 238, 144) <span class="comment">' Light Green</span>
            cell.Font.Bold = True
        End If
    Next cell
    
    MsgBox "Macro Execution Completed Successfully!", vbInformation, "Automation Done"
End Sub
      </div>
    

---

## Module 21: Protection, Security & Collaboration
**Category:** Security | **Description:** Protect Sheet, Protect Workbook Structure, Unlocking Specific Cells, Encrypt File Password & Track Changes


      ### 21.1 Cell Level Protection Setup

      <ol>
        <li>By default, Excel ke sabhi cells "Locked" state me hote hain (lekin protection OFF hota hai).</li>
        <li>Jus cells ko input ke liye OPEN (Editable) rakhna hai, unhe select karke **Ctrl + 1 → Protection → Uncheck 'Locked'**.</li>
        <li>Review Tab → **Protect Sheet** → Password set karo. Ab locked cells read-only ho jayenge, aur sirf unlocked cells me entry ho sakegi!</li>
      </ol>
      
      ### 21.2 Workbook Structure & File Passwords

      <ul>
        <li>**Protect Workbook Structure:** Sheets ko delete, rename, hide ya add karne se rokta hai.</li>
        <li>**Encrypt with Password (File Password):** File → Info → Protect Workbook → Encrypt with Password. File bina password ke open hi nahi hogi.</li>
      </ul>
    

---

## Module 22: 100+ Essential Keyboard Shortcuts Master List
**Category:** Productivity | **Description:** Categorized Shortcuts for Navigation, Selection, Formatting, Editing, Formulas, View & Tables


      <table class="notes-table">
        <thead>
          <tr>
            <th>Shortcut Key</th>
            <th>Category</th>
            <th>Exact Action Performed</th>
          </tr>
        </thead>
        <tbody>
          <tr><td>`Ctrl + T`</td><td>Table</td><td>Convert selected range to Excel Structured Table</td></tr>
          <tr><td>`Ctrl + Shift + L`</td><td>Filter</td><td>Toggle AutoFilter On / Off</td></tr>
          <tr><td>`Ctrl + E`</td><td>Flash Fill</td><td>Instant pattern auto-fill data extraction</td></tr>
          <tr><td>`Ctrl + 1`</td><td>Formatting</td><td>Open Format Cells dialog box</td></tr>
          <tr><td>`Alt + =`</td><td>Formulas</td><td>Instant AutoSum formula insertion</td></tr>
          <tr><td>`F4`</td><td>Editing</td><td>Toggle Cell Reference ($A$1/$A1/A$1) or Repeat last action</td></tr>
          <tr><td>`F2`</td><td>Editing</td><td>Enter cell edit mode</td></tr>
          <tr><td>`Ctrl + Shift + ~`</td><td>Formatting</td><td>Apply General Number format</td></tr>
          <tr><td>`Ctrl + Shift + $`</td><td>Formatting</td><td>Apply Currency format</td></tr>
          <tr><td>`Ctrl + Shift + %`</td><td>Formatting</td><td>Apply Percentage format</td></tr>
          <tr><td>`Ctrl + Arrow Keys`</td><td>Navigation</td><td>Jump to edge of data region</td></tr>
          <tr><td>`Ctrl + Shift + Arrow`</td><td>Selection</td><td>Select entire range to edge of data region</td></tr>
          <tr><td>`Ctrl + ;`</td><td>Date</td><td>Insert Current System Date</td></tr>
          <tr><td>`Ctrl + Shift + ;`</td><td>Time</td><td>Insert Current System Time</td></tr>
          <tr><td>`Alt + F11`</td><td>Developer</td><td>Open Visual Basic Editor (VBE)</td></tr>
          <tr><td>`Ctrl + Alt + F5`</td><td>Data</td><td>Refresh All Data Sources & PivotTables</td></tr>
          <tr><td>`Ctrl + ``</td><td>Auditing</td><td>Show formulas instead of values toggle</td></tr>
          <tr><td>`Alt + H + O + I`</td><td>Formatting</td><td>Auto-fit Column Widths automatically</td></tr>
        </tbody>
      </table>
    

---

## Module 23: Complete Excel Error Troubleshooting Cheat Sheet
**Category:** Debugging | **Description:** Diagnostic blueprint and resolution steps for #N/A, #REF!, #DIV/0!, #VALUE!, #NAME?, ####, #SPILL!, #CALC!


      <table class="notes-table">
        <thead>
          <tr>
            <th>Error Code</th>
            <th>Root Cause Diagnostic</th>
            <th>Exact Resolution Steps</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>**#N/A**</td>
            <td>Value Not Available. Lookup function ko target item search range me nahi mila.</td>
            <td>Check trailing spaces via TRIM(), verify exact match vs approx match, or wrap in `IFNA()` / `IFERROR()`.</td>
          </tr>
          <tr>
            <td>**#REF!**</td>
            <td>Invalid Cell Reference. Formula jiss cell/row/column ko refer kar raha tha wo delete ho gaya hai.</td>
            <td>Undo (Ctrl+Z) delete action, or update formula references to valid ranges.</td>
          </tr>
          <tr>
            <td>**#DIV/0!**</td>
            <td>Division by Zero. Formula zero ya rami (empty) cell se divide kar raha hai.</td>
            <td>Wrap in `=IFERROR(A1/B1, 0)` or `=IF(B1=0, 0, A1/B1)`.</td>
          </tr>
          <tr>
            <td>**#VALUE!**</td>
            <td>Wrong Data Type. Math formula me Text string use ho raha hai.</td>
            <td>Check data types, remove text characters from numeric cells using `VALUE()` or `NUMBERVALUE()`.</td>
          </tr>
          <tr>
            <td>**#NAME?**</td>
            <td>Function Name Misspelled ya invalid Range Name.</td>
            <td>Correct formula spelling (e.g., `VLOOKUP` instead of `VLOOKP`), verify Named Ranges in Name Manager.</td>
          </tr>
          <tr>
            <td>**####**</td>
            <td>Column width is too narrow to display the formatted number/date, or negative date value.</td>
            <td>Double click column header boundary to auto-fit width (Shortcut: `Alt + H + O + I`).</td>
          </tr>
          <tr>
            <td>**#SPILL!**</td>
            <td>Dynamic Array formula output path is blocked by non-empty cells.</td>
            <td>Clear all data cells in the spill range path below/right of the formula.</td>
          </tr>
        </tbody>
      </table>
    

---

## Module 24: 100 MIS Data Analyst Practical Real-World Case Studies
**Category:** Job & Interview Prep | **Description:** Complete solved scenarios: VLOOKUP vs XLOOKUP, Data Cleaning, Running Totals, Ageing Analysis, Year-over-Year Growth, Dynamic Dashboards


      ### 24.1 Top Real-World Business Case Studies

      
      #### Case Study 1: Combine First & Last Name with Proper Casing

      **Problem:** Column A contains lower case first names ("lokesh"), Column B contains last names ("ahire"). Clean & join them into Full Name.


      <div class="code-snippet">=PROPER(A2) & " " & PROPER(B2)</div>

      #### Case Study 2: Extract Username & Domain from Email Address

      **Username:** `=TEXTBEFORE(A2, "@")` or `=LEFT(A2, FIND("@", A2)-1)`


      **Domain:** `=TEXTAFTER(A2, "@")` or `=MID(A2, FIND("@", A2)+1, LEN(A2))`



      #### Case Study 3: Running Total Sales Calculation

      **Problem:** Column C me Daily Sales hain. Column D me Cumulative Running Total calculate karna hai.


      <div class="code-snippet">=SUM($C$2:C2)  <span class="comment">-- Copy/drag down column D</span></div>

      #### Case Study 4: Employee Salary Increment Calculation

      **Rules:** Performance Rating "A" → 15%, "B" → 10%, "C" → 5% increment.


      <div class="code-snippet">=Salary * (1 + SWITCH(Rating, "A", 0.15, "B", 0.10, "C", 0.05, 0))</div>

      #### Case Study 5: Debtors Invoice Ageing Bracket Analysis

      **Brackets:** 0-30 Days ("Current"), 31-60 Days ("30+ Days"), 61-90 Days ("60+ Days"), >90 Days ("Critical Overdue").


      <div class="code-snippet">
=IFS(
   (TODAY()-DueDate)<=30, "Current",
   (TODAY()-DueDate)<=60, "30+ Days",
   (TODAY()-DueDate)<=90, "60+ Days",
   TRUE, "Critical Overdue"
)
      </div>
    

---

