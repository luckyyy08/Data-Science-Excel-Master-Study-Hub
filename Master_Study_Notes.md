# 📊 Data Science & Advanced Excel Master Study Notes

Welcome to the **Complete Study Notes & Cheat Sheet** compiled from your Excel study folder (`c:\Users\Lucky\Desktop\Advance Excel`). This master document bridges **Excel Data Analytics** with **Python (Pandas/NumPy) & SQL** for Data Science learners.

---

## 🔰 Module 0: Excel Basics & Core Fundamentals

### 1. Excel Screen Layout & Cell Address
* **Row & Column Structure:** Columns are labeled with letters (`A, B, C...`) and Rows with numbers (`1, 2, 3...`).
* **Cell Address:** Intersection of Column and Row (e.g., `A1`, `B5`, `C15`).
* **Name Box:** Displays the current active cell address.
* **Formula Bar:** Displays the formula or value stored in the active cell.

### 2. Basic Data Types
| Data Type | Screen Alignment | Description & Example |
| :--- | :--- | :--- |
| **Text** | Left-Aligned | Names, labels, strings (`"Rahul"`, `"Sales"`). |
| **Number** | Right-Aligned | Numeric values used for calculation (`15000`, `99.5`). |
| **Date** | Right-Aligned | Stored internally as serial numbers starting from `1 Jan 1900` = `1`. |
| **Boolean** | Centered / System | Logical values (`TRUE` or `FALSE`). |

> [!NOTE]
> If a number is aligned to the **LEFT**, it is formatted as **Text**, which will prevent arithmetic calculations!

### 3. Basic Arithmetic Operators
* **Addition (`+`):** `=A1 + B1`
* **Subtraction (`-`):** `=A1 - B1`
* **Multiplication (`*`):** `=A1 * B1`
* **Division (`/`):** `=A1 / B1`
* **Percentage (`%`):** `=A1 * 10%`
* **Exponentiation (`^`):** `=A1 ^ 2` (Square of A1)

### 4. Formula vs. Function
* **Formula:** A custom mathematical expression created by the user (e.g., `=A1 + A2 + A3`).
* **Function:** A built-in predefined command in Excel (e.g., `=SUM(A1:A3)`).
* **Rule:** Every formula or function **MUST** start with an equals sign (`=`).

### 5. Cell Referencing & Lock (`$ Sign`)
* **Relative Reference (`A1`):** Changes automatically when copied down or across cells.
* **Absolute Reference (`$A$1`):** Locked in both row and column using **`F4` key**. Does not change when copied.
* **Mixed Reference (`$A1` or `A$1`):** Locks only column or only row.

### 6. Basic Aggregation Functions
* **SUM:** `=SUM(A1:A10)` $\rightarrow$ Calculates total sum of numbers.
* **AVERAGE:** `=AVERAGE(A1:A10)` $\rightarrow$ Calculates arithmetic mean.
* **MIN:** `=MIN(A1:A10)` $\rightarrow$ Finds smallest value.
* **MAX:** `=MAX(A1:A10)` $\rightarrow$ Finds largest value.
* **COUNT:** `=COUNT(A1:A10)` $\rightarrow$ Counts cells containing **numeric** values only.

---

## 🧠 Module 1: Logical Functions & Error Handling

### 1. Basic IF Function
Evaluates a condition and returns one value if `TRUE` and another if `FALSE`.
* **Syntax:** `=IF(logical_test, value_if_true, value_if_false)`
* **Example:** `=IF(C2 > 15000, 1000, 0)` *(If Basic Salary > 15,000, Bonus = 1,000, else 0)*

### 2. Nested IF (Multi-Slab Decision Trees)
Chains multiple IF statements for multi-tiered rules.
* **Example (Designation Bonus):**
  `=IF(Desg="Manager", 5000, IF(Desg="Supervisor", 3000, 1000))`

### 3. Income Tax Slab Calculation (Real-World Case Study)
* **Slabs:**
  * Income $\le$ ₹2,50,000 $\rightarrow$ No Tax
  * ₹2,50,001 to ₹5,00,000 $\rightarrow$ 5%
  * Income > ₹5,00,000 $\rightarrow$ 20%
* **Formula:**
  `=IF(Income <= 250000, 0, IF(Income <= 500000, Income * 5%, Income * 20%))`

### 4. IF with AND / OR
* **AND (All conditions must be TRUE):**
  `=IF(AND(Gender="Female", Salary<50000), "Eligible for Gift", "Not Eligible")`
* **OR (At least one condition must be TRUE):**
  `=IF(OR(Dept="Sales", Dept="Marketing"), 25000, 10000)`

### 5. Error Handling (`IFERROR`)
Replaces error codes like `#DIV/0!`, `#N/A`, or `#VALUE!` with clean default values.
* **Syntax:** `=IFERROR(value, value_if_error)`
* **Example:** `=IFERROR(A2 / B2, 0)`

---

## 📊 Module 2: Counting & Aggregation Functions

### 1. SUMIF & SUMIFS
* **SUMIF (Single Criterion):** `=SUMIF(range, criteria, [sum_range])`
  * Example: `=SUMIF(Dept_Range, "Sales", Salary_Range)`
* **SUMIFS (Multiple Criteria):** `=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)`
  * Example: `=SUMIFS(Salary_Range, Dept_Range, "Sales", Region_Range, "North")`

### 2. COUNTIF & COUNTIFS
* **COUNTIF:** `=COUNTIF(range, criteria)` $\rightarrow$ `=COUNTIF(Type_Range, "Cash")`
* **COUNTIFS:** `=COUNTIFS(range1, criteria1, range2, criteria2)` $\rightarrow$ `=COUNTIFS(Type_Range, "Cash", Mode_Range, "Credit")`

### 3. AVERAGEIF & AVERAGEIFS
* **AVERAGEIF:** `=AVERAGEIF(Type_Range, "Online", Amount_Range)`
* **AVERAGEIFS:** `=AVERAGEIFS(Amount_Range, Region_Range, "West", Product_Range, "Laptop")`

### 4. COUNTA vs. COUNTBLANK
* **`COUNTA(range)`:** Counts all non-empty cells (numbers, text, symbols).
* **`COUNTBLANK(range)`:** Counts empty/blank cells in a dataset.

### 5. SUBTOTAL Function
Calculates summaries for filtered datasets while ignoring hidden/filtered rows.
* **Syntax:** `=SUBTOTAL(function_num, ref1, [ref2], ...)`
* **Common Function Codes:** `9` = SUM, `1` = AVERAGE, `3` = COUNTA, `2` = COUNT.

### 6. SUMPRODUCT Function
Multiplies corresponding components in two or more arrays and returns the sum of products.
* **Formula:** `=SUMPRODUCT(Array1, Array2)`
* **Example:** `=SUMPRODUCT(Quantity_Range, Rate_Range)` $\rightarrow$ *(Matrix Dot Product)*

---

## 🔤 Module 3: Text Engineering & Data Cleaning

### 1. Character Extraction (`LEFT`, `RIGHT`, `MID`)
* **`LEFT(text, num_chars)`:** `=LEFT("Mumbai-400086", 6)` $\rightarrow$ `"Mumbai"`
* **`RIGHT(text, num_chars)`:** `=RIGHT("Mumbai-400086", 6)` $\rightarrow$ `"400086"`
* **`MID(text, start_num, num_chars)`:** `=MID("Mumbai-400086", 8, 6)` $\rightarrow$ `"400086"`

### 2. Position Lookup & Substitution
* **`FIND(find_text, within_text)`:** Returns 1-based index of character (Case-sensitive).
  * Combination Example: `=MID(A2, FIND("-", A2) + 1, 6)`
* **`SUBSTITUTE(text, old_text, new_text)`:** Replaces specific string segments.
  * Example: `=SUBSTITUTE("9812-34-56-78", "-", "")` $\rightarrow$ `"9812345678"`

### 3. Cleaning & Case Formatting
* **`TRIM(text)`:** Removes leading, trailing, and extra spaces.
* **`LEN(text)`:** Returns total character length.
* **`UPPER(text)` / `LOWER(text)` / `PROPER(text)`:** Case transformations.
* **Flash Fill (`Ctrl + E`):** AI pattern recognition to auto-extract text without formulas.

---

## 🔍 Module 4: Lookup & Reference Analytics

### 1. VLOOKUP (Vertical Lookup)
Searches for a value in the leftmost column of a table and returns a value in the same row from a specified column.
* **Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
* **Example:** `=VLOOKUP(Emp_ID, A2:D100, 3, 0)` *(0 = Exact Match)*

### 2. HLOOKUP (Horizontal Lookup)
Searches for a value in the top row of a table.
* **Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, 0)`

### 3. INDEX + MATCH Combination (Dynamic Left Lookup)
Replaces VLOOKUP limitations by enabling lookups to the left and improving computational speed.
* **Formula:** `=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))`

---

## 📅 Module 5: Temporal Analytics & Date Functions

* **`TODAY()`:** Returns current system date (`31-07-2026`).
* **`NOW()`:** Returns current system date and timestamp.
* **`YEAR(date)` / `MONTH(date)` / `DAY(date)`:** Extracts date parts.
* **Age Calculation (`DATEDIF`):**
  * Years: `=DATEDIF(Birthdate, TODAY(), "Y")`
  * Months: `=DATEDIF(Birthdate, TODAY(), "YM")`
  * Formatted String: `=DATEDIF(A2, TODAY(), "Y") & " Years " & DATEDIF(A2, TODAY(), "YM") & " Months"`

---

## 💰 Module 6: Financial Modeling & Risk Analytics

### 1. Loan EMI Calculation (`PMT`)
Calculates monthly payment for a loan based on constant payments and interest rate.
* **Syntax:** `=PMT(rate / 12, nper_months, -loan_amount)`
* **Example:** `=PMT(10%/12, 36, -100000)` $\rightarrow$ Monthly EMI for ₹1,00,000 loan at 10% interest for 3 years.

### 2. Interest vs. Principal Components
* **`IPMT` (Interest Portion):** `=IPMT(10%/12, 1, 36, -100000)`
* **`PPMT` (Principal Portion):** `=PPMT(10%/12, 1, 36, -100000)`
* **Rule:** `IPMT + PPMT = PMT` (Total Monthly EMI).

### 3. Valuation Functions
* **`PV` (Present Value):** `=PV(rate, nper, pmt)`
* **`FV` (Future Value):** `=FV(rate, nper, pmt)`

---

## 🗃️ Module 7: Database Functions, Data Validation & Pivots

* **Data Validation:** Restricts cell entry (e.g., `Data -> Data Validation -> Whole Number -> 1 to 1000`).
* **Database Functions (`DSUM`, `DAVERAGE`, `DCOUNT`):** Queries database ranges using external criteria blocks.
  * Example: `=DSUM(Database_Range, "Amount", Criteria_Block)`
* **Pivot Tables:** Generates dynamic summaries, group-bys, and cross-tabulations (`Insert -> Pivot Table`).

---

## 🐍 Module 10: Excel to Python/Pandas Matrix

| Feature / Operation | Excel Formula / Action | Python (Pandas / NumPy) Equivalent |
| :--- | :--- | :--- |
| **Conditional IF** | `=IF(A2>15000, 1000, 0)` | `np.where(df['Salary'] > 15000, 1000, 0)` |
| **Sum by Group** | `=SUMIF(Dept, "Sales", Salary)` | `df[df['Dept']=='Sales']['Salary'].sum()` |
| **Multi-Group Sum** | `=SUMIFS(Salary, Dept, "Sales", Region, "North")` | `df.groupby(['Dept', 'Region'])['Salary'].sum()` |
| **VLOOKUP / Join** | `=VLOOKUP(A2, Table, 2, 0)` | `pd.merge(df1, df2, on='ID', how='left')` |
| **Pivot Summary** | Insert Pivot Table | `df.pivot_table(index='Region', values='Sales', aggfunc='sum')` |
| **Text Replace** | `=SUBSTITUTE(A2, "-", "")` | `df['Phone'].str.replace('-', '')` |
| **String Slice** | `=LEFT(A2, 5)` | `df['Code'].str[:5]` |

---

## ❓ Topic-wise Master Practice Q&A Bank

### Basics & Formatting
1. **Q:** What is the cell address of Column C and Row 15?  
   **Ans:** `C15`.
2. **Q:** How do you lock cell references while dragging a formula?  
   **Ans:** Press `F4` key to make it absolute (`$A$1`).

### Logical & Aggregations
3. **Q:** Write a formula for 10% bonus if Salary > 15000, else 0.  
   **Ans:** `=IF(C2>15000, C2*10%, 0)`.
4. **Q:** What is the difference between SUMIF and SUMIFS sum ranges?  
   **Ans:** `SUMIF` places sum range at the end, while `SUMIFS` places sum range at the beginning.
5. **Q:** Which function calculates sums on filtered data while ignoring hidden rows?  
   **Ans:** `=SUBTOTAL(9, Range)` or `=SUBTOTAL(109, Range)`.

### Lookups & Financials
6. **Q:** What argument is passed in VLOOKUP for an exact match?  
   **Ans:** `0` or `FALSE`.
7. **Q:** Write a formula to calculate monthly EMI for ₹1,00,000 at 10% annual interest for 36 months.  
   **Ans:** `=PMT(10%/12, 36, -100000)`.
