# Excel Cheat Sheet

> Part of the **35-module Excel Mastery** curriculum. See the [curriculum roadmap](README.md) or the other [quick reference guides](README.md#quick-reference-guides).

One-page memory aids for functions, formats, errors and recipes. For explanations go to the lesson folders.

_Function names and minimum versions checked on November 7, 2026._

**Version key:** `2016+` works in Excel 2016 and later (and all newer versions). `2019+`, `2021+` need that version or newer. `365 / 2024` needs Microsoft 365 or Excel 2024. `365 only` needs Microsoft 365 (not in Excel 2024).

Jump to: [Functions](#functions-by-category) · [Which function?](#which-function-do-i-use) · [References and operators](#references-and-operators) · [Criteria](#criteria-for-sumifs-countifs-and-friends) · [Number formats](#number-format-codes) · [Date and time codes](#date-and-time-format-codes) · [Errors](#error-codes-and-fixes) · [Recipes](#copy-ready-recipes) · [Versions](#what-needs-which-version)

## Functions by category

### Math and totals

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **SUM** | `SUM(number1, [number2], ...)` | Adds numbers | `=SUM(B2:B10)` | 2016+ |
| **ROUND** | `ROUND(number, num_digits)` | Rounds to a set number of decimals | `=ROUND(3.14159,2)` | 2016+ |
| **ROUNDUP** | `ROUNDUP(number, num_digits)` | Always rounds away from zero | `=ROUNDUP(2.1,0)` | 2016+ |
| **ROUNDDOWN** | `ROUNDDOWN(number, num_digits)` | Always rounds toward zero | `=ROUNDDOWN(2.9,0)` | 2016+ |
| **INT** | `INT(number)` | Rounds down to the nearest whole number | `=INT(7.8)` | 2016+ |
| **MOD** | `MOD(number, divisor)` | Remainder after division | `=MOD(10,3)` | 2016+ |
| **ABS** | `ABS(number)` | Removes the minus sign | `=ABS(-5)` | 2016+ |
| **POWER** | `POWER(number, power)` | Raises a number to a power | `=POWER(2,3)` | 2016+ |
| **SQRT** | `SQRT(number)` | Square root | `=SQRT(81)` | 2016+ |
| **SUMIF** | `SUMIF(range, criteria, [sum_range])` | Adds cells that meet one rule | `=SUMIF(C2:C9,"North",J2:J9)` | 2016+ |
| **SUMIFS** | `SUMIFS(sum_range, criteria_range1, criteria1, ...)` | Adds cells that meet many rules | `=SUMIFS(J:J,C:C,"North",F:F,"Hardware")` | 2016+ |
| **COUNTIF** | `COUNTIF(range, criteria)` | Counts cells that meet one rule | `=COUNTIF(E2:E9,"Fail")` | 2016+ |
| **COUNTIFS** | `COUNTIFS(criteria_range1, criteria1, ...)` | Counts cells that meet many rules | `=COUNTIFS(C:C,"North",F:F,"Software")` | 2016+ |
| **AVERAGEIF** | `AVERAGEIF(range, criteria, [average_range])` | Average of cells that meet one rule | `=AVERAGEIF(C:C,"East",J:J)` | 2016+ |
| **AVERAGEIFS** | `AVERAGEIFS(average_range, criteria_range1, criteria1, ...)` | Average with many rules | `=AVERAGEIFS(F:F,C:C,"Sales",G:G,">3")` | 2016+ |
| **MAXIFS** | `MAXIFS(max_range, criteria_range1, criteria1, ...)` | Largest value that meets rules | `=MAXIFS(J:J,C:C,"West")` | 2019+ |
| **MINIFS** | `MINIFS(min_range, criteria_range1, criteria1, ...)` | Smallest value that meets rules | `=MINIFS(J:J,C:C,"West")` | 2019+ |
| **SUMPRODUCT** | `SUMPRODUCT(array1, [array2], ...)` | Multiplies arrays and adds the results | `=SUMPRODUCT(G2:G9,H2:H9)` | 2016+ |
| **SUBTOTAL** | `SUBTOTAL(function_num, ref1, ...)` | Totals that ignore other subtotals and filtered rows | `=SUBTOTAL(109,J2:J41)` | 2016+ |
| **AGGREGATE** | `AGGREGATE(function_num, options, ref1, ...)` | Totals that can skip errors and hidden rows | `=AGGREGATE(9,6,J2:J41)` | 2016+ |

### Statistics

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **AVERAGE** | `AVERAGE(number1, [number2], ...)` | Mean of numbers | `=AVERAGE(B2:B10)` | 2016+ |
| **COUNT** | `COUNT(value1, [value2], ...)` | Counts numbers | `=COUNT(B2:B10)` | 2016+ |
| **COUNTA** | `COUNTA(value1, [value2], ...)` | Counts non-empty cells | `=COUNTA(A2:A10)` | 2016+ |
| **COUNTBLANK** | `COUNTBLANK(range)` | Counts empty cells | `=COUNTBLANK(H2:H31)` | 2016+ |
| **MAX** | `MAX(number1, [number2], ...)` | Largest number | `=MAX(B2:B10)` | 2016+ |
| **MIN** | `MIN(number1, [number2], ...)` | Smallest number | `=MIN(B2:B10)` | 2016+ |
| **LARGE** | `LARGE(array, k)` | k-th largest value | `=LARGE(J2:J41,3)` | 2016+ |
| **SMALL** | `SMALL(array, k)` | k-th smallest value | `=SMALL(J2:J41,3)` | 2016+ |
| **MEDIAN** | `MEDIAN(number1, [number2], ...)` | Middle value | `=MEDIAN(F2:F25)` | 2016+ |
| **MODE.SNGL** | `MODE.SNGL(number1, [number2], ...)` | Most frequent number | `=MODE.SNGL(G2:G25)` | 2016+ |
| **STDEV.S** | `STDEV.S(number1, [number2], ...)` | Spread of a sample | `=STDEV.S(C2:C21)` | 2016+ |
| **STDEV.P** | `STDEV.P(number1, [number2], ...)` | Spread of a whole population | `=STDEV.P(C2:C21)` | 2016+ |
| **VAR.S** | `VAR.S(number1, [number2], ...)` | Variance of a sample | `=VAR.S(C2:C21)` | 2016+ |
| **PERCENTILE.INC** | `PERCENTILE.INC(array, k)` | Value at a percentile (0 to 1) | `=PERCENTILE.INC(C2:C21,0.9)` | 2016+ |
| **QUARTILE.INC** | `QUARTILE.INC(array, quart)` | Quartile value (0 to 4) | `=QUARTILE.INC(C2:C21,1)` | 2016+ |
| **RANK.EQ** | `RANK.EQ(number, ref, [order])` | Rank of a number in a list | `=RANK.EQ(C2,C$2:C$21)` | 2016+ |
| **CORREL** | `CORREL(array1, array2)` | Correlation between two lists (-1 to 1) | `=CORREL(B2:B21,C2:C21)` | 2016+ |
| **SLOPE** | `SLOPE(known_y's, known_x's)` | Slope of the best-fit line | `=SLOPE(C2:C21,B2:B21)` | 2016+ |
| **INTERCEPT** | `INTERCEPT(known_y's, known_x's)` | Intercept of the best-fit line | `=INTERCEPT(C2:C21,B2:B21)` | 2016+ |
| **RSQ** | `RSQ(known_y's, known_x's)` | R-squared of the best-fit line | `=RSQ(C2:C21,B2:B21)` | 2016+ |
| **FORECAST.LINEAR** | `FORECAST.LINEAR(x, known_y's, known_x's)` | Predicts a value on a straight line | `=FORECAST.LINEAR(8,C2:C21,B2:B21)` | 2016+ |
| **T.TEST** | `T.TEST(array1, array2, tails, type)` | p-value of a t-test | `=T.TEST(A2:A9,B2:B9,2,2)` | 2016+ |

### Logical and error handling

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **IF** | `IF(logical_test, value_if_true, [value_if_false])` | Chooses between two results | `=IF(G2>=4,"Good","Review")` | 2016+ |
| **IFS** | `IFS(logical_test1, value_if_true1, ...)` | Many conditions, no nesting | `=IFS(G2>=4,"A",G2>=3,"B",TRUE,"C")` | 2019+ |
| **AND** | `AND(logical1, [logical2], ...)` | TRUE only if all are TRUE | `=AND(A2>0,B2>0)` | 2016+ |
| **OR** | `OR(logical1, [logical2], ...)` | TRUE if any is TRUE | `=OR(A2>0,B2>0)` | 2016+ |
| **NOT** | `NOT(logical)` | Flips TRUE and FALSE | `=NOT(A2>0)` | 2016+ |
| **XOR** | `XOR(logical1, [logical2], ...)` | TRUE if an odd number are TRUE | `=XOR(A2>0,B2>0)` | 2016+ |
| **SWITCH** | `SWITCH(expression, value1, result1, ..., [default])` | Matches one value to many results | `=SWITCH(C2,"North","N","South","S","Other")` | 2019+ |
| **IFERROR** | `IFERROR(value, value_if_error)` | Replaces any error with a safe value | `=IFERROR(A2/B2,0)` | 2016+ |
| **IFNA** | `IFNA(value, value_if_na)` | Replaces only #N/A | `=IFNA(VLOOKUP(A2,F:G,2,0),"Not found")` | 2016+ |

### Lookup and reference

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **VLOOKUP** | `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])` | Looks right in a table | `=VLOOKUP(A2,F:H,3,FALSE)` | 2016+ |
| **HLOOKUP** | `HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])` | Looks down a table with headings on top | `=HLOOKUP("Q1",A1:E4,2,FALSE)` | 2016+ |
| **INDEX** | `INDEX(array, row_num, [column_num])` | Returns the value at a position | `=INDEX(B2:B9,3)` | 2016+ |
| **MATCH** | `MATCH(lookup_value, lookup_array, [match_type])` | Finds a position | `=MATCH("Asha",D2:D41,0)` | 2016+ |
| **XLOOKUP** | `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Modern lookup in any direction | `=XLOOKUP(A2,F:F,H:H,"Not found")` | 2021+ |
| **XMATCH** | `XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])` | Modern MATCH | `=XMATCH("Asha",D2:D41)` | 2021+ |
| **CHOOSE** | `CHOOSE(index_num, value1, [value2], ...)` | Picks the n-th value from a list | `=CHOOSE(2,"Low","Mid","High")` | 2016+ |
| **ROW** | `ROW([reference])` | Row number | `=ROW(A5)` | 2016+ |
| **COLUMN** | `COLUMN([reference])` | Column number | `=COLUMN(C1)` | 2016+ |
| **ROWS** | `ROWS(array)` | Counts rows in a range | `=ROWS(A2:A41)` | 2016+ |
| **TRANSPOSE** | `TRANSPOSE(array)` | Swaps rows and columns | `=TRANSPOSE(A1:C1)` | 2016+ |

### Text

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **LEFT** | `LEFT(text, [num_chars])` | First characters | `=LEFT(A2,3)` | 2016+ |
| **RIGHT** | `RIGHT(text, [num_chars])` | Last characters | `=RIGHT(A2,4)` | 2016+ |
| **MID** | `MID(text, start_num, num_chars)` | Characters from the middle | `=MID(A2,5,2)` | 2016+ |
| **LEN** | `LEN(text)` | Number of characters | `=LEN(A2)` | 2016+ |
| **TRIM** | `TRIM(text)` | Removes extra spaces | `=TRIM(A2)` | 2016+ |
| **CLEAN** | `CLEAN(text)` | Removes non-printing characters | `=CLEAN(A2)` | 2016+ |
| **UPPER** | `UPPER(text)` | CAPITAL LETTERS | `=UPPER(A2)` | 2016+ |
| **LOWER** | `LOWER(text)` | small letters | `=LOWER(A2)` | 2016+ |
| **PROPER** | `PROPER(text)` | Capitalises Each Word | `=PROPER(A2)` | 2016+ |
| **FIND** | `FIND(find_text, within_text, [start_num])` | Position of text (case-sensitive) | `=FIND("@",A2)` | 2016+ |
| **SEARCH** | `SEARCH(find_text, within_text, [start_num])` | Position of text (not case-sensitive, allows wildcards) | `=SEARCH("sales",A2)` | 2016+ |
| **SUBSTITUTE** | `SUBSTITUTE(text, old_text, new_text, [instance_num])` | Replaces matching text | `=SUBSTITUTE(A2,"-","")` | 2016+ |
| **REPLACE** | `REPLACE(old_text, start_num, num_chars, new_text)` | Replaces by position | `=REPLACE(A2,1,3,"XYZ")` | 2016+ |
| **TEXT** | `TEXT(value, format_text)` | Turns a number into formatted text | `=TEXT(B2,"dd-mmm-yyyy")` | 2016+ |
| **VALUE** | `VALUE(text)` | Turns number-looking text into a number | `=VALUE("123")` | 2016+ |
| **NUMBERVALUE** | `NUMBERVALUE(text, [decimal_separator], [group_separator])` | Text to number, any locale | `=NUMBERVALUE("1.234,5",",",".")` | 2016+ |
| **CONCATENATE** | `CONCATENATE(text1, [text2], ...)` | Joins text (older function) | `=CONCATENATE(A2," ",B2)` | 2016+ |
| **CONCAT** | `CONCAT(text1, [text2], ...)` | Joins text and ranges | `=CONCAT(A2:C2)` | 2019+ |
| **TEXTJOIN** | `TEXTJOIN(delimiter, ignore_empty, text1, ...)` | Joins with a separator | `=TEXTJOIN(", ",TRUE,A2:A6)` | 2019+ |
| **REPT** | `REPT(text, number_times)` | Repeats text | `=REPT("\|",5)` | 2016+ |
| **EXACT** | `EXACT(text1, text2)` | Case-sensitive comparison | `=EXACT(A2,B2)` | 2016+ |
| **TEXTBEFORE** | `TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Text before a delimiter | `=TEXTBEFORE(A2,"@")` | 365 / 2024 |
| **TEXTAFTER** | `TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])` | Text after a delimiter | `=TEXTAFTER(A2,"@")` | 365 / 2024 |
| **TEXTSPLIT** | `TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])` | Splits text into cells | `=TEXTSPLIT(A2," ")` | 365 / 2024 |

### Date and time

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **TODAY** | `TODAY()` | Today's date | `=TODAY()` | 2016+ |
| **NOW** | `NOW()` | Date and time now | `=NOW()` | 2016+ |
| **DATE** | `DATE(year, month, day)` | Builds a date | `=DATE(2026,11,7)` | 2016+ |
| **YEAR** | `YEAR(serial_number)` | Year of a date | `=YEAR(B2)` | 2016+ |
| **MONTH** | `MONTH(serial_number)` | Month of a date | `=MONTH(B2)` | 2016+ |
| **DAY** | `DAY(serial_number)` | Day of a date | `=DAY(B2)` | 2016+ |
| **WEEKDAY** | `WEEKDAY(serial_number, [return_type])` | Day of the week as a number | `=WEEKDAY(B2,2)` | 2016+ |
| **WEEKNUM** | `WEEKNUM(serial_number, [return_type])` | Week of the year | `=WEEKNUM(B2)` | 2016+ |
| **EOMONTH** | `EOMONTH(start_date, months)` | Last day of a month | `=EOMONTH(B2,0)` | 2016+ |
| **EDATE** | `EDATE(start_date, months)` | Same day, some months later | `=EDATE(B2,3)` | 2016+ |
| **DATEDIF** | `DATEDIF(start_date, end_date, unit)` | Years, months or days between dates (works but is not listed in Excel's function list) | `=DATEDIF(E2,TODAY(),"Y")` | 2016+ |
| **NETWORKDAYS** | `NETWORKDAYS(start_date, end_date, [holidays])` | Working days between dates | `=NETWORKDAYS(B2,C2)` | 2016+ |
| **WORKDAY** | `WORKDAY(start_date, days, [holidays])` | Date after n working days | `=WORKDAY(B2,10)` | 2016+ |
| **DAYS** | `DAYS(end_date, start_date)` | Days between two dates | `=DAYS(C2,B2)` | 2016+ |
| **YEARFRAC** | `YEARFRAC(start_date, end_date, [basis])` | Fraction of a year between dates | `=YEARFRAC(B2,C2)` | 2016+ |
| **TIME** | `TIME(hour, minute, second)` | Builds a time | `=TIME(9,30,0)` | 2016+ |
| **HOUR** | `HOUR(serial_number)` | Hour of a time | `=HOUR(B2)` | 2016+ |

### Dynamic arrays, LET and LAMBDA

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **FILTER** | `FILTER(array, include, [if_empty])` | Returns the rows that meet a rule | `=FILTER(A2:J41,C2:C41="North")` | 2021+ |
| **SORT** | `SORT(array, [sort_index], [sort_order], [by_col])` | Sorts a range with a formula | `=SORT(A2:J41,10,-1)` | 2021+ |
| **SORTBY** | `SORTBY(array, by_array1, [sort_order1], ...)` | Sorts by another range | `=SORTBY(A2:A41,J2:J41,-1)` | 2021+ |
| **UNIQUE** | `UNIQUE(array, [by_col], [exactly_once])` | Removes duplicates | `=UNIQUE(C2:C41)` | 2021+ |
| **SEQUENCE** | `SEQUENCE(rows, [columns], [start], [step])` | Makes a list of numbers | `=SEQUENCE(5,1,1,2)` | 2021+ |
| **RANDARRAY** | `RANDARRAY([rows], [columns], [min], [max], [whole_number])` | Random numbers | `=RANDARRAY(3,1,1,6,TRUE)` | 2021+ |
| **LET** | `LET(name1, value1, [name2, value2], ..., calculation)` | Names a value inside a formula | `=LET(x,5,x*2)` | 2021+ |
| **LAMBDA** | `LAMBDA([parameter1, ...], calculation)` | Builds your own function | `=LAMBDA(x,x*2)(5)` | 365 / 2024 |
| **MAP** | `MAP(array1, lambda)` | Runs a LAMBDA on each item | `=MAP(A2:A6,LAMBDA(v,v*2))` | 365 / 2024 |
| **REDUCE** | `REDUCE([initial_value], array, lambda)` | Rolls a list into one value | `=REDUCE(0,A2:A6,LAMBDA(a,v,a+v))` | 365 / 2024 |
| **SCAN** | `SCAN([initial_value], array, lambda)` | Running total with a LAMBDA | `=SCAN(0,A2:A6,LAMBDA(a,v,a+v))` | 365 / 2024 |
| **BYROW** | `BYROW(array, lambda)` | Runs a LAMBDA on each row | `=BYROW(A2:C6,LAMBDA(r,SUM(r)))` | 365 / 2024 |
| **BYCOL** | `BYCOL(array, lambda)` | Runs a LAMBDA on each column | `=BYCOL(A2:C6,LAMBDA(c,SUM(c)))` | 365 / 2024 |
| **MAKEARRAY** | `MAKEARRAY(rows, columns, lambda)` | Builds an array with a LAMBDA | `=MAKEARRAY(3,3,LAMBDA(r,c,r*c))` | 365 / 2024 |
| **VSTACK** | `VSTACK(array1, [array2], ...)` | Stacks ranges on top of each other | `=VSTACK(A2:B3,A6:B7)` | 365 / 2024 |
| **HSTACK** | `HSTACK(array1, [array2], ...)` | Puts ranges side by side | `=HSTACK(A2:A5,C2:C5)` | 365 / 2024 |
| **TAKE** | `TAKE(array, rows, [columns])` | Takes the first or last rows or columns | `=TAKE(A2:C20,5)` | 365 / 2024 |
| **DROP** | `DROP(array, rows, [columns])` | Drops rows or columns | `=DROP(A1:C20,1)` | 365 / 2024 |
| **CHOOSECOLS** | `CHOOSECOLS(array, col_num1, [col_num2], ...)` | Picks columns | `=CHOOSECOLS(A2:J41,1,10)` | 365 / 2024 |
| **CHOOSEROWS** | `CHOOSEROWS(array, row_num1, [row_num2], ...)` | Picks rows | `=CHOOSEROWS(A2:J41,1,-1)` | 365 / 2024 |
| **TOCOL** | `TOCOL(array, [ignore], [scan_by_column])` | Turns a block into one column | `=TOCOL(A2:C6)` | 365 / 2024 |
| **TOROW** | `TOROW(array, [ignore], [scan_by_column])` | Turns a block into one row | `=TOROW(A2:C6)` | 365 / 2024 |
| **EXPAND** | `EXPAND(array, rows, [columns], [pad_with])` | Pads an array to a size | `=EXPAND(A2:B3,4,3,"")` | 365 / 2024 |
| **GROUPBY** | `GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship])` | Summary table with one formula | `=GROUPBY(C2:C41,J2:J41,SUM)` | 365 only |
| **PIVOTBY** | `PIVOTBY(row_fields, col_fields, values, function, ...)` | Pivot-style summary with one formula | `=PIVOTBY(C2:C41,F2:F41,J2:J41,SUM)` | 365 only |
| **PERCENTOF** | `PERCENTOF(data_subset, data_all)` | Share of a total | `=GROUPBY(C2:C41,J2:J41,PERCENTOF)` | 365 only |

### Financial

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **PMT** | `PMT(rate, nper, pv, [fv], [type])` | Loan payment per period | `=PMT(0.08/12,60,-25000)` | 2016+ |
| **IPMT** | `IPMT(rate, per, nper, pv, [fv], [type])` | Interest part of a payment | `=IPMT(0.08/12,1,60,-25000)` | 2016+ |
| **PPMT** | `PPMT(rate, per, nper, pv, [fv], [type])` | Principal part of a payment | `=PPMT(0.08/12,1,60,-25000)` | 2016+ |
| **FV** | `FV(rate, nper, pmt, [pv], [type])` | Future value of savings | `=FV(0.06/12,120,-200)` | 2016+ |
| **PV** | `PV(rate, nper, pmt, [fv], [type])` | Today's value of future payments | `=PV(0.06/12,60,-500)` | 2016+ |
| **NPV** | `NPV(rate, value1, [value2], ...)` | Net present value of later cash flows | `=NPV(0.1,B3:B7)+B2` | 2016+ |
| **IRR** | `IRR(values, [guess])` | Return rate that makes NPV zero | `=IRR(B2:B7)` | 2016+ |
| **XNPV** | `XNPV(rate, values, dates)` | NPV with exact dates | `=XNPV(0.1,B2:B7,A2:A7)` | 2016+ |
| **XIRR** | `XIRR(values, dates, [guess])` | IRR with exact dates | `=XIRR(B2:B7,A2:A7)` | 2016+ |
| **RATE** | `RATE(nper, pmt, pv, [fv], [type], [guess])` | Interest rate per period | `=RATE(60,-500,25000)` | 2016+ |
| **NPER** | `NPER(rate, pmt, pv, [fv], [type])` | Number of periods | `=NPER(0.08/12,-500,25000)` | 2016+ |
| **SLN** | `SLN(cost, salvage, life)` | Straight-line depreciation | `=SLN(10000,1000,5)` | 2016+ |
| **MIRR** | `MIRR(values, finance_rate, reinvest_rate)` | IRR with a reinvestment rate | `=MIRR(B2:B7,0.08,0.06)` | 2016+ |

### Information and other

| Function | Syntax | What it does | Example | Needs |
|---|---|---|---|---|
| **ISNUMBER** | `ISNUMBER(value)` | TRUE if a number | `=ISNUMBER(A2)` | 2016+ |
| **ISTEXT** | `ISTEXT(value)` | TRUE if text | `=ISTEXT(A2)` | 2016+ |
| **ISBLANK** | `ISBLANK(value)` | TRUE if empty | `=ISBLANK(A2)` | 2016+ |
| **ISERROR** | `ISERROR(value)` | TRUE for any error | `=ISERROR(A2)` | 2016+ |
| **ISNA** | `ISNA(value)` | TRUE for #N/A only | `=ISNA(A2)` | 2016+ |
| **ISFORMULA** | `ISFORMULA(reference)` | TRUE if the cell holds a formula | `=ISFORMULA(A2)` | 2016+ |
| **FORMULATEXT** | `FORMULATEXT(reference)` | Shows a cell's formula as text | `=FORMULATEXT(B2)` | 2016+ |
| **HYPERLINK** | `HYPERLINK(link_location, [friendly_name])` | Makes a clickable link | `=HYPERLINK("https://example.com","Open")` | 2016+ |
| **IMAGE** | `IMAGE(source, [alt_text], [sizing], [height], [width])` | Shows a picture from a web address | `=IMAGE("https://example.com/logo.png","Logo")` | 365 / 2024 |

## Which function do I use?

| I want to... | Use |
|---|---|
| Add numbers | SUM |
| Add numbers that meet rules | SUMIFS |
| Count cells that meet rules | COUNTIFS |
| Average numbers that meet rules | AVERAGEIFS |
| Get the biggest or smallest that meets a rule | MAXIFS, MINIFS (2019+) |
| Choose between two or more results | IF, IFS (2019+), SWITCH (2019+) |
| Find a value in another table | XLOOKUP (2021+), INDEX with MATCH, VLOOKUP |
| Hide an error | IFERROR, IFNA |
| Clean text | TRIM, CLEAN, PROPER, UPPER, LOWER, SUBSTITUTE |
| Cut out part of text | LEFT, MID, RIGHT, TEXTBEFORE, TEXTAFTER |
| Join text | &, CONCAT (2019+), TEXTJOIN (2019+) |
| Split text into cells | Text to Columns, TEXTSPLIT (365 / 2024) |
| Build or split a date | DATE, YEAR, MONTH, DAY |
| Move a date | EDATE, EOMONTH, WORKDAY |
| Measure the gap between dates | DATEDIF, DAYS, NETWORKDAYS |
| Round | ROUND, ROUNDUP, ROUNDDOWN |
| Rank or take the k-th value | RANK.EQ, LARGE, SMALL |
| Filter, sort, de-duplicate with a formula | FILTER, SORT, UNIQUE (2021+) |
| Summarise like a pivot table | PivotTable, GROUPBY and PIVOTBY (365 only) |
| Name a value inside a formula | LET (2021+) |
| Predict a value | FORECAST.LINEAR, TREND |
| Loan payment | PMT, IPMT, PPMT |
| Value of an investment | NPV, IRR, XNPV, XIRR |

## References and operators

| Item | Meaning |
|---|---|
| `A1` | Relative reference (moves when copied) |
| `$A$1` | Absolute reference (never moves) |
| `A$1`, `$A1` | Mixed references (row or column locked) |
| `A1:B5` | Range |
| `Sheet2!A1` | Cell on another sheet |
| `tblSales[Revenue]` | Table column (structured reference) |
| `+ - * / ^` | Add, subtract, multiply, divide, power |
| `&` | Join text |
| `= < > <= >= <>` | Equal, less, greater, less or equal, greater or equal, not equal |
| `%` | Percent (`=50%` is 0.5) |
| `( )` | Brackets are calculated first |
| `A1#` | The whole spill range of a dynamic array (2021+) |

Order of calculation: brackets, powers, multiply and divide, add and subtract, then join (&), then comparisons.

## Criteria for SUMIFS, COUNTIFS and friends

| You want | Criteria |
|---|---|
| Equal to text | `"North"` |
| Not equal to | `"<>North"` |
| Greater than a number | `">10"` |
| Greater than or equal to | `">=10"` |
| Compare to a cell | `">"&B1` |
| Date on or after | `">="&DATE(2026,1,1)` |
| Starts with | `"A*"` |
| Ends with | `"*Plan"` |
| Contains | `"*soft*"` |
| Any one character | `"?at"` |
| Blank cells | `""` |
| Not blank | `"<>"` |
| A real star or question mark | `"~*"` and `"~?"` |

## Number format codes

| Code | Shows | Example |
|---|---|---|
| `General` | Excel's default | 1234.5 |
| `0` | Whole number | 1235 |
| `0.00` | Two decimals | 1234.50 |
| `#,##0` | Thousands separator | 1,235 |
| `#,##0.00` | Thousands and two decimals | 1,234.50 |
| `0%` and `0.0%` | Percent | 25% and 25.6% |
| `$#,##0.00` | Dollar amount | $1,234.50 |
| `#,##0,"K"` | Thousands as K | 1,235K |
| `00000` | Leading zeros | 00123 |
| `0.0E+00` | Scientific | 1.2E+03 |
| `@` | Text | keeps leading zeros |
| `;;;` | Hide the value | (shows nothing) |

## Date and time format codes

| Code | Shows | Example |
|---|---|---|
| `d`, `dd` | Day number | 7, 07 |
| `ddd`, `dddd` | Day name | Sat, Saturday |
| `m`, `mm` | Month number | 11, 11 |
| `mmm`, `mmmm` | Month name | Nov, November |
| `yy`, `yyyy` | Year | 26, 2026 |
| `yyyy-mm-dd` | ISO date | 2026-11-07 |
| `h`, `hh`, `mm`, `ss` | Hours, minutes, seconds | 9, 09, 05, 30 |
| `h:mm AM/PM` | 12-hour clock | 9:05 PM |
| `[h]:mm` | Hours above 24 | 27:30 |

(`mm` means minutes when it follows `h` and means month otherwise.)

## Error codes and fixes

| Error | Meaning | Usual fix |
|---|---|---|
| `#DIV/0!` | Divided by zero or an empty cell | Test the divisor with IF, or use IFERROR |
| `#N/A` | A lookup found nothing | Check spaces and types, then IFNA |
| `#NAME?` | Unknown function or name | Fix the spelling, add quotes around text, check the Excel version |
| `#NULL!` | Wrong range operator (a space) | Use a colon or a comma |
| `#NUM!` | Impossible number | Check the arguments, such as a negative square root |
| `#REF!` | A referenced cell was deleted | Undo or rebuild the reference |
| `#VALUE!` | Wrong type of value | Convert text to numbers with VALUE; check the range sizes |
| `#SPILL!` | A dynamic array has no room | Clear the cells in the spill range |
| `#CALC!` | A dynamic array calculation failed | Check LAMBDA inputs and empty arrays |
| `#####` | Not an error: the column is too narrow | Widen the column |

## Copy-ready recipes

Replace the ranges and names with yours.

| Task | Formula |
|---|---|
| Total by one rule | `=SUMIFS(sum_range, criteria_range, "North")` |
| Total by two rules | `=SUMIFS(sum_range, range1, "North", range2, "Hardware")` |
| Percent of total | `=B2/SUM($B$2:$B$20)` |
| Running total | `=SUM($B$2:B2)` |
| Growth percent | `=(new/old)-1` |
| Weighted average | `=SUMPRODUCT(values, weights)/SUM(weights)` |
| Safe lookup | `=IFERROR(VLOOKUP(A2, table, 2, FALSE), "Not found")` |
| Look left | `=INDEX(return_col, MATCH(key, key_col, 0))` |
| Two-way lookup | `=INDEX(data, MATCH(row_key, row_keys, 0), MATCH(col_key, col_keys, 0))` |
| Clean a name | `=PROPER(TRIM(A2))` |
| First name | `=LEFT(TRIM(A2), FIND(" ", TRIM(A2)&" ")-1)` |
| Years of service | `=DATEDIF(start, end, "Y")` |
| Last day of the month | `=EOMONTH(A2, 0)` |
| Quarter number | `=ROUNDUP(MONTH(A2)/3, 0)` |
| Rank | `=RANK.EQ(B2, $B$2:$B$20)` |
| Count unique (older Excel) | `=SUMPRODUCT(1/COUNTIF(A2:A100, A2:A100))` |
| Count unique (2021+) | `=COUNTA(UNIQUE(A2:A100))` |
| Loan payment per month | `=PMT(rate/12, years*12, -amount)` |
| NPV with year 0 | `=NPV(rate, year1_to_n) + year0` |

## What needs which version


| Version needed | Functions |
|---|---|
| Excel 2016 and later (everything else in this list) | 113 functions such as SUM, IF, SUMIFS, VLOOKUP, INDEX, MATCH, TEXT, DATE, PMT, NPV |
| Excel 2019 and later | MAXIFS, MINIFS, IFS, SWITCH, CONCAT, TEXTJOIN |
| Excel 2021 and later | XLOOKUP, XMATCH, FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, LET |
| Microsoft 365 and Excel 2024 | TEXTBEFORE, TEXTAFTER, TEXTSPLIT, LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS, CHOOSEROWS, TOCOL, TOROW, EXPAND, IMAGE |
| Microsoft 365 only | GROUPBY, PIVOTBY, PERCENTOF |

Note: Excel 2024 includes LAMBDA, the text and array functions (TEXTSPLIT, VSTACK, TAKE and others) but not GROUPBY, PIVOTBY or PERCENTOF.
