# Interview Questions and Answers

> Part of the **35-module Excel Mastery** curriculum. See the [curriculum roadmap](README.md) for live lesson links as each module publishes.

**206 questions with answers**, grouped by level. Every lesson contributes 5 questions (they also appear at the end of each lesson), and the last section adds scenario questions that interviewers ask in practical tests.

How to use this file: read the question, say your answer out loud, then open the answer and compare. Use the pattern **explain, example, pitfall**. Then practise the formula in the lesson's workbook.

Jump to: [Beginner](#-beginner-questions) · [Intermediate](#-intermediate-questions) · [Advanced](#-advanced-questions) · [Scenario questions](#-scenario-and-practical-questions)

## 🟢 Beginner questions

50 questions from 10 lessons.

### 01 Excel Basics and Interface

_Lesson 01_

<details>
<summary><b>Q1. What is the difference between a workbook and a worksheet?</b></summary>

A workbook is the whole Excel file. A worksheet is one tab (page) inside it. One workbook can hold many worksheets.

</details>

<details>
<summary><b>Q2. What is a cell reference?</b></summary>

The address of a cell: column letter plus row number, like C7. Formulas use it so results update when the data changes.

</details>

<details>
<summary><b>Q3. Why use cell references instead of typing numbers in a formula?</b></summary>

The formula stays correct when inputs change, you can copy it down, and other people can see where numbers come from.

</details>

<details>
<summary><b>Q4. What does `=2+3*4` return and why?</b></summary>

14. Excel multiplies before it adds, the same as normal maths. Brackets change the order.

</details>

<details>
<summary><b>Q5. How do you quickly total a column?</b></summary>

Select the cell below the numbers and press Alt + = (AutoSum), or type =SUM(range). You can also select the numbers and read the total on the status bar.

</details>

### 02 Data Entry and Number Formats

_Lesson 02_

<details>
<summary><b>Q6. What is the difference between a cell's value and its displayed text?</b></summary>

The value is what Excel stores and uses in formulas. The displayed text is what the number format shows. For example 2.456 can display as 2.5 but formulas still use 2.456.

</details>

<details>
<summary><b>Q7. How does Excel store dates?</b></summary>

As serial numbers that count days from a start date. That is why you can add and subtract dates. Times are fractions of a day.

</details>

<details>
<summary><b>Q8. Your numbers will not add up and they sit on the left of the cells. What is wrong?</b></summary>

They are stored as text. Convert them with Convert to Number, Text to Columns or VALUE, then SUM works.

</details>

<details>
<summary><b>Q9. When would you use the TEXT function?</b></summary>

When you need a number or date as a formatted label, for example in a report title or a joined sentence. Do not use it when you still need to calculate with the result.

</details>

<details>
<summary><b>Q10. How do you keep a leading zero in a code like 00123?</b></summary>

Format the cells as Text before typing, use a custom format 00000, or use TEXT(value,"00000") in a formula.

</details>

### 03 Formulas and Cell References

_Lesson 03_

<details>
<summary><b>Q11. What is the difference between a relative and an absolute reference?</b></summary>

A relative reference changes when you copy the formula. An absolute reference ($B$3) is locked and always points to the same cell.

</details>

<details>
<summary><b>Q12. What does the F4 key do while editing a formula?</b></summary>

It cycles the reference through $A$1, A$1, $A1 and A1 so you can lock the row, the column, both or neither.

</details>

<details>
<summary><b>Q13. When would you use a mixed reference?</b></summary>

When a formula must be copied in two directions, such as a price grid or multiplication table where one input lives in a left column and the other in a top row.

</details>

<details>
<summary><b>Q14. What are the benefits of named ranges?</b></summary>

Formulas are easier to read (=Revenue*TaxRate), names do not shift when copied, and one change to the name updates every formula that uses it.

</details>

<details>
<summary><b>Q15. A formula you copied down shows the same result in every row. What is the likely cause?</b></summary>

A reference that should be relative was locked with dollar signs, or the calculation mode is set to manual.

</details>

### 04 Logical Functions

_Lesson 04_

<details>
<summary><b>Q16. What does IF return when the logical test is TRUE?</b></summary>

The second argument (value_if_true). If the test is FALSE it returns the third argument (value_if_false).

</details>

<details>
<summary><b>Q17. What is the difference between AND and OR?</b></summary>

AND needs every condition to be TRUE. OR needs at least one condition to be TRUE.

</details>

<details>
<summary><b>Q18. Why use IFS instead of nested IF?</b></summary>

IFS is easier to read and has less brackets to match. It needs Excel 2019 or 365; in Excel 2016 you must nest IFs.

</details>

<details>
<summary><b>Q19. When is SWITCH better than IFS?</b></summary>

When you compare one value against a list of exact matches, such as department codes. SWITCH repeats the value only once.

</details>

<details>
<summary><b>Q20. How can you count the cells that meet a logical test without a helper column?</b></summary>

Use SUMPRODUCT with a double minus, for example =SUMPRODUCT(--(G2:G25>=4)). COUNTIF also works for simple tests.

</details>

### 06 Text Functions

_Lesson 06_

<details>
<summary><b>Q21. What is the difference between FIND and SEARCH?</b></summary>

FIND is case-sensitive and does not allow wildcards. SEARCH ignores case and allows ? and * wildcards. Both return the starting position of the text.

</details>

<details>
<summary><b>Q22. How do you split a full name into first and last name with older functions?</b></summary>

First name: =LEFT(A2,FIND(" ",A2)-1). Last name: =MID(A2,FIND(" ",A2)+1,100). Clean with TRIM first.

</details>

<details>
<summary><b>Q23. What does TRIM do and why does it matter?</b></summary>

It removes leading and trailing spaces and reduces repeated inner spaces to one. Hidden spaces make lookups and comparisons fail.

</details>

<details>
<summary><b>Q24. How do you remove all dashes from a phone number?</b></summary>

=SUBSTITUTE(A2,"-",""). Nest more SUBSTITUTE functions to remove other symbols.

</details>

<details>
<summary><b>Q25. How is TEXTJOIN better than the & operator?</b></summary>

It joins a whole range with one separator, can skip empty cells and is much shorter when you have many items. It needs Excel 2019 or later.

</details>

### 11 Error Handling

_Lesson 11_

<details>
<summary><b>Q26. What does #N/A usually mean and how do you handle it?</b></summary>

A lookup could not find the value. Check for spaces and data types first. Handle it with IFNA when a missing value is acceptable.

</details>

<details>
<summary><b>Q27. What is the difference between IFERROR and IFNA?</b></summary>

IFERROR catches every error type. IFNA catches only #N/A, so you still see real mistakes like #REF!.

</details>

<details>
<summary><b>Q28. When does #REF! appear?</b></summary>

When a formula refers to a cell, row, column or sheet that was deleted.

</details>

<details>
<summary><b>Q29. How do you total a range that includes errors?</b></summary>

Use AGGREGATE(9,6,range). Function 9 is SUM and option 6 ignores errors. Better still, fix the cause.

</details>

<details>
<summary><b>Q30. Why can IFERROR be dangerous?</b></summary>

It hides all errors, including real mistakes, so wrong totals look fine. Use it only where the error is expected and the replacement is sensible.

</details>

### 13 Sorting, Filtering and Tables

_Lesson 13_

<details>
<summary><b>Q31. What is the difference between sorting and filtering?</b></summary>

Sorting reorders all rows. Filtering hides rows that do not meet a rule, without deleting them.

</details>

<details>
<summary><b>Q32. What are the benefits of Excel Tables?</b></summary>

They grow automatically, fill formulas down, keep formatting, have built-in filters, support structured references and slicers, and feed pivot tables and Power Query cleanly.

</details>

<details>
<summary><b>Q33. Why does SUBTOTAL(109, ...) behave differently from SUM?</b></summary>

It ignores rows hidden by a filter (and manually hidden rows with 109), while SUM adds everything.

</details>

<details>
<summary><b>Q34. How do you sort by more than one column?</b></summary>

Data, Sort, Add Level. Put the main column first, then the next level, each with its own order.

</details>

<details>
<summary><b>Q35. What is a structured reference?</b></summary>

A reference that uses the table and column names, for example tblSales[Revenue]. It adjusts automatically when the table changes.

</details>

### 14 Conditional Formatting

_Lesson 14_

<details>
<summary><b>Q36. What is the difference between a built-in rule and a formula rule?</b></summary>

Built-in rules cover common cases (greater than, top 10, duplicates). A formula rule can test anything that gives TRUE or FALSE, including other columns.

</details>

<details>
<summary><b>Q37. How do you highlight a whole row based on one cell?</b></summary>

Select the whole range, create a formula rule that locks the column of the test cell, for example =$G2>=10.

</details>

<details>
<summary><b>Q38. Why does a formula rule use the first cell of the selection?</b></summary>

Excel applies the formula to the other cells by shifting relative references, as if you copied it. So the formula must be written for the top-left cell.

</details>

<details>
<summary><b>Q39. How do you remove or edit rules?</b></summary>

Home, Conditional Formatting, Manage Rules, then choose the scope (This Worksheet), and edit or delete a rule. Clear Rules removes them from selected cells or the sheet.

</details>

<details>
<summary><b>Q40. What are good practices for colour?</b></summary>

Use few colours, colour-blind-safe palettes, consistent meaning, a legend, and never rely on colour alone.

</details>

### 31 Excel for Managers and CEOs

_Lesson 31_

<details>
<summary><b>Q41. How do you choose KPIs for a team?</b></summary>

Start from the team's goals and decisions, pick three to six measures that people can influence, define each clearly and set a target and an owner. Include a leading indicator.

</details>

<details>
<summary><b>Q42. What would you check when a colleague sends you a model?</b></summary>

Sources and date, assumptions, whether totals reconcile, whether inputs are separated from formulas, hard-coded numbers and a quick sense check of the trend.

</details>

<details>
<summary><b>Q43. What is run-rate and what is its limit?</b></summary>

The annualised result from a recent period, such as half-year revenue times 2. It ignores seasonality and growth changes, so it is a rough guide.

</details>

<details>
<summary><b>Q44. Why compare KPIs with a target or last period?</b></summary>

A number alone has no meaning. Comparison shows whether performance is good or bad and which way it is moving.

</details>

<details>
<summary><b>Q45. What should an executive dashboard contain?</b></summary>

A few KPIs against target, one or two charts that explain the movement, key risks and the decisions needed, with the data date and owner.

</details>

### 35 Cheat Sheet and Shortcuts

_Lesson 35_

<details>
<summary><b>Q46. Name five Excel shortcuts you use every day.</b></summary>

For example F2 (edit), F4 (absolute reference), Ctrl + 1 (format cells), Alt + = (AutoSum) and Ctrl + Shift + L (filters).

</details>

<details>
<summary><b>Q47. What does F4 do while you edit a formula?</b></summary>

It cycles the selected reference through $A$1, A$1, $A1 and A1.

</details>

<details>
<summary><b>Q48. How do you quickly see all formulas on a sheet?</b></summary>

Press Ctrl + ` (the key left of the number 1) or use Formulas, Show Formulas.

</details>

<details>
<summary><b>Q49. Which function would you use to find a value in another table?</b></summary>

XLOOKUP in Excel 2021 or later, or INDEX with MATCH, or VLOOKUP with an exact match.

</details>

<details>
<summary><b>Q50. How do you learn shortcuts effectively?</b></summary>

Learn a few at a time, use them in real work, review them with spaced practice and keep a short list near your screen.

</details>

## 🟡 Intermediate questions

90 questions from 18 lessons.

### 05 Lookup Functions

_Lesson 05_

<details>
<summary><b>Q51. Why is FALSE (or 0) used in VLOOKUP?</b></summary>

It asks for an exact match. Without it Excel does an approximate match, which needs a sorted table and can return wrong results.

</details>

<details>
<summary><b>Q52. What are the limits of VLOOKUP?</b></summary>

It can only look to the right of the search column, it breaks when columns are inserted if the column number is typed, and by default it finds only the first match.

</details>

<details>
<summary><b>Q53. Why do many analysts prefer INDEX-MATCH to VLOOKUP?</b></summary>

It can look left, it does not depend on a column number, it is faster on large data and it still works in older Excel versions.

</details>

<details>
<summary><b>Q54. What does XLOOKUP add over VLOOKUP?</b></summary>

It looks in any direction, uses exact match by default, has a built-in if-not-found argument, can search from the bottom, and does not need a column number.

</details>

<details>
<summary><b>Q55. When would you use an approximate match?</b></summary>

For bands and slabs such as tax brackets, grades or commission tiers where you want the row with the largest start value not above the lookup value. The table must be sorted ascending.

</details>

### 07 Date and Time Functions

_Lesson 07_

<details>
<summary><b>Q56. How does Excel store dates and times?</b></summary>

Dates are whole numbers counting days. Times are fractions of a day (0.5 is noon). That is why you can add, subtract and compare them.

</details>

<details>
<summary><b>Q57. How do you calculate someone's age in complete years?</b></summary>

=DATEDIF(birth_date, reference_date, "Y"). Use TODAY() as the reference date only when you really want the age today.

</details>

<details>
<summary><b>Q58. What is the difference between EDATE and EOMONTH?</b></summary>

EDATE keeps the same day number and moves by months. EOMONTH returns the last day of the month you land on.

</details>

<details>
<summary><b>Q59. How do you find the due date that skips weekends and holidays?</b></summary>

=WORKDAY(start_date, days, holidays) where holidays is a range of dates.

</details>

<details>
<summary><b>Q60. Why is TODAY() called volatile?</b></summary>

It recalculates every time the workbook calculates, so results change daily and many volatile functions can slow big files.

</details>

### 08 Math and Statistical Functions

_Lesson 08_

<details>
<summary><b>Q61. What is the syntax of SUMIFS and what is different from SUMIF?</b></summary>

SUMIFS(sum_range, criteria_range1, criteria1, ...). The sum range comes first and you can use many rules. In SUMIF the sum range is the last argument.

</details>

<details>
<summary><b>Q62. How do you do an OR condition with SUMIFS?</b></summary>

Add two SUMIFS (one per option) or use SUMPRODUCT with a plus between the conditions.

</details>

<details>
<summary><b>Q63. What does SUMPRODUCT do?</b></summary>

It multiplies matching items of two or more lists and adds the results. With tests it can count and sum by many conditions without helper columns.

</details>

<details>
<summary><b>Q64. What is the difference between COUNT, COUNTA and COUNTBLANK?</b></summary>

COUNT counts numbers. COUNTA counts any non-empty cell. COUNTBLANK counts empty cells.

</details>

<details>
<summary><b>Q65. How do you sum values between two dates?</b></summary>

=SUMIFS(sum_range, date_range, ">="&start, date_range, "<="&end).

</details>

### 09 Dynamic Arrays

_Lesson 09_

<details>
<summary><b>Q66. What is a spill range?</b></summary>

The block of cells a dynamic array formula fills automatically. It resizes when the source data changes.

</details>

<details>
<summary><b>Q67. What causes #SPILL!?</b></summary>

Something blocks the spill range: a non-empty cell, a merged cell, a Table, or a result that is too big. Clear the obstacle.

</details>

<details>
<summary><b>Q68. How do you combine two conditions in FILTER?</b></summary>

Multiply the tests for AND, for example (C2:C41="North")*(G2:G41>=10). Add them for OR.

</details>

<details>
<summary><b>Q69. How does UNIQUE differ from Remove Duplicates?</b></summary>

UNIQUE is a formula and leaves the original data alone. It updates when data changes. Remove Duplicates deletes rows once and does not update.

</details>

<details>
<summary><b>Q70. Which Excel versions support FILTER and SORT?</b></summary>

Excel 2021, Excel 2024 and Microsoft 365. Not Excel 2019 or 2016.

</details>

### 12 Data Cleaning and Validation

_Lesson 12_

<details>
<summary><b>Q71. What are the first checks you do on a new data set?</b></summary>

Count rows and columns, look for blanks, duplicates, text in number columns, mixed formats and outliers. I also keep a raw copy and note every change.

</details>

<details>
<summary><b>Q72. How do you remove duplicates without losing data?</b></summary>

Copy the data to a cleaning sheet. Select the whole table, use Data, Remove Duplicates and choose the columns that make a row unique. Or flag them first with COUNTIF.

</details>

<details>
<summary><b>Q73. How do you convert numbers stored as text?</b></summary>

Use Convert to Number, Text to Columns, VALUE, or multiply by 1. For text with symbols, SUBSTITUTE first.

</details>

<details>
<summary><b>Q74. What types of data validation rules exist?</b></summary>

Any value, whole number, decimal, list, date, time, text length and custom formula.

</details>

<details>
<summary><b>Q75. Does data validation stop bad data completely?</b></summary>

No. It checks typed entries but not pasted values. Existing data is not checked either, except by Circle Invalid Data.

</details>

### 15 Pivot Tables, Slicers and Timelines

_Lesson 15_

<details>
<summary><b>Q76. What are the four areas of the PivotTable field list?</b></summary>

Filters, Columns, Rows and Values.

</details>

<details>
<summary><b>Q77. Why does a pivot table not show my new rows?</b></summary>

The pivot needs to be refreshed, and its source range may not include the new rows. Use a Table as the source and press Refresh.

</details>

<details>
<summary><b>Q78. What is a slicer and how is it different from a filter?</b></summary>

A slicer is a visual set of buttons that filters a pivot (or a table). It shows the current filter state on screen and can drive several pivots at once.

</details>

<details>
<summary><b>Q79. How do you group dates by month and year?</b></summary>

Right-click a date in the pivot, Group, then choose Months and Years (Excel does this automatically in newer versions when dates are added).

</details>

<details>
<summary><b>Q80. How can you check that a pivot table is correct?</b></summary>

Reconcile its grand total with the source data (SUM) and spot check one or two cells with SUMIFS.

</details>

### 16 Charts and Graphs

_Lesson 16_

<details>
<summary><b>Q81. When would you use a line chart and when a column chart?</b></summary>

A line chart for trends over many time points. A column chart for comparing a few categories or a few periods.

</details>

<details>
<summary><b>Q82. What is wrong with most pie charts?</b></summary>

People compare angles poorly. With many slices or close values it is hard to read. A sorted bar chart is clearer.

</details>

<details>
<summary><b>Q83. How do you show a relationship between two numeric variables?</b></summary>

A scatter chart with a trendline. Report the correlation or R-squared.

</details>

<details>
<summary><b>Q84. What is a waterfall chart used for?</b></summary>

To show how a starting value changes through increases and decreases to reach an end value, such as revenue to profit.

</details>

<details>
<summary><b>Q85. How do you make a chart update automatically when new data is added?</b></summary>

Base it on an Excel Table or a dynamic named range, so the series range grows with the data.

</details>

### 17 Dashboard Design

_Lesson 17_

<details>
<summary><b>Q86. What makes a good dashboard?</b></summary>

A clear purpose and audience, a few meaningful KPIs compared to targets, simple charts, a logical layout and filters, all driven by formulas so it refreshes easily.

</details>

<details>
<summary><b>Q87. How do you keep a dashboard updated with new data?</b></summary>

Base calculations on Tables or pivots, use Power Query for imports, and press Refresh All. The dashboard sheet only points to those results.

</details>

<details>
<summary><b>Q88. What is the difference between a report and a dashboard?</b></summary>

A report gives detail for study. A dashboard gives a quick view of the key numbers and trends on one screen.

</details>

<details>
<summary><b>Q89. How do slicers help a dashboard?</b></summary>

They let users filter tables, pivots and charts with a click and show the current selection clearly.

</details>

<details>
<summary><b>Q90. How do you choose which KPIs to show?</b></summary>

Start from the decisions the audience must make, pick measures that link to goals, and drop vanity numbers that change no decision.

</details>

### 18 Power Query

_Lesson 18_

<details>
<summary><b>Q91. What is Power Query and when would you use it instead of formulas?</b></summary>

A tool to import, clean and combine data using recorded steps. Use it for repeated clean-ups, large or messy sources, combining files and anything that must be refreshed.

</details>

<details>
<summary><b>Q92. What is the difference between Merge and Append?</b></summary>

Merge joins two tables side by side on a key (like a lookup or SQL JOIN). Append stacks tables with the same columns on top of each other (like UNION).

</details>

<details>
<summary><b>Q93. What does unpivot do?</b></summary>

It turns column headings into values in one column, giving a tidy table with one measure per row. It is needed for summarising data that came as Q1, Q2, Q3 columns.

</details>

<details>
<summary><b>Q94. How do you refresh a query?</b></summary>

Data, Refresh All, or Ctrl + Alt + F5. You can also set a query to refresh when the file opens or every few minutes.

</details>

<details>
<summary><b>Q95. What is M?</b></summary>

The functional language that Power Query writes for each step. You can read and edit it in the Advanced Editor.

</details>

### 20 What-If Analysis Tools

_Lesson 20_

<details>
<summary><b>Q96. What is the difference between Goal Seek and Solver?</b></summary>

Goal Seek changes one input to reach one target value. Solver can change many inputs, maximise or minimise a result and respect constraints.

</details>

<details>
<summary><b>Q97. When would you use a Data Table?</b></summary>

To see a result for many combinations of one or two inputs at once, such as a loan payment for different rates and years.

</details>

<details>
<summary><b>Q98. How is Scenario Manager different from a Data Table?</b></summary>

Scenario Manager saves named sets of several inputs (Worst, Base, Best) and compares them. A Data Table varies one or two inputs across a grid.

</details>

<details>
<summary><b>Q99. What do you need before you can run Solver?</b></summary>

A model with a result cell, changing cells and a formula connecting them, and the Solver add-in switched on.

</details>

<details>
<summary><b>Q100. What is break-even?</b></summary>

The volume where revenue equals total costs, so profit is zero. It equals fixed costs divided by the profit per unit.

</details>

### 22 Financial Functions and Models

_Lesson 22_

<details>
<summary><b>Q101. What does PMT return and what are its main arguments?</b></summary>

The payment per period for a loan. Arguments: rate per period, number of periods, present value (loan amount). Optional future value and payment type.

</details>

<details>
<summary><b>Q102. Why does NPV in Excel need the investment added separately?</b></summary>

NPV discounts all values as if they come at the end of each period starting with period 1. A cash flow today (year 0) is not discounted, so you add it outside NPV.

</details>

<details>
<summary><b>Q103. What is IRR?</b></summary>

The discount rate at which the NPV of the cash flows is zero. It is the project's rate of return.

</details>

<details>
<summary><b>Q104. How do you build a loan amortisation schedule?</b></summary>

Opening balance, then interest (balance x rate per period), principal (payment minus interest), and closing balance (opening minus principal). Repeat for each period.

</details>

<details>
<summary><b>Q105. What is the difference between IPMT and PPMT?</b></summary>

IPMT is the interest part of a payment in a given period. PPMT is the principal part. Together they equal the payment.

</details>

### 23 Business Analytics Cases

_Lesson 23_

<details>
<summary><b>Q106. How would you find the best marketing channel?</b></summary>

Compare cost per lead, conversion rate and ROI per channel, using totals. Check the volume, because a cheap channel may not scale.

</details>

<details>
<summary><b>Q107. How do you find the items to reorder?</b></summary>

Compare stock with the reorder level (SUMPRODUCT or a helper column), calculate the quantity needed and multiply by unit cost.

</details>

<details>
<summary><b>Q108. How do you calculate net margin?</b></summary>

Net profit divided by revenue, where net profit is revenue minus cost of goods sold minus operating expenses.

</details>

<details>
<summary><b>Q109. Why use totals when calculating a conversion rate across months?</b></summary>

Because averaging monthly rates gives equal weight to small and large months. Total customers divided by total leads weights each lead equally.

</details>

<details>
<summary><b>Q110. How do you present a result to a manager?</b></summary>

Start with the answer and the decision, then show the key numbers and a simple chart, and mention assumptions and risks in one or two lines.

</details>

### 24 Research and Survey Analysis

_Lesson 24_

<details>
<summary><b>Q111. How do you analyse a Likert scale in Excel?</b></summary>

Count each answer with COUNTIF, show the percentages, the mean and the top-2-box share, and chart it as a stacked bar. Compare groups with AVERAGEIFS or a pivot table.

</details>

<details>
<summary><b>Q112. How is NPS calculated?</b></summary>

Percentage of promoters (scores 9 and 10) minus percentage of detractors (scores 0 to 6). Passives (7 and 8) are ignored. The result ranges from -100 to 100.

</details>

<details>
<summary><b>Q113. What is a margin of error?</b></summary>

The range around a survey result that reflects sampling uncertainty. It gets smaller as the sample size increases.

</details>

<details>
<summary><b>Q114. What is reverse coding and when do you do it?</b></summary>

Flipping the scale of negatively worded items (for example 6 minus the score on a 1 to 5 scale) so that all items point in the same direction before combining them.

</details>

<details>
<summary><b>Q115. How would you clean survey data?</b></summary>

Remove duplicates and test rows, check for blanks, impossible values and straight-lining, recode answers and keep a log of every change.

</details>

### 25 Excel for Software Engineers and Testers

_Lesson 25_

<details>
<summary><b>Q116. How would you track test cases in Excel?</b></summary>

A Table with Test ID, module, priority, tester, status, run date and defect ID, with validation lists, conditional formatting and a pivot table or formulas for pass rate and failures by module.

</details>

<details>
<summary><b>Q117. How do you import JSON into Excel?</b></summary>

Data, Get Data, From File, From JSON. Power Query opens the data; expand the list and records, set types and Close and Load.

</details>

<details>
<summary><b>Q118. How can you call a REST API from Excel without code?</b></summary>

Use Power Query, Get Data, From Web, with the API address (and headers if needed). Then expand the JSON into columns and load it.

</details>

<details>
<summary><b>Q119. How do you compare two lists to find missing items?</b></summary>

Use COUNTIF or XLOOKUP in a helper column to flag items that are not found, or Power Query Merge with a Left Anti join.

</details>

<details>
<summary><b>Q120. When should an engineer stop using Excel and write code?</b></summary>

When the process runs repeatedly, the data is large, you need tests and version control, or multiple systems must be integrated.

</details>

### 28 Collaboration, Sharing, Protection and Security

_Lesson 28_

<details>
<summary><b>Q121. How do you let people edit inputs but not formulas?</b></summary>

Unlock the input cells (Format Cells, Protection), then protect the sheet (Review, Protect Sheet). Locked cells with formulas cannot be changed.

</details>

<details>
<summary><b>Q122. Is sheet protection secure?</b></summary>

No. It prevents accidental changes. Real protection comes from file encryption, access permissions and keeping sensitive data out of the file.

</details>

<details>
<summary><b>Q123. How does co-authoring work in Excel?</b></summary>

Save the file in OneDrive or SharePoint, share it, and several people edit it at the same time with AutoSave. Changes appear for everyone almost immediately.

</details>

<details>
<summary><b>Q124. How can you remove hidden personal information before sharing?</b></summary>

Use Inspect Document (File, Info, Check for Issues) to find comments, hidden sheets, personal properties and external links, then remove them.

</details>

<details>
<summary><b>Q125. What should you do if you forget a workbook's open password?</b></summary>

Nothing can be recovered if the file is properly encrypted. That is why you keep a copy of the password in a trusted password manager.

</details>

### 30 AI in Excel

_Lesson 30_

<details>
<summary><b>Q126. How can AI help in Excel?</b></summary>

It can draft formulas, explain existing ones, suggest charts and pivot tables, clean data steps, write macros and scripts and summarise results. It speeds up first drafts.

</details>

<details>
<summary><b>Q127. How do you verify AI-generated formulas?</b></summary>

Ask for the formula, test on a few rows by hand, reconcile totals with SUM or a pivot, check the range covers all rows, confirm functions exist in your version and test edge cases.

</details>

<details>
<summary><b>Q128. What are the main limits of AI in Excel?</b></summary>

It can be confidently wrong, may invent functions, can miss rows in large data, does not know business rules, may give inconsistent answers and raises privacy concerns.

</details>

<details>
<summary><b>Q129. What data should you never paste into an AI tool?</b></summary>

Anything confidential or personal that your organisation has not approved for that tool: customer details, salaries, health or financial account data, passwords and keys.

</details>

<details>
<summary><b>Q130. Why ask for the Excel version in a prompt?</b></summary>

Functions differ by version. XLOOKUP, FILTER and LAMBDA are not in older versions and will show #NAME?.

</details>

### 32 Excel vs Sheets, Power BI, SQL and Python

_Lesson 32_

<details>
<summary><b>Q131. When would you choose Power BI over Excel?</b></summary>

When you need interactive dashboards for many users, large refreshable data and central sharing. Excel is better for ad-hoc analysis and detailed models.

</details>

<details>
<summary><b>Q132. How does a SQL GROUP BY compare with an Excel PivotTable?</b></summary>

Both group rows by one or more fields and summarise values with SUM, COUNT or AVG. GROUP BY returns a table from a database and a PivotTable does the same in Excel.

</details>

<details>
<summary><b>Q133. What can Google Sheets do that Excel cannot, and the reverse?</b></summary>

Sheets is strong for real-time browser collaboration and has QUERY and IMPORTRANGE. Excel has richer charts, pivot tables, Power Query, Power Pivot, macros and works well offline.

</details>

<details>
<summary><b>Q134. Why would an analyst learn Python if they know Excel?</b></summary>

Python handles large data, repeats work automatically, offers statistics and machine learning libraries and works well with version control.

</details>

<details>
<summary><b>Q135. How do you decide which tool to use?</b></summary>

Look at data size, whether the job repeats, who consumes the result, whether it must be interactive, and the skills of your team.

</details>

### 34 Interview Preparation

_Lesson 34_

<details>
<summary><b>Q136. How would you prepare for an Excel practical test?</b></summary>

Practise the common tasks (lookups, SUMIFS, IF logic, pivot tables, cleaning, charts) on sample data with a timer, learn the key shortcuts and practise checking each result with a second method.

</details>

<details>
<summary><b>Q137. What do you do if you get stuck on a task?</b></summary>

Break it into smaller steps, test each part in a helper cell, say what you tried, move on and return if time allows.

</details>

<details>
<summary><b>Q138. How do you show that your work is accurate?</b></summary>

Reconcile totals, spot check rows by hand, handle errors and edge cases, keep inputs separate and document assumptions.

</details>

<details>
<summary><b>Q139. How do you explain a formula you wrote?</b></summary>

Say the goal, walk through the arguments from the inside out, give an example row and mention a limit or risk.

</details>

<details>
<summary><b>Q140. Which Excel skills matter most for an analyst?</b></summary>

Lookups, SUMIFS and COUNTIFS, pivot tables, data cleaning, charts, dates and text functions, and clear communication of results. Power Query and a little Python or SQL set you apart.

</details>

## 🔴 Advanced questions

35 questions from 7 lessons.

### 10 LET, LAMBDA and Array Shaping

_Lesson 10_

<details>
<summary><b>Q141. What are the benefits of LET?</b></summary>

It names intermediate results, so formulas are shorter and easier to read, and Excel calculates a repeated expression only once.

</details>

<details>
<summary><b>Q142. What is LAMBDA and how is it different from a VBA function?</b></summary>

LAMBDA lets you create a custom function from a formula, saved as a name. It needs no macros, so the file can stay .xlsx and nothing is blocked by macro security.

</details>

<details>
<summary><b>Q143. What does REDUCE do?</b></summary>

It starts with an initial value and applies a LAMBDA to each item of a list, carrying the result forward, for example a running total, and returns the final value.

</details>

<details>
<summary><b>Q144. How do TEXTSPLIT and TEXTBEFORE improve on LEFT, FIND and MID?</b></summary>

They split or cut text directly at a delimiter with one function, so you no longer need to find the position first.

</details>

<details>
<summary><b>Q145. Which versions support GROUPBY?</b></summary>

Microsoft 365 only. Microsoft lists it as applying to Excel for Microsoft 365, and it is not part of Excel 2021 or Excel 2024.

</details>

### 19 Power Pivot, Data Model and DAX

_Lesson 19_

<details>
<summary><b>Q146. What is the Data Model?</b></summary>

A hidden database in the workbook that holds multiple tables and relationships. Pivot tables can use it to summarise across tables.

</details>

<details>
<summary><b>Q147. What is the difference between a measure and a calculated column?</b></summary>

A calculated column is evaluated row by row when the data is loaded and stored. A measure is evaluated on demand using the current filters, so it is used for totals, ratios and KPIs.

</details>

<details>
<summary><b>Q148. What does CALCULATE do?</b></summary>

It evaluates an expression after modifying the filter context, for example the total revenue for the West only.

</details>

<details>
<summary><b>Q149. What is a star schema?</b></summary>

One fact table (transactions) linked to several dimension tables (products, customers, dates) with one-to-many relationships.

</details>

<details>
<summary><b>Q150. When would you use Power Pivot instead of normal pivot tables?</b></summary>

When data has several tables, many rows (beyond a worksheet), or needs measures such as year-to-date, ratios and distinct counts.

</details>

### 21 Statistics in Excel

_Lesson 21_

<details>
<summary><b>Q151. What is the difference between mean and median, and when would you use the median?</b></summary>

The mean is the average; the median is the middle value. Use the median when data is skewed or has outliers, such as salaries.

</details>

<details>
<summary><b>Q152. What does a correlation of 0.95 tell you?</b></summary>

A very strong positive straight-line relationship. It does not prove one variable causes the other.

</details>

<details>
<summary><b>Q153. What does R-squared mean?</b></summary>

The share of the variation in y that the model explains. 0.90 means 90 percent.

</details>

<details>
<summary><b>Q154. How do you read a p-value of 0.03 in a t-test?</b></summary>

If there were really no difference, a result this extreme would happen about 3 percent of the time, so it is evidence of a difference. It does not say how big or important the difference is.

</details>

<details>
<summary><b>Q155. When do you use ANOVA instead of a t-test?</b></summary>

When you compare the means of three or more groups. Repeated t-tests raise the chance of a false positive.

</details>

### 26 Macros and VBA

_Lesson 26_

<details>
<summary><b>Q156. What is the difference between recording a macro and writing VBA?</b></summary>

Recording captures your clicks as code with no thinking. Writing VBA lets you add loops, decisions, error handling and flexibility. Many people record first, then clean the code.

</details>

<details>
<summary><b>Q157. How do you find the last used row in column A in VBA?</b></summary>

ws.Cells(ws.Rows.Count, "A").End(xlUp).Row. It starts at the bottom and moves up to the first filled cell.

</details>

<details>
<summary><b>Q158. Why are macros a security risk?</b></summary>

A macro can run any code the user can run, including deleting files or sending data. Only enable macros from trusted sources and keep the Trust Center defaults.

</details>

<details>
<summary><b>Q159. What file type do you use for workbooks with macros?</b></summary>

.xlsm (macro-enabled workbook). A normal .xlsx cannot store macros.

</details>

<details>
<summary><b>Q160. When would you choose Power Query or Office Scripts instead of VBA?</b></summary>

Use Power Query for repeated data import and clean-up. Use Office Scripts for Excel on the web and cloud automation. VBA remains the tool for desktop automation and older files.

</details>

### 27 Python in Excel and Office Scripts

_Lesson 27_

<details>
<summary><b>Q161. What is Python in Excel?</b></summary>

A feature that lets you run Python (with pandas, matplotlib and more) in cells. It runs in a Microsoft cloud container, reads sheet data with xl() and returns results to the sheet.

</details>

<details>
<summary><b>Q162. Can Python in Excel read files from your computer?</b></summary>

No. The Python environment is isolated from local files and the internet. Bring data into the workbook with Power Query first.

</details>

<details>
<summary><b>Q163. What are Office Scripts and how are they different from VBA?</b></summary>

Office Scripts are TypeScript automation scripts for Excel (web, Windows and Mac) that you can record, edit and run from Power Automate. VBA is older, desktop-only and macro-based.

</details>

<details>
<summary><b>Q164. When would you use Python in Excel instead of formulas?</b></summary>

For statistics, large data transformations, machine learning and charts that need pandas or matplotlib.

</details>

<details>
<summary><b>Q165. How do you check that a Python result is right?</b></summary>

Compare it with an Excel formula or a pivot table for a few groups, and test with a small known data set.

</details>

### 29 Performance and Auditing

_Lesson 29_

<details>
<summary><b>Q166. What makes a workbook slow and how do you fix it?</b></summary>

Heavy whole-column array formulas, many volatile functions, large lookups on unsorted data and too much formatting. Fix with bounded ranges or Tables, INDEX instead of OFFSET, helper columns, Power Query and by saving as .xlsb.

</details>

<details>
<summary><b>Q167. What is a volatile function?</b></summary>

One that recalculates every time anything changes, such as TODAY, NOW, RAND, OFFSET and INDIRECT.

</details>

<details>
<summary><b>Q168. How do you audit someone else's workbook?</b></summary>

Read the structure, trace precedents and dependents, use Go To Special to see constants and formulas, check for hard-coded numbers and errors, reconcile totals and test with edge cases.

</details>

<details>
<summary><b>Q169. What is a reconciliation check?</b></summary>

Two independent calculations of the same total with a check cell showing the difference, which should be zero.

</details>

<details>
<summary><b>Q170. What do Ctrl + [ and Ctrl + ] do?</b></summary>

Ctrl + [ selects the cells that feed the active formula (precedents). Ctrl + ] selects the cells that use the active cell (dependents).

</details>

### 33 Capstone Projects

_Lesson 33_

<details>
<summary><b>Q171. How do you choose a good portfolio project?</b></summary>

Choose a real-looking business question, enough data to be interesting, a clear decision at the end and skills that match the job you want.

</details>

<details>
<summary><b>Q172. What should a project README contain?</b></summary>

The problem, data and its source, the method and key formulas, results with a chart or screenshot, insights and recommendation, and what you would do next.

</details>

<details>
<summary><b>Q173. How do you show that your numbers are right?</b></summary>

Reconcile totals with a second method, keep a checks area, show sample hand calculations and document assumptions.

</details>

<details>
<summary><b>Q174. How do you present a result to a non-technical manager?</b></summary>

One page: the answer first, two or three KPIs, one clear chart, a recommendation and the risks. No jargon.

</details>

<details>
<summary><b>Q175. What would you improve in a first version of your project?</b></summary>

Add data validation, protect formulas, add a slicer, document assumptions and test with new data.

</details>

## 🧩 Scenario and practical questions

### 🟢 Beginner

<details>
<summary><b>Q176. How do you keep the top row visible while you scroll?</b></summary>

View, Freeze Panes, Freeze Top Row. To freeze rows and columns, select the cell below and to the right of what you want frozen, then choose Freeze Panes.

</details>

<details>
<summary><b>Q177. How do you create serial numbers 1, 2, 3 automatically?</b></summary>

Type 1 and 2, select both and drag the fill handle. Or use =ROW()-1 in the first data row, or =SEQUENCE(100) in Excel 2021 and later.

</details>

<details>
<summary><b>Q178. How do you find duplicates quickly?</b></summary>

Home, Conditional Formatting, Highlight Cells Rules, Duplicate Values. Or add a helper column with =COUNTIF($A$2:$A$100,A2)>1.

</details>

<details>
<summary><b>Q179. What is the difference between Clear Contents and Delete?</b></summary>

Clear Contents (the Delete key) empties the cells and keeps them in place. Delete (right-click, Delete) removes the cells and shifts neighbouring cells up or left.

</details>

<details>
<summary><b>Q180. How do you print a wide sheet on one page across?</b></summary>

Page Layout, Scale to Fit: set Width to 1 page. Use Print Titles to repeat the header row on every page, and set a Print Area.

</details>

<details>
<summary><b>Q181. How do you turn formulas into fixed values?</b></summary>

Copy the cells, then Paste Special, Values (Ctrl + Alt + V, then V).

</details>

<details>
<summary><b>Q182. How do you change text to upper, lower or proper case?</b></summary>

Use UPPER, LOWER or PROPER in a helper column, then paste as values. Flash Fill (Ctrl + E) can also do it from an example.

</details>

<details>
<summary><b>Q183. How do you total a column quickly?</b></summary>

Select the cell under the numbers and press Alt + = (AutoSum). Or select the numbers and read the total on the status bar.

</details>

### 🟡 Intermediate

<details>
<summary><b>Q184. VLOOKUP returns #N/A but the value is in the table. What do you check?</b></summary>

Spaces (use TRIM), numbers stored as text versus numbers, a missing FALSE for exact match, a locked table range, and whether the key is in the first column of the table.

</details>

<details>
<summary><b>Q185. How do you sum revenue for March only?</b></summary>

With a date column: =SUMIFS(rev, date, ">="&DATE(2026,3,1), date, "<"&DATE(2026,4,1)). This also handles the year. A pivot table grouped by month and year works as well.

</details>

<details>
<summary><b>Q186. How do you count unique values?</b></summary>

In Excel 2021 and later: =COUNTA(UNIQUE(range)). In older versions: =SUMPRODUCT(1/COUNTIF(range,range)) (no blank cells). In a pivot table, add the data to the Data Model and use Distinct Count.

</details>

<details>
<summary><b>Q187. How do you highlight rows whose due date has passed and are not done?</b></summary>

Conditional formatting with a formula rule applied to the whole table: =AND($C2<TODAY(),$D2<>"Done"). The $ locks the columns and lets the row move.

</details>

<details>
<summary><b>Q188. How do you do a two-way lookup?</b></summary>

=INDEX(data, MATCH(row_key, row_headers, 0), MATCH(column_key, column_headers, 0)). XLOOKUP can be nested for the same job in newer Excel.

</details>

<details>
<summary><b>Q189. How do you combine twelve monthly files with the same columns?</b></summary>

Power Query: Data, Get Data, From File, From Folder, then Combine and Transform. New files are picked up on Refresh.

</details>

<details>
<summary><b>Q190. How do you build a dependent dropdown list?</b></summary>

In Excel 2021 and later, use UNIQUE and FILTER in a helper range and point validation lists at the spill range (for example E2#). In older versions use named ranges with INDIRECT.

</details>

<details>
<summary><b>Q191. How do you find the second largest value that meets a condition?</b></summary>

=AGGREGATE(14,6,values/(criteria_range="X"),2). The division turns non-matching rows into errors that option 6 ignores. In Excel 365 you can also use SORT with FILTER.

</details>

<details>
<summary><b>Q192. How do you calculate someone's age or years of service?</b></summary>

=DATEDIF(start_date, end_date, "Y"). Use a fixed end date for repeatable reports or TODAY() for the age today.

</details>

<details>
<summary><b>Q193. How do you compare two lists to find items missing from one?</b></summary>

In a helper column use =COUNTIF(list_b, A2)=0, or XLOOKUP with a not-found message. For big lists use Power Query Merge with a Left Anti join.

</details>

<details>
<summary><b>Q194. A pivot table shows Count of Revenue instead of Sum of Revenue. Why?</b></summary>

Some cells in the column are blank or text, so Excel defaults to Count. Clean the data to numbers and change the Value Field Settings to Sum.

</details>

<details>
<summary><b>Q195. How do you protect formulas but let people enter inputs?</b></summary>

Unlock the input cells (Format Cells, Protection), then protect the sheet (Review, Protect Sheet). Remember that sheet protection is not strong security.

</details>

### 🔴 Advanced

<details>
<summary><b>Q196. Explain LET and when you would use it.</b></summary>

LET names values or calculations inside a formula. Use it when the same expression repeats, to make long formulas readable and to calculate each repeated part once.

</details>

<details>
<summary><b>Q197. How would you build an interactive dashboard?</b></summary>

Keep data in Tables or the Data Model, calculate with pivots or SUMIFS on a calculations sheet, build charts and KPI cells on a dashboard sheet, add slicers and a timeline, hide gridlines, protect the sheet and set the print area.

</details>

<details>
<summary><b>Q198. Explain the Data Model and relationships.</b></summary>

The Data Model is an in-workbook database that holds several tables. Relationships link them on key columns (one to many) so a pivot table can summarise across tables. DAX measures calculate on demand using the current filters.

</details>

<details>
<summary><b>Q199. A workbook takes two minutes to calculate. How do you find the cause?</b></summary>

Check calculation mode, volatile functions, whole-column array formulas, many lookups on unsorted data, large conditional formatting ranges and file size. Test by switching sections off. Move heavy transformations to Power Query.

</details>

<details>
<summary><b>Q200. How do you audit a workbook you did not build?</b></summary>

Read the structure, use Trace Precedents and Dependents, Go To Special for constants and formulas, search for hard-coded numbers in formulas, test totals with an independent method and try edge cases.

</details>

<details>
<summary><b>Q201. Compare XLOOKUP and INDEX-MATCH.</b></summary>

Both look in any direction. XLOOKUP is simpler, has a built-in not-found argument and a search direction, but needs Excel 2021 or later. INDEX-MATCH works in every version and gives more control for two-way lookups.

</details>

<details>
<summary><b>Q202. How do you analyse five million rows?</b></summary>

Not in a normal worksheet (the limit is 1,048,576 rows). Load the data with Power Query into the Data Model and use Power Pivot, or summarise in a database with SQL, or use Python. Then bring the summary into Excel.

</details>

<details>
<summary><b>Q203. How can you automate a task that must run every day without opening Excel?</b></summary>

Use Office Scripts with Power Automate for files in the cloud, a scheduled Python script, or a database job. VBA macros need Excel to be open on a desktop computer.

</details>

<details>
<summary><b>Q204. How do you calculate a rolling 12-month total?</b></summary>

=SUMIFS(rev, date, ">"&EDATE(end_date,-12), date, "<="&end_date). Put the end date in a cell and copy the formula across a list of month ends.

</details>

<details>
<summary><b>Q205. What are GROUPBY and PIVOTBY?</b></summary>

Microsoft 365 functions that build a grouped summary (GROUPBY) or a pivot-style cross table (PIVOTBY) with one formula that updates with the data. They are not in Excel 2021 or Excel 2024.

</details>

<details>
<summary><b>Q206. How would you check an AI-generated formula before using it?</b></summary>

Ask for the formula and an explanation, test a few rows by hand, reconcile totals with SUM or a pivot, check the range covers all rows, confirm the function exists in your Excel version and test edge cases such as blanks and zeros.

</details>

---
[⬅️ Back to the course index](README.md)
