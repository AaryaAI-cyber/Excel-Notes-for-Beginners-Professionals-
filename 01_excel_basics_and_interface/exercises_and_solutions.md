# 🏋️ Exercises and Solutions · 01 Excel Basics and Interface

Try each task yourself first. Open [practice_workbook.xlsx](practice_workbook.xlsx), use the data sheets, and type your answer in the yellow cells of the `Exercises` sheet. Then open the solution here (or the `Solutions` sheet).

| # | Level | Task |
|---|---|---|
| 1 | 🟢 Easy | In a blank cell, find the total number of units sold. |
| 2 | 🟢 Easy | How many orders are listed on the Sales sheet? |
| 3 | 🟢 Easy | What keyboard shortcut inserts an AutoSum on Windows? |
| 4 | 🟡 Medium | Find the gap between the biggest and smallest order revenue. |
| 5 | 🟡 Medium | Show the average revenue per order, rounded to 2 decimals. |
| 6 | 🟡 Medium | Show total revenue in thousands (for example 47.48 instead of 47,476.75). |
| 7 | 🔴 Hard | What does =10+2*3^2 return? Work it out by hand, then confirm. |
| 8 | 🔴 Hard | Total revenue of only the first 10 orders (rows 2 to 11). |

## Exercise 1 · 🟢 Easy

**Task:** In a blank cell, find the total number of units sold.

<details>
<summary>Show solution</summary>

```excel
=SUM(Sales!$G$2:$G$41)
```
**Result:** `339`

**Explanation:** Add the Units column (G) with SUM.

</details>

## Exercise 2 · 🟢 Easy

**Task:** How many orders are listed on the Sales sheet?

<details>
<summary>Show solution</summary>

```excel
=COUNTA(Sales!$A$2:$A$41)
```
**Result:** `40`

**Explanation:** Use COUNTA on the OrderID column because the IDs are text.

</details>

## Exercise 3 · 🟢 Easy

**Task:** What keyboard shortcut inserts an AutoSum on Windows?

<details>
<summary>Show solution</summary>

**Explanation:** Alt + = (hold Alt, press the equals key). On a Mac use Cmd + Shift + T.

</details>

## Exercise 4 · 🟡 Medium

**Task:** Find the gap between the biggest and smallest order revenue.

<details>
<summary>Show solution</summary>

```excel
=MAX(Sales!$J$2:$J$41)-MIN(Sales!$J$2:$J$41)
```
**Result:** `2709`

**Explanation:** Subtract MIN from MAX.

</details>

## Exercise 5 · 🟡 Medium

**Task:** Show the average revenue per order, rounded to 2 decimals.

<details>
<summary>Show solution</summary>

```excel
=ROUND(AVERAGE(Sales!$J$2:$J$41),2)
```
**Result:** `1,186.92`

**Explanation:** Wrap AVERAGE inside ROUND.

</details>

## Exercise 6 · 🟡 Medium

**Task:** Show total revenue in thousands (for example 47.48 instead of 47,476.75).

<details>
<summary>Show solution</summary>

```excel
=SUM(Sales!$J$2:$J$41)/1000
```
**Result:** `47.4768`

**Explanation:** Divide the total by 1000.

</details>

## Exercise 7 · 🔴 Hard

**Task:** What does =10+2*3^2 return? Work it out by hand, then confirm.

<details>
<summary>Show solution</summary>

```excel
=10+2*3^2
```
**Result:** `28`

**Explanation:** Powers first (3^2 = 9), then multiply (2*9 = 18), then add (10+18 = 28).

</details>

## Exercise 8 · 🔴 Hard

**Task:** Total revenue of only the first 10 orders (rows 2 to 11).

<details>
<summary>Show solution</summary>

```excel
=SUM(Sales!J2:J11)
```
**Result:** `11,838`

**Explanation:** Give SUM a smaller range. The range is written first cell, colon, last cell.

</details>

---
[⬅️ Back to the notes](README.md)
