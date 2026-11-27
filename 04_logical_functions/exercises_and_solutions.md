# 🏋️ Exercises and Solutions · 04 Logical Functions: IF, IFS, AND, OR, NOT and SWITCH

Try each task yourself first. Open [practice_workbook.xlsx](practice_workbook.xlsx), use the data sheets, and type your answer in the yellow cells of the `Exercises` sheet. Then open the solution here (or the `Solutions` sheet).

| # | Level | Task |
|---|---|---|
| 1 | 🟢 Easy | Show "High" if the first employee's salary is above 60,000, otherwise "Standard". |
| 2 | 🟢 Easy | What does AND return if one of its tests is FALSE? |
| 3 | 🟢 Easy | Mark the first sales order as "Bulk" if it has 10 or more units, otherwise "Normal". |
| 4 | 🟡 Medium | Classify the first order's unit price: 500 or more is Premium, 100 or more is Standard, otherwise Budget. |
| 5 | 🟡 Medium | Count the employees with a rating of exactly 5 without a helper column. |
| 6 | 🟡 Medium | Show the regional manager for the first order: North is Anita, South is Bala, East is Chetan, West is Dev, anything else Unknown. |
| 7 | 🔴 Hard | For the first employee: rating 4 or more and joined before 1 Jan 2020 gives "Loyalty bonus"; rating 4 or more gives "Performance bonus"; otherwise "None". |
| 8 | 🔴 Hard | Count the orders from West or East that have 10 or more units. |

## Exercise 1 · 🟢 Easy

**Task:** Show "High" if the first employee's salary is above 60,000, otherwise "Standard".

<details>
<summary>Show solution</summary>

```excel
=IF(Employees!F2>60000,"High","Standard")
```
**Result:** `High`

**Explanation:** A single IF with a greater-than test.

</details>

## Exercise 2 · 🟢 Easy

**Task:** What does AND return if one of its tests is FALSE?

<details>
<summary>Show solution</summary>

**Explanation:** FALSE. AND needs every test to be TRUE.

</details>

## Exercise 3 · 🟢 Easy

**Task:** Mark the first sales order as "Bulk" if it has 10 or more units, otherwise "Normal".

<details>
<summary>Show solution</summary>

```excel
=IF(Sales!G2>=10,"Bulk","Normal")
```
**Result:** `Normal`

**Explanation:** IF with >=.

</details>

## Exercise 4 · 🟡 Medium

**Task:** Classify the first order's unit price: 500 or more is Premium, 100 or more is Standard, otherwise Budget.

<details>
<summary>Show solution</summary>

```excel
=IFS(Sales!H2>=500,"Premium",Sales!H2>=100,"Standard",TRUE,"Budget")
```
**Result:** `Budget`

**Explanation:** IFS checks the strictest test first.

</details>

## Exercise 5 · 🟡 Medium

**Task:** Count the employees with a rating of exactly 5 without a helper column.

<details>
<summary>Show solution</summary>

```excel
=SUMPRODUCT(--(Employees!$G$2:$G$25=5))
```
**Result:** `4`

**Explanation:** Compare the whole range, turn TRUE/FALSE into 1/0 and add.

</details>

## Exercise 6 · 🟡 Medium

**Task:** Show the regional manager for the first order: North is Anita, South is Bala, East is Chetan, West is Dev, anything else Unknown.

<details>
<summary>Show solution</summary>

```excel
=SWITCH(Sales!C2,"North","Anita","South","Bala","East","Chetan","West","Dev","Unknown")
```
**Result:** `Dev`

**Explanation:** SWITCH with a default.

</details>

## Exercise 7 · 🔴 Hard

**Task:** For the first employee: rating 4 or more and joined before 1 Jan 2020 gives "Loyalty bonus"; rating 4 or more gives "Performance bonus"; otherwise "None".

<details>
<summary>Show solution</summary>

```excel
=IF(AND(Employees!G2>=4,Employees!E2<DATE(2020,1,1)),"Loyalty bonus",IF(Employees!G2>=4,"Performance bonus","None"))
```
**Result:** `Performance bonus`

**Explanation:** Test the strictest case first and use AND for two conditions.

</details>

## Exercise 8 · 🔴 Hard

**Task:** Count the orders from West or East that have 10 or more units.

<details>
<summary>Show solution</summary>

```excel
=SUMPRODUCT(((Sales!$C$2:$C$41="West")+(Sales!$C$2:$C$41="East")>0)*(Sales!$G$2:$G$41>=10))
```
**Result:** `12`

**Explanation:** Add the two region tests (OR as plus), compare to 0, multiply by the units test, then add with SUMPRODUCT.

</details>

---
[⬅️ Back to the notes](README.md)
