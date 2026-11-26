# 🏋️ Exercises and Solutions · 03 Formulas and Cell References

Try each task yourself first. Open [practice_workbook.xlsx](practice_workbook.xlsx), use the data sheets, and type your answer in the yellow cells of the `Exercises` sheet. Then open the solution here (or the `Solutions` sheet).

| # | Level | Task |
|---|---|---|
| 1 | 🟢 Easy | What does the dollar sign do in $B$3? |
| 2 | 🟢 Easy | What does the F4 key do while the cursor is inside a reference? |
| 3 | 🟢 Easy | Using the named range, total all revenue. |
| 4 | 🟡 Medium | What commission does the third order in the Refs sheet (row 20) earn? |
| 5 | 🟡 Medium | Total revenue of the West region using names. |
| 6 | 🟡 Medium | Read the last cell of the multiplication table (6 times 6) in Refs!G14. |
| 7 | 🔴 Hard | You copy =A1*B1 from cell C1 to cell D2. What formula does D2 hold? |
| 8 | 🔴 Hard | Work out the tax amount for all orders at the TaxRate in the Refs sheet. |

## Exercise 1 · 🟢 Easy

**Task:** What does the dollar sign do in $B$3?

<details>
<summary>Show solution</summary>

**Explanation:** It locks the column and the row, so the reference does not change when the formula is copied.

</details>

## Exercise 2 · 🟢 Easy

**Task:** What does the F4 key do while the cursor is inside a reference?

<details>
<summary>Show solution</summary>

**Explanation:** It cycles through absolute, row-locked, column-locked and relative versions of the reference.

</details>

## Exercise 3 · 🟢 Easy

**Task:** Using the named range, total all revenue.

<details>
<summary>Show solution</summary>

```excel
=SUM(Revenue)
```
**Result:** `47,476.75`

**Explanation:** The name Revenue points to Sales!J2:J41.

</details>

## Exercise 4 · 🟡 Medium

**Task:** What commission does the third order in the Refs sheet (row 20) earn?

<details>
<summary>Show solution</summary>

```excel
=Refs!B20*Refs!$B$3
```
**Result:** `115.2`

**Explanation:** Row 20 holds the revenue of the third order. Multiply by the locked rate.

</details>

## Exercise 5 · 🟡 Medium

**Task:** Total revenue of the West region using names.

<details>
<summary>Show solution</summary>

```excel
=SUMIF(Region,"West",Revenue)
```
**Result:** `15,552.5`

**Explanation:** SUMIF with the names Region and Revenue.

</details>

## Exercise 6 · 🟡 Medium

**Task:** Read the last cell of the multiplication table (6 times 6) in Refs!G14.

<details>
<summary>Show solution</summary>

```excel
=Refs!G14
```
**Result:** `36`

**Explanation:** The table formula is =$A9*B$8 copied across and down.

</details>

## Exercise 7 · 🔴 Hard

**Task:** You copy =A1*B1 from cell C1 to cell D2. What formula does D2 hold?

<details>
<summary>Show solution</summary>

**Explanation:** Both references move one column right and one row down, so D2 holds =B2*C2.

</details>

## Exercise 8 · 🔴 Hard

**Task:** Work out the tax amount for all orders at the TaxRate in the Refs sheet.

<details>
<summary>Show solution</summary>

```excel
=SUM(Revenue)*TaxRate
```
**Result:** `8,545.815`

**Explanation:** Multiply the total revenue by the named rate.

</details>

---
[⬅️ Back to the notes](README.md)
