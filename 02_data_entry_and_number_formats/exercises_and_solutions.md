# 🏋️ Exercises and Solutions · 02 Data Entry and Number Formats

Try each task yourself first. Open [practice_workbook.xlsx](practice_workbook.xlsx), use the data sheets, and type your answer in the yellow cells of the `Exercises` sheet. Then open the solution here (or the `Solutions` sheet).

| # | Level | Task |
|---|---|---|
| 1 | 🟢 Easy | Show 0.256 as a percentage with no decimals using TEXT. |
| 2 | 🟢 Easy | What is actually stored in a date cell? |
| 3 | 🟢 Easy | Make the number 7 display as 00007 using a formula. |
| 4 | 🟡 Medium | Show the first order date as month and year, like Jan-2025. |
| 5 | 🟡 Medium | Contacts!D2 holds the text 2024-03-05. Convert it into a real date. |
| 6 | 🟡 Medium | What date is 30 days after the first order? |
| 7 | 🔴 Hard | Add up the text amounts in Contacts column E (E2:E13) as real numbers. |
| 8 | 🔴 Hard | A cell holds 2.456 but is formatted to show one decimal (2.5). What does a formula that multiplies it by 10 return? |

## Exercise 1 · 🟢 Easy

**Task:** Show 0.256 as a percentage with no decimals using TEXT.

<details>
<summary>Show solution</summary>

```excel
=TEXT(0.256,"0%")
```
**Result:** `26%`

**Explanation:** The format "0%" multiplies by 100 and rounds for display.

</details>

## Exercise 2 · 🟢 Easy

**Task:** What is actually stored in a date cell?

<details>
<summary>Show solution</summary>

**Explanation:** A serial number (a whole number that counts days). The date you see is only a number format.

</details>

## Exercise 3 · 🟢 Easy

**Task:** Make the number 7 display as 00007 using a formula.

<details>
<summary>Show solution</summary>

```excel
=TEXT(7,"00000")
```
**Result:** `00007`

**Explanation:** Use TEXT with five zeros.

</details>

## Exercise 4 · 🟡 Medium

**Task:** Show the first order date as month and year, like Jan-2025.

<details>
<summary>Show solution</summary>

```excel
=TEXT(Sales!B2,"mmm-yyyy")
```
**Result:** `Jan-2025`

**Explanation:** Use the pattern "mmm-yyyy".

</details>

## Exercise 5 · 🟡 Medium

**Task:** Contacts!D2 holds the text 2024-03-05. Convert it into a real date.

<details>
<summary>Show solution</summary>

```excel
=DATEVALUE(Contacts!D2)
```
**Result:** `2024-03-05`

**Explanation:** DATEVALUE turns date-looking text into a serial number. Format the cell as a date to read it.

</details>

## Exercise 6 · 🟡 Medium

**Task:** What date is 30 days after the first order?

<details>
<summary>Show solution</summary>

```excel
=Sales!B2+30
```
**Result:** `2025-02-03`

**Explanation:** Dates are numbers, so just add 30.

</details>

## Exercise 7 · 🔴 Hard

**Task:** Add up the text amounts in Contacts column E (E2:E13) as real numbers.

<details>
<summary>Show solution</summary>

```excel
=SUMPRODUCT(VALUE(SUBSTITUTE(SUBSTITUTE(Contacts!E2:E13,"$",""),",","")))
```
**Result:** `15,647.8`

**Explanation:** SUBSTITUTE removes the symbols, VALUE converts, SUMPRODUCT adds the array without needing Ctrl+Shift+Enter.

</details>

## Exercise 8 · 🔴 Hard

**Task:** A cell holds 2.456 but is formatted to show one decimal (2.5). What does a formula that multiplies it by 10 return?

<details>
<summary>Show solution</summary>

```excel
=2.456*10
```
**Result:** `24.56`

**Explanation:** The format does not change the value, so Excel multiplies 2.456 and returns 24.56, not 25.

</details>

---
[⬅️ Back to the notes](README.md)
