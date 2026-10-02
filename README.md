# Excel_Assignment - Data Exploration

## 1. Sum, Count, Average

**Total Price:** Calculates the sum of all product prices using `=SUM(D2:D35)`.

**Total Products:** Counts the total number of products using `=COUNTA(B2:B35)`.

**Average Price:** Calculates the average price of all products using `=AVERAGE(D2:D35)`.

## 2. Min and Max

**Minimum Price:** Finds the lowest product price using `=MIN(D2:D35)`.

**Maximum Price:** Finds the highest product price using `=MAX(D2:D35)`.

## 3. IF Function (Price Range)

Categorizes products as **"High Price"** if the price is $500 or more, and **"Standard Price"** otherwise.

Formula:

`=IF(D2>=500,"High Price","Standard Price")`

## 4. SUMIF and COUNTIF

**Electronics Total:** Calculates the total price of products in the **Electronics** category using:

`=SUMIF(F2:F35,"Electronics",D2:D35)`

**Count Under $100:** Counts the number of products priced below $100 using:

`=COUNTIF(D2:D35,"<100")`

## 5. Text Formatting (LEFT, RIGHT, MID)

**Day:** Extracts the first 2 characters from the Product ID using:

`=LEFT(A2,2)`

**Country Code:** Extracts the last 2 characters from the Product ID using:

`=RIGHT(A2,2)`

**Month:** Extracts the month from the middle of the Product ID using:

`=MID(A2,4,3)`
