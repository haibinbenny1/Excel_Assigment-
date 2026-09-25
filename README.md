# Excel_Assigment- Data Exploration

1. Sum, Count, Average:

Total Price: Calculates the sum of all product prices using =SUM(D2:D35).

Total Products: Counts the total number of products using =COUNTA(B2:B35).

Average Price: Calculates the mean price of all products using =AVERAGE(D2:D35).

2. Min and Max:

Minimum Price: Finds the lowest product price using =MIN(D2:D35).

Maximum Price: Finds the highest product price using =MAX(D2:D35).

3. IF Function (Price Range):

Categorizes products as 'High Price' if they are $500 or more, and 'Standard Price' otherwise using =IF(D2>=500, "High Price", "Standard Price").

4. SUMIF and COUNTIF:

Electronics Total: Sums up prices specifically for the 'Electronics' category using =SUMIF(F2:F35, "Electronics", D2:D35).

Count Under $100: Counts how many products cost less than $100 using =COUNTIF(D2:D35, "<100").

5. Text Formatting (LEFT, RIGHT, MID):

Day: Extracts the first 2 characters from the Product ID using =LEFT(A2, 2).

Country Code: Extracts the last 2 characters from the Product ID using =RIGHT(A2, 2).

Month: Extracts characters from the middle of the Product ID to get the month using =MID(A2, 4, 3).

