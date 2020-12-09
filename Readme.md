Open excel
enable macros
Go to sheet "MainOperation"
Enter the names of products and prices of products
Use macro7 with shortcut Ctrl-m
Voila
Your qoute is ready on them main sheet with 6 currencies!

Limitation:-
For now you can only list 15 products, withtout the editing/modifying the excel sheet.
contact me at https://www.linkedin.com/in/rajeshjoshi607/ for changes or other modifications


How does it work?

1) it imports  currency exhange rate data from  http://www.floatrates.com/daily/usd.xml & pastes it on the sheet "RawData1"
2) This imported data is unsorted and not indexed, hence we must sort it and index it.
2) Index number listed  are on the sheet "Indexnumbers"
3) Press ctrl-m as a macro shortcut to:-
 Copy the index numbers from "Indexnumbers" to "Sorting" and paste it on A coloum.
Copy the  imported index numbers from "RawData1" to "Sorting" and paste it on B column.
Sort this imported data on "B column" in assending order.
4) The index number assigned to the relevant currencies are:-
	EUR - 45
	AUD -   9
	GBP -  140
	INR -   59
	CAD - 25
	CNY - 39
5) We use VLOOPUP(Indexnumber's cell,Range of sorted data,column no 4,FALSE) to extract the exhange rates of such 6 currencies with respect to USD.
6) In the  sheet "MainOperation", we enter the prices of our products in USD, and multiply it witht the exhange rates.
7) In the sheet "RoundOff" we  use excel fucntion ROUNDOFF(convertedPrice,2) to round off the prices to qoute to two decimal points.
8) Finally we import the data in the Product Price List which is correctly formated with the relevant currencies.
9) Such Qouted price list can be sent to customers all over the world !
