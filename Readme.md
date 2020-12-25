## Multiple Currency Price Quoter

Summary
An excel utility to generate Price quotes for products in multiple currencies. Useful for businesses having customers all over the world with a single set of products. The utility solves the need to update currencies manually. It requires the prices of products in USD, which are converted on the basis of a current hourly rate. 

Further proposed plans:-
Usage of base currency other than USD.
Utility to allow price discrimination, charging different prices to different countries on basis of tiers of countries. e.g. Allow to charge 50% price for Kenya than USA.

Steps to use:-

1. Download the repository from the top-right download option github.com https://github.com/Wanderingbubble/PriceQuoter/archive/main.zip  
2. Unzip and Open excel  
3. enable macros  
4. Go to sheet "MainOperation"  
5. Enter the names of products and prices of products  
6. Use macro7 with shortcut Ctrl-m  
 
7. And Voila! your quote is ready on the main sheet with 6 currencies!  

Limitation:-
For now you can only list 15 products, without editing/modifying the excel sheet.
contact me at https://www.linkedin.com/in/rajeshjoshi607/ for changes or other modifications


How does it work?

1) it imports  currency exchange rate data from  http://www.floatrates.com/daily/usd.xml & pastes it on the sheet "RawData1"
2) This imported data is unsorted and not indexed, hence we must sort it and index it.
2) Index number listed  are on the sheet "Indexnumbers"
3) Press ctrl-m as a macro shortcut to:-
 Copy the index numbers from "Indexnumbers" to "Sorting" and paste it on A column.
Copy the  imported index numbers from "RawData1" to "Sorting" and paste it on B column.
Sort this imported data on "B column" in ascending order.
4) The index number assigned to the relevant currencies are:-
	AUD -   9
	GBP -  140
	INR -   59
	CAD - 25
	CNY - 39
5) We use VLOOPUP(Indexnumber's cell,Range of sorted data,column no 4,FALSE) to extract the exchange rates of such 6 currencies with respect to USD.
6) In the  sheet "MainOperation", we enter the prices of our products in USD, and multiply it with the exchange rates.
7) In the sheet "RoundOff" we  use the excel function ROUNDOFF(convertedPrice,2) to round off the prices to quote to two decimal points.
8) Finally we import the data in the Product Price List which is correctly formated with the relevant currencies.
9) Such quoted price list can be sent to customers all over the world!
