# Multiple Currency Price Quoter

## Summary
An excel utility to generate Price quotes for products in multiple currencies. Useful for businesses having customers all over the world with a single set of products. The utility solves the need to update currencies manually. It requires the prices of products in USD, which are converted on the basis of a current hourly rate. 

## Further proposed plans:-
Usage of base currency other than USD.
Utility to allow price discrimination, charging different prices to different countries on basis of tiers of countries. e.g. Allow to charge 50% price for Kenya than USA.

## Steps to use:-

1. Download the repository from the top-right download option github.com https://github.com/Wanderingbubble/PriceQuoter/archive/main.zip  
2. Unzip and Open excel  
3. enable macros  
4. Go to sheet "MainOperation"  
5. Enter the names of products and prices of products  
6. Use macro7 with shortcut Ctrl-m  
 
7. And Voila! your quote is ready on the main sheet with 6 currencies!  

## Limitation:-
For now you can only list 15 products, without editing/modifying the excel sheet.
contact me at https://www.linkedin.com/in/rajeshjoshi607/ for changes or other modifications


## How does it work?

1. it imports  currency exchange rate data from  http://www.floatrates.com/daily/usd.xml & pastes it on the sheet "RawData1"
2. This imported data is unsorted and not indexed, hence we must sort it and index it.
3. Index number listed  are on the sheet "Indexnumbers"
4. Press ctrl-m as a macro shortcut to:-
 Copy the index numbers from "Indexnumbers" to "Sorting" and paste it on A column.
Copy the  imported index numbers from "RawData1" to "Sorting" and paste it on B column.
Sort this imported data on "B column" in ascending order.
5. The index number assigned to the relevant currencies are:-
	AUD -   9
	GBP -  140
	INR -   59
	CAD - 25
	CNY - 39
6. We use VLOOPUP(Indexnumber's cell,Range of sorted data,column no 4,FALSE) to extract the exchange rates of such 6 currencies with respect to USD.
7. In the  sheet "MainOperation", we enter the prices of our products in USD, and multiply it with the exchange rates.
8. In the sheet "RoundOff" we  use the excel function ROUNDOFF(convertedPrice,2) to round off the prices to quote to two decimal points.
9. Finally we import the data in the Product Price List which is correctly formated with the relevant currencies.
10. Such quoted price list can be sent to customers all over the world!

## Debugging:-
For data import errors
1. open the data tab in excel, go to sheet "Rawdata1".
2. Click on the A1 cell in "RawData1", and use cnrl-a to select the whole table and delete it.
3. select the cell A1, Select import data from other sources, from xml
4. Make sure you have a working internet
3. Instead of the file name enter the link "http://www.floatrates.com/daily/usd.xml" in it and click open.
4. Import data dialog will appear, check 2nd option "existing worksheet" and enter cell "=$A1$1" (make sure you are in sheet "RawData1")
4. Run Macro with cnrl-m to sort the new data and the excel will update itself

## Changelog
Dec 2020
-Updated the readme.mc
-Found issues in macro creating duplicate connections which slows the application.

https://docs.github.com/en/free-pro-team@latest/github/writing-on-github/basic-writing-and-formatting-syntax
formatting tools for reference