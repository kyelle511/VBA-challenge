# VBA-challenge
Module 2 Challenge

Files included with descriptions:

1- VBAChallenge.vbs: includes two Macro Scripts 'SummaryStats' and 'NewDataActions'. 'SummaryStats' should be run first and then 'NewDataActions'.  'SummaryStats' pulls the ticker symbol, calculates the yearly change for that ticker, calculates the percentage change for that ticker and calculates the total stock volume.  'NewDataActions' will conditionally format the yearly change column that was calculated and determine which stock had the greatest % increase, greatest % decrease, and the greatest total volume. 

2-ScreenshotWorksheet1.png: shows the first worksheet (2018) in the excel file after running the scripts

3-ScreenshotWorksheet2.png:shows the second worksheet (2019) in the excel file after running the scripts

4-ScreenshotWorksheet3.png:shows the third worksheet (2020) in the excel file after running the scripts

Citations/Code Sources:
All are located in the VBAChallenge.vbs file

Figuring out how to format percentage for both macros in the Percentage change column and then again for greatest % increase and greatest % decrease. I used this in both 'SummaryStats' and 'NewDataActions' Macros (https://excelchamps.com/vba/functions/formatpercent/#:~:text=The%20VBA%20FORMATPERCENT%20function%20is,as%20a%20string%20data%20type).

Auto fit columns so that they accommodated new values as part of the 'SummaryStats' Macro (https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit)

Max function (which I then used to determine min) whish I used in the 'NewDataActions' to determine greatest amounts (https://www.wallstreetmojo.com/vba-max/)

To fix the Total Volume formatting so it doesn’t show up in exponential number format.  Which I used in the 'NewDataActions' for the Greatest Volume Cell (https://stackoverflow.com/questions/57299072/show-number-without-exponential-expression-after-remove-non-numeric-character)
