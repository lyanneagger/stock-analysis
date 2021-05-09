# Stock Analysis with Visual Basic

## Overview of Project

### Purpose

This analysis will simplify and refactor the original VBA script to streamline the process and reduce execution time. By refactoring the original code, it will work for a larger number of stocks more quickly and present the information in a way that is quickly formatted and easy to understand.

## Results

The original script looped through the data repeatedly to check for each of the values requested, using nested 'for' loops. Creating a 'tickerIndex' variable allows the code to gather and store the information without the nested 'for loop. For example, the original code opens a 'for' loop for each index of the ticker array and then has a nested 'for' loop to check each row's data for the Volume:

```
For i = 0 to 11
	For j = rowStart to rowEnd
		If Cells(j, 1).Value = ticker Then
			totalVolume = totalVolume + Cells(j, 8).Value
		End If
	Next j
Next i
```

However, in the refactored code, collecting the same information is simplified to:
```
For i = 2 to RowCount
	tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
Next i
```

Reducing nested 'for' loops is a major factor in reducing the run time. Below is the comparison of the execution times for the original and the refactored code for both 2017 and 2018. 

Execution time for 2017 analysis - original vs. refactored code:

![Alt Text](https://github.com/lyanneagger/stock-analysis/blob/main/Resources/VBA_Challenge_2017_v1.png)
![Alt Text](https://github.com/lyanneagger/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

Execution time for 2018 analysis - original vs. refactored code:

![Alt Text](https://github.com/lyanneagger/stock-analysis/blob/main/Resources/VBA_Challenge_2018_v1.png)
![Alt Text](https://github.com/lyanneagger/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

It should also be noted that the original code also does not inclue formatting, as that was a separate subroutine.


###Summary

Refactoring code has several advantages and disadvantages. While refactoring this code is a bit time consuming, going back through half of the lines to adjust one by one. However, there are strong advantages in doing so. The refactored code runs much more efficiently as seen in the run time comparisons, and it shortened and simplified the code. With the code simplified and shortened, refactoring it in the future would be easier, and any adjustments in the data would be streamlined. For example, if someone wanted to change the stock tickers to an array of 90, they would only need to adjust the list of ticker names and the length of the array or index range in  7 lines.
