# stock-analysis
Analyzing stock data

## Overview of Project: 
Steve, a recent finance degree graduate, had previously requested to help him analyzing stocks data in Wall Street for a green energy company DAQO New Energy Corp (Wall Street ticker = DQ) for his parents toknow how well DQ was actively traded in 2018. He wanted to analyze a handful of other companies in addition to DAQO Company. We have helped him by developing a Visual Basic Applications (VBA) macro to automate this process. In addition, We have also created a button to click so that he can analyze an entire dataset easily. He was happy with the workbook that was prepared for him. He wanted to help him again to do a bit more research for his parents and he wants to expand the dataset to include the entire stock market data in the spreadsheet over the last two years. The previous code worked well for analyzing a small subset of data (in that case, we only analyzed companies in 2018). It might become harder to analyze many thousands of stocks using the code and it may even take a long time to execute the code. 

In this challenge project, we have edited, or refactored, the code, Module2_VBA_Script, to loop through all the data onetime for collecting similar information. Then, we have performed the analysis, and evaluated the performance of the VBA script by measuring the the taken to execute the code. Refactoring is a key part during the coding process. The goal was to make to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

## Results:

The spreasdsheet with the whole analysis was uploaded to the GitHub and tthe link to the spreassheet (macro enabled) is included here: [VBA-Challenge](VBA_Challenge.xlsm). The refractired VB script is available in the followiung link: [VBA-Challenge-Script](VBA_Challenge.vbs)

### Improvements to Code (Refactoring)

In the refactored code, the following improvements were made by using a ticker index and arrays. Using an the arrays that used indexes to store data improved the efficiency of running the code and are briefly explained below. 

    a) a ticker index (tickerIndex) was created and set equal to zero before iterating over all the rows: tickerIndex = 0
    b) Three output arrays for storing data (tickerStartingPrices, and tickerEndingPrices. tickerVolumes) were created. 

    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
    c) Used a "for" loop to initiset to set equal to tickerVolumes = 0
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    
    d) Using tickerIndex, the current ticker volume for each ticker is increased by using the formula: 
    
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    e) checked if the current row is the first row with the selected tickerIndex
    
     If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
    
    f) similarly, checked if the current row is the last row with the selected tickerIndex
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
    g) Looped through the arrays to output the Ticker, Total Daily Volume, and Return.
    
   
### Analysis of the stocks for 2017 and 2018

The results of this analysis showed that the DAQO Company performed well in 2017 (199.4%) followed by the company with the ticker SEDG. Among the top most companies those had the largest positive returns in 2017 were DQ (199.4%), SEDG (184.5%), ENPH (129.5%). However, DAQO performance of DAQO was bad in 2018. Almost all companies workedwell in 2017 and only two companies showed positive returns (81.9% for ENPH, 84.0% for RUN) in 2018.  

### Performance of VBA Scripts
The time taken to run the whole analysis for 2017 data was 0.04 s (40 ms). Figure 1 below shows the pop up. ![VBA-Challenge - 2017 - time](/resources/VBA_Challenge_2017.png). The analysis time taken before refactoring for 2017 data was 0.26 s.It was 6.5 times faster after refactoring the code.

### Figure 1: Screenshot showing the time needed to analyze the entire datasheet for 2017 (3012 data).

The time taken to run the whole analysis for 2017 data was 0.05 s (50 ms). Figure 2 below shows the pop up. The analysis time taken before for 2018 data was 0.25 s. It was 5 times faster after refactoring.

![VBA-Challenge - 2018 - time](/resources/VBA_Challenge_2018.png)

### Figure 2: Screenshot showing the time needed to analyze the entire datasheet for 2018.

## Summary:

In summary, advantages of refactoring a code includes the following: 1) the running time for the code becomes faster, 2) data in variables can be stored in memory efficiently when using indecies. 3) calculations are more effective and faster. Some of the disadvantages include the following: 1) there may be situations of plagiarism when refactoring someones code, 2) any intellectual property issues may arise when refactoring.3) if there is an undetected error in the code (especially in formulas for calculations, even though the code produces results before refactoring), it can easily be missed and remained in the code undetected UNLESS the code is thoroughly checked (quality check step).   

In our analysis, we have seen a much faster code after refactoring the code. It is definetely an advantage (pros). However, I did not see any disadvantages (cons) of using refactoring the VBA script in this analysis. 





