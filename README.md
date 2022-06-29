#Refactored Stock Analysis

##Project Overview
1.In this project I used data from twelve different stocks to analyze the performance of the stocks during the years of 2018 and 2017. The challenge was to refactor the code to shorten the amount of time that it takes for the code to run the analysis. This can be helpful in the world of data analytics because when writing code for the first time may be clunky and need some revisions to run smoother. 

###Process
1. I utilized the starter code that was provided at the end of the module. With this code there we several additions that I needed to add.
- First, I needed to add a ticker index and set it equal to zero (tickerIndex = 0)
- Next, I created three output arrays 
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
2. Once I had done that, I now needed to create a for loop that would initialize the tickerVolumes to zero
- For i = 0 To 11       
        tickerVolumes(i) = 0    
  Next i
- Now it was the time to create a nested for loop that would run through all the data and populate the necessary data for each stock.
- To do this I started the loop to count from the first row of data. (For i = 2 To RowCount)
- Next, I needed to increase the volume of the current ticker
- tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
- From here I needed to check if the current row is the first row with the selected ticker index. 
- If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then      
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value           
  End If
- After that I needed to check if the current row was the last row with the selected ticker index
- If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then      
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
- Next I needed the for loop to increase the tickerIndex by 1 and close the nested loop
- tickerIndex = tickerIndex + 1
    End If
 Next i
 
####Analysis
1. I found that the code ran much faster than the code that was taught within the modules.Although the code functioned and was accurate it was unfortunately very slow to produce a result. When the code was refactored with steps above the results were much faster.  

![Moudule_Analysis](https://github.com/mselover21/StockAnalysis/blob/main/Module_Analysis.PNG)

![VBA_Challenge_2017]
![VBA_Challenge_2018]

#####Summary
1. Refactoring code has many advantages as well as disadvantages. Among the advantages is the reduction of time that it takes for analysis to take place. This when applied in the real world is crucial when you are dealing with increasingly large sets of data. I also found that the refactoring of code can simplify it and make it easier for others who you may be working with know what each line of code is used for. Among the disadvantages I found the amount of time it took to write the code far longer than I had anticipated. I am learning an entirely new language and can estimate that that was part of the issue. I do feel that even with the length of time that it took to refactor that it is certainly well worth the effort when applied to real world scenarios.
2. The original code that we used had a couple of advantages and disadvantages. I feel strongly that writing original code was better for the learning experience. It was broken down further than the refactored code which was beneficial for me to understand this new language. After refactoring the code, I see that the disadvantage comes from the length of time it takes to create an overly defined code. This could be helpful in a real-world scenario in relation to creating a rough draft for the code. However, to produce a clean well written and faster code would be better when utilizing very large sets of data.
