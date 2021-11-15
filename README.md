# Stock-Analysis - VBA of Wall Street 
# Overview of Project 
For the given dataset of different stocks and their information relating to the performance have been analysed initialy to assess the performance of the each stocks for the given years. By clicking the button "Run analysis for all the stocks " the entire data set can be analysed. As the given data set is not included the whole stock market, the main pupose of this analysis to find out if the whole data set is expanded in the stock market in the last few years and determine how fast that the code can be run to get the desired results by refatoring the initial VBA script in the given data set. For that purpose the given data set carring 12 stocks for 2017 and 2018 will be analysed to see their performance and run time for future investments. Please see the attached link for the original data set and output worksheet. 



# Results 
## Comparison of stock performance in 2017 and 2018

As per the below comparison of the 2 tables it cleary shows the total daily volume and the return of each fund in the given 2 years. It had been a really a good year in 2017 for all the given stocks except the TERP. The stocks ENPH and RUN had been performed very good in both years. However it indicates that eventhough the total daily volume increases it still can be given a negative or less return over a year. For example the stocks DQ, HASI, SEDG, VSLR shows a negative return while ENPH a less return eventhough the volume has increased.

<img width="271" alt="Performance - 2017" src="https://user-images.githubusercontent.com/93173498/141703639-611b65bc-5f38-4165-b4da-6fc03504ce95.png">



<img width="241" alt="Performance - 2018" src="https://user-images.githubusercontent.com/93173498/141703692-e42a035b-774b-499b-9ed8-08b3ab56f1b6.png">

## Refactored Run time 

The refactored codes took 0.38 seconds to run in 2018 and took 0.35 seconds in 2017. It was less time than the original code run time. 
See the attached images of the refactored run time for each year. 

<img width="282" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/93173498/141704478-16a26154-1888-4b6f-a43a-62ab99323415.png">

<img width="341" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/93173498/141704488-e0b44497-4f91-400a-aead-b478abb06a0e.png">

## Initial code vs refactored code 
* As per the initial code, the search is done by one ticker at a time. The search has to be finished to start a new ticker. 

* As per the refactored code , the search is read any ticker, identify the value , add the volume in the same index as the code consists new 3 output arrays.
See the attached refactored code for the refernece. 

   
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
   
   
 Next i
        


    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
    tickerVolumes(i) = 0
   
 Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
       If Cells(i, 1).Value = tickers(tickerIndex) Then
    
        '3a) Increase volume for current ticker
        
        
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
  End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
        
        Worksheets("All Stocks Analysis").Activate
        
        For i = 0 To 11
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
   
   
   
   
  # Summary 
  
   
   ## Advantages and disadvantages of refactoring the code in general.
   
   * Advantages - Save time with the analysis as it runs faster than the original.
   * Disadvantage - Time consuming to write the codes. 
   
   ## Advantages and disadvantages of the original and refactored VBA script.
   * Advantages - Save time with analysis as it runs faster than the original as mentioned above. If this code work for more than 12 stocks it would be great time saving          to do a analysis based on the retun and the volume. 
   * Disadvantage - The new refactored code run for the given sorted data set. if it is unsorted, a new code has to be added to sort the data. Also it is time consuming if        we have to add more names to tickers if there are so many new stock names to be added. 

   
    
    



