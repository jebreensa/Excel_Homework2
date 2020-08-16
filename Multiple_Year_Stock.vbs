Sub Multiple_Year_Stock_Data()
    
   'Declare ws for looping each worksheet
    Dim ws As Worksheet
    
   'Declaring all variables within each worksheet with Autofitting all columns
    For Each ws In Worksheets
    ws.Cells.EntireColumn.AutoFit
    
    
    Dim Ticker_Symbol As String
    Dim LastRow As Long
    Dim L As Long
    Dim Total_Ticker_Volume As Double
    Dim Summary_Table_Row As Long
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Yearly_Open As Double
    Dim Percent_Change As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    
       
   
   'Defining the name of each column used in the summary table
    ws.Range("I1").Value = "Ticker_Symbol"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Total_Stock_Volume"
    ws.Range("O1").Value = " (%) Trends Category"
    ws.Range("O2").Value = "Greatest_Increase(%)"
    ws.Range("O3").Value = "Greatest_Decrease(%)"
    ws.Range("O4").Value = "Greatest_Total_Volume"
    ws.Range("P1").Value = "Ticker_Symbol"
    ws.Range("Q1").Value = "Value in (%)"
    
   'Declaring initial values for each calcualted value
    Total_Ticker_Volume = 0
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    Summary_Table_Row = 2
    'Similar to having LastCol but since we do not want to be limited by the number of columns and keeping the loop for the open price=i+1 once the if statement is met
    L = 2
       

    'Defining the number of rows for each loop
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   

        
    'The First Loop is used to run through all tickers within each year gathering: Ticker Symbol, Yearly Change, Percent Change and Total Stock Volume
    For i = 2 To LastRow
       
    'Total_Ticker_Volume as intial calcualtion running through each ws within the loop:
        Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
        
    'First If Statement includes defining when is the next change to input each ticker as a header then calculate the following: Yearly Change, Total_Stock_Volume
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

          
            Ticker_Symbol = ws.Cells(i, 1).Value
           
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
         
            ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
        
            Total_Ticker_Volume = 0
            
           'Defining the opening and closing price: Opening will loop ahead of the closing price since we are defining the beginning and end year
            Opening_Price = ws.Range("C" & L)
          
            Closing_Price = ws.Range("F" & i)
            
            Yearly_Change = Closing_Price - Opening_Price
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
            'Number Foramting Command
            ws.Range("J" & Summary_Table_Row).NumberFormat = "$0.00"

            ' Define an If statment for no change if the open price=0
            If Opening_Price = 0 Then
               Percent_Change = 0
                    
           'Percent Change calcualted within the same Ticker Symbol
                Else
                Yearly_Opening_Price = ws.Range("C" & L)
                Percent_Change = Yearly_Change / Yearly_Opening_Price
                        
            End If
                
           'Formatting the Percent Change Values (%)and highlighting the yearly change to determine whether we have negative or positve values
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                
           
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

        
            If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                Else
          
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
            'Determine the The Summary Table Row Loop for each range and the range for the open price from the closing price
  
            Summary_Table_Row = Summary_Table_Row + 1
            L = i + 1
              
                
        End If
                
       
        Next i
        
       'The second loop is to calcualte the greatest increase, greatest decrease, and Greatest Total Volume with formatting the numbers in (%)
                       
        For i = 2 To LastRow
            
         
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If

         
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
                    
            End If

            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
                    
            End If

            Next i
            
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

'Loop Through all sheets
    Next ws

End Sub







