Attribute VB_Name = "Module11"
Sub StockMarket():

    '----------------------------------------
    'LOOP THROUGH ALL WORKSHEETS
    '----------------------------------------
    
    For Each ws In Worksheets
    
    '----------------------------------------
    'PART I
    'Loop through all stocks and output the following Information:
    'Ticker Symbol (Column I, j = 9)
    'Yearly Change (Column H, j = 10)
    'Percent Change (Column J, j = 11)
    'Total Stock Volume (Column K, j = 12
    'For the Yearly Change column, highlight positive change in green and negative change in red
    '----------------------------------------
    
        'Set variable for holding the ticker symbol (TicSymbol)
        Dim TicSymbol As String
    
        'Set variable for holding the opening Price and determine the initial amount (OpenPrice)
        Dim OpenPrice As Double
        OpenPrice = ws.Cells(2, 3).Value
        
        'Set variable for holding the closing Price and determining the intial amount(ClosePrice)
        Dim ClosePrice As Double
        ClosePrice = 0
    
        'Set variable for holding the stock volumn and determining the intial amount(StockVol)
        Dim StockVol As Double
        StockVol = 0
    
        'Set variable for holding the percent change(PercentChange)
        Dim PercentChange As Double
    
        'Set variable for holding the price change(PriceChange)
        Dim PriceChange As Double
    
        'Set variable for summary table rows and initial row (TabSumRow)
        Dim TabSumRow As Integer
        TabSumRow = 2
        
        'Set variable for the row count (lastrow)
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        'Create Header Titles for the Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'ws.Cells(1, 13).Value = "OpenPrice" '(used to verify data in table was correct)
        'ws.Cells(1, 14).Value = "ClosePrice" '(used to verify data in table was correct)
   
            'Loop through all rows of ticker symbols
            For i = 2 To lastrow
    
            'If the ticker in the current row is NOT the same as the next row, then
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the ticker symbol name, closing value and the total stock volume
                TicSymbol = ws.Cells(i, 1).Value
                ClosePrice = ws.Cells(i, 6).Value
                StockVol = StockVol + ws.Cells(i, 7).Value
            
                'Compute Price Change and Percentage Change
                PriceChange = ClosePrice - OpenPrice
                
                If OpenPrice = 0 Then
                    PercentChange = 0
                
                Else
                    PercentChange = PriceChange / OpenPrice
                
                End If
                            
                'Drop ticker symbol, percent change and total stock volume into summary table
                ws.Range("I" & TabSumRow).Value = TicSymbol
                ws.Range("J" & TabSumRow).Value = PriceChange
                ws.Range("K" & TabSumRow).Value = PercentChange
                ws.Range("L" & TabSumRow).Value = StockVol
                'ws.Range("m" & TabSumRow).Value = OpenPrice '(used to verify data in table was correct)
                'ws.Range("n" & TabSumRow).Value = ClosePrice '(used to verify data in table was correct)
                        
                'Color Format the Percent Change Column
                'If the change is positive, color the column green
                If ws.Range("J" & TabSumRow).Value > 0 Then
                    ws.Range("J" & TabSumRow).Interior.ColorIndex = 4
                    
                'If the change is negative, color the column red
                ElseIf ws.Range("J" & TabSumRow).Value < 0 Then
                    ws.Range("J" & TabSumRow).Interior.ColorIndex = 3
                
                'If there is no change do not color the column
                Else
                    ws.Range("J" & TabSumRow).Interior.ColorIndex = 0
            
                End If
                        
                'Move to the next row on the summary table
                TabSumRow = TabSumRow + 1
            
                'Reset the opening price and the total stock volume
                OpenPrice = ws.Cells(i + 1, 3).Value
                StockVol = 0
            
            'If the ticker in the current row is NOT the same as the next row, then
            Else
         
                StockVol = StockVol + ws.Cells(i, 7).Value
        
            End If
        
    
            Next i
    
            '------------------------------------------
            'PART II BONUS
            '------------------------------------------
       
            'Set variable for holding the greatest % increase (GPercentIncr)
            Dim GPercentIncr As Double
            GPercentIncr = 0
                
            'Set variable for holding the Ticker of the greatest % increase (GTickerIncr)
            Dim GTickerIncr As String
    
            'Set variable for holding the greatest % decrease (GPercentDecr)
            Dim GPercentDecr As Double
            GPercentDecr = 0
    
            'Set variable for holding the Ticker of the greatest % decrease (GTickerDecr)
            Dim GTickerDecr As String
    
            'Set variable for holding the greatest total volume (GVolume)
            Dim GVolume As Double
            GVolume = 0
    
            'Set variable for holding the Ticker of the greatest total volume (GTickerVol)
            Dim GTickerVol As String
    
            'Set variable for the Ticker Column row count (Tlastrow)
            lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
            'Loop through all rows of Ticker Column
            For i = 2 To lastrow
    
            'Create Column and Row titles for the Bonus Table
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % increase"
            ws.Cells(3, 15).Value = "Greatest % decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
                                       
            'Determine the greatest percent Increase Stock
            If GPercentIncr < ws.Cells(i, 11).Value Then
                'Track the greatest % increase and Ticker
                GPercentIncr = ws.Cells(i, 11).Value
                GTickerIncr = ws.Cells(i, 9).Value
  
            End If
            
            'Determine the greatest percent Decrease Stock
            If GPercentDecr > ws.Cells(i, 11).Value Then
                'Track the greatest % increase and Ticker
                GPercentDecr = ws.Cells(i, 11).Value
                GTickerDecr = ws.Cells(i, 9).Value
  
            End If
            
            'Determine the greatest percent Increase Stock
            If GVolume < ws.Cells(i, 12).Value Then
                'Track the greatest % increase and Ticker
                GVolume = ws.Cells(i, 12).Value
                GTickerVol = ws.Cells(i, 9).Value
                
            End If
    
            Next i
    
            'insert the Tickers, Percents and Volume for the Greatest % Increase and % Decrease and Volume
            ws.Cells(2, 16).Value = GTickerIncr
            ws.Cells(2, 17).Value = GPercentIncr
            ws.Cells(3, 16).Value = GTickerDecr
            ws.Cells(3, 17).Value = GPercentDecr
            ws.Cells(4, 16).Value = GTickerVol
            ws.Cells(4, 17).Value = GVolume
            
            
            '----------------------------------------------
            'FORMAT WORKSHEETS
            '-----------------------------------------------
    
            'Format data in columns and cells
            'source: https://stackoverflow.com/questions/27141944/autofit-column-size-in-excel-based-on-content-of-column
            ws.Columns("A:Q").EntireColumn.AutoFit
            ws.Columns("J").EntireColumn.Style = "Currency"
            ws.Columns("K").EntireColumn.Style = "Percent"
            ws.Columns("K").EntireColumn.NumberFormat = "0.00%"
            'source: https://stackoverflow.com/questions/44409090/percent-style-formatting-in-excel-vba
            ws.Columns("L").EntireColumn.Style = "Comma"
            'source: https://stackoverflow.com/questions/44409090/percent-style-formatting-in-excel-vba
            ws.Columns("L").EntireColumn.NumberFormat = "_(* #,###_);_(* (#,###);_(* ""-""??_);_(@_)"
            
            'Format Greatest Table
            ws.Range("Q2:Q3").Style = "Percent"
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("Q4").Style = "Comma"
            ws.Range("Q4").NumberFormat = "_(* #,###_);_(* (#,###);_(* ""-""??_);_(@_)"
   
    
    'Move to next worksheet
    Next ws
    
    'Message Box to inform that Code has successfully run
    MsgBox ("Complete")
    
End Sub
