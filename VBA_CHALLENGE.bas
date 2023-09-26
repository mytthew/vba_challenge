Attribute VB_Name = "Module1"
Sub StockData()

        ' Go through each of the 3 worksheets
        For Each ws In Worksheets
            
            'Declare all of our variables
            Dim WorksheetName As String
            
            Dim i As Long
            Dim j As Long
            
            Dim Ticker As String
            
            Dim Ticker_Volume As Double
            
            Dim Summary_Row_Table As Integer
            
            Dim YearlyChange As Double
            Dim PercentChange As Double
            Dim GreatestIncrease  As Double
            Dim GreatestDecrease As Double
            Dim HighestVolume As Double
            
            WorksheetName = ws.Name
            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            
            'Set our Summary Table to Row 2
            Summary_Row_Table = 2
            
            'Set the start row to 2
            j = 2
            
            'Identify the last row in Column A
            LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'Begin our For Loop
            For i = 2 To LastRowA
            
                    'Check if the ticker in Column A changed and then paste those tickers in Column I
                    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(Summary_Row_Table, 9).Value = ws.Cells(i, 1).Value
                    
                    'Formula for Yearly Change
                    Yearly_Change = ws.Cells(Summary_Row_Table, 10)
                    ws.Cells(Summary_Row_Table, 10) = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    
                    'Color Code cells based on positive or negative - red will be nagtive and green will be positive
                    If ws.Cells(Summary_Row_Table, 10).Value < 0 Then
                    ws.Cells(Summary_Row_Table, 10).Interior.ColorIndex = 3
                
                    Else
                    ws.Cells(Summary_Row_Table, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Formula for percent change
                    If ws.Cells(j, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    'Format as a Percentage
                    ws.Cells(Summary_Row_Table, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    ws.Cells(Summary_Row_Table, 11).Value = Format(0, "Percent")
                    
                    End If
                
                    'Calculate total volume for each ticker
                    ws.Cells(Summary_Row_Table, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                    
                    'Advance to next ticker in the summary table
                    Summary_Row_Table = Summary_Row_Table + 1
                    
                    'Set the new start row
                    j = i + 1
                    
                    End If
            
        Next i
            
            'Find the last row in Column I
            LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            'Headers for our Summary Table
            HighestVolume = ws.Cells(2, 12).Value
            GreatestIncrease = ws.Cells(2, 11).Value
            GreatestDecrease = ws.Cells(2, 11).Value
        
            'Loop from row 2 to the last row in Column I
            For i = 2 To LastRowI
            
                'Identify highest volume by checking to see if the subsequent value is larger. If it is, then past it into cell (4,16)
                If ws.Cells(i, 12).Value > HighestVolume Then
                HighestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                HighestVolume = HighestVolume
                
                End If
        
                'Same concept as highest volume
                If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                GreatestIncrease = GreatestIncrease
                
                End If
                
                'Same concept as highest volume
                If ws.Cells(i, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                GreatestDecrease = GreatestDecrease
                
                End If
                
            'Place values into the summary table and format them properly
            ws.Cells(2, 17).Value = Format(GreatestIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GreatestDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(HighestVolume, "Scientific")
            
        Next i
            
        Next ws
            
End Sub

