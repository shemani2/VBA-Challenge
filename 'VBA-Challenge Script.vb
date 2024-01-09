'VBA-Challenge Script
Sub MarketStockAnalysis():

            'Name variables
                Dim TickerName As String
                Dim TickerVolume As Double
                Dim SummaryRow As Integer
                Dim OpeningPrice As Double
                Dim ClosingPrice As Double
                Dim YearlyChange As Double
                Dim PercentChange As Double

            'Define total volume
                TickerVolume = 0
            'Define location of summary row of ticker; 2 represents columns
                SummaryRow = 2

            'Define opening price location and value
                OpeningPrice = Cells(2, 3).Value
    
            'Summary Table Headers
                Cells(1, 9).Value = "Ticker Symbol"
                Cells(1, 10).Value = "Yearly Change"
                Cells(1, 11).Value = "Yearly % Change"
                Cells(1, 12).Value = "Total Stock Volume"

            'Find number of rows in first column to define last row; ( reference from class solved code)
                LastRow = Cells(Rows.Count, 1).End(xlUp).Row

            'Create loop of ticker names from beginning to end (last row)
                'for loop defined from 2nd row to last row
                For i = 2 To LastRow
                         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                         
                            'Define TickerName
                            TickerName = Cells(i, 1).Value
                            
                            'Add TickerVolume to value in "vol" column
                            TickerVolume = TickerVolume + Cells(i, 7).Value
                            
                            'Add TickerName to Summary Table
                            Range("I" & SummaryRow).Value = TickerName
                            
                            'Add TickerVolume for each ticker to Summary Table
                            Range("L" & SummaryRow).Value = TickerVolume
                            
                            'Define ClosingPrice
                            ClosingPrice = Cells(i, 6).Value
                            
                            'Add YearlyChange to Summary Table
                            Range("J" & SummaryRow).Value = YearlyChange
                            
                            'Calculate YearlyChange
                            'YearlyChange is difference between closing and opening price
                            YearlyChange = ClosingPrice - OpeningPrice
                            
                            'Calculate PercentChange and define non-divisble numbers
                            If OpeningPrice = 0 Then
                                PercentChange = 0
                            Else
                                PercentChange = YearlyChange / OpeningPrice
                            End If
                
                            'Add PercentChange to Summary Table for each row
                            Range("K" & SummaryRow).Value = PercentChange
                            Range("K" & SummaryRow).NumberFormat = "0.00%"
                            
                            'Reset row counter by adding one to SummaryRow
                            SummaryRow = SummaryRow + 1
                            
                            'Reset trade volume to zero
                            TickerVolume = 0
                            
                            'Reset OpeningPrice
                            OpeningPrice = Cells(i + 1, 3)
                    
                    Else
                            'Add trade volume
                            TickerVolume = TickerVolume + Cells(i, 7).Value
                    
                    End If
    
                Next i

    'Find the last row of summary table
    SummaryTableLastRow = Cells(Rows.Count, 9).End(xlUp).Row

    'Conditional Formatting: positive change is green and negative change is red
        For i = 2 To SummaryTableLastRow
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        'Label cells to find greatest % increase, greatest % decrease, and greatest total volume
            Cells(2, 14).Value = "Greatest % Increase"
            Cells(3, 14).Value = "Greatest % Decrease"
            Cells(4, 14).Value = "Greatest Total Volume"
            Cells(1, 15).Value = "Ticker"
            Cells(1, 16).Value = "Value"
        
        'Calculate max and min values for percent change
        
            For i = 2 To SummaryTableLastRow
            
                'Greatest % Increase
                If Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & SummaryTableLastRow)) Then
                    Cells(2, 15).Value = Cells(i, 9).Value
                    Cells(2, 16).Value = Cells(i, 11).Value
                    Cells(2, 16).NumberFormat = "0.00%"
                
                'Greatest % decrease
                ElseIf Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & SummaryTableLastRow)) Then
                            Cells(3, 15).Value = Cells(i, 9).Value
                            Cells(3, 16).Value = Cells(i, 11).Value
                            Cells(3, 16).NumberFormat = "0.00%"
                            
                'Greatest Total Volume
                ElseIf Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & SummaryTableLastRow)) Then
                            Cells(4, 15).Value = Cells(i, 9).Value
                            Cells(4, 16).Value = Cells(i, 12).Value
                End If
                
            Next i
           
End Sub
