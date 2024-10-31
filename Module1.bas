Attribute VB_Name = "Module1"
Sub MultipleYearStockData()
    ' Define variables (Worksheet(ws), Ticker Stock Name, Volume, Totals and Integers for rows(i))
    Dim ws As Worksheet
    Dim Ticker_Stock_Name As String
    Dim Next_Ticker_Stock_Name As String
    Dim Stock_Volume As LongLong
    Dim Stock_Volume_Total As LongLong
    Dim i As Long
    Dim Leader_Board_Row As Long
    Dim Last_Row As Long
    
    ' Set Variable for Stocks to be used in calculations
    Dim Stock_Open_Price As Double
    Dim Stock_Closing_Price As Double
    Dim Price_Change As Double
    Dim Percent_Change As Double
  
    For Each ws In ThisWorkbook.Worksheets
        ' Set Column Headers for the four tabs on the workbook
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"

        ' Reset per Ticker
        Stock_Volume_Total = 0
        Stock_Open_Price = ws.Cells(2, 3).Value
        Leader_Board_Row = 2
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To Last_Row
            ' Extract values from workbook
            Ticker_Stock_Name = ws.Cells(i, 1).Value
            Stock_Volume = ws.Cells(i, 7).Value
            Next_Ticker_Stock_Name = ws.Cells(i + 1, 1).Value
    
            ' Check to see if on the same ticker, if not, total out and move to next
            If (Ticker_Stock_Name <> Next_Ticker_Stock_Name) Then
                ' Compute total for ticker
                Stock_Volume_Total = Stock_Volume_Total + Stock_Volume
                ' Calculate Percent Change
                Stock_Closing_Price = ws.Cells(i, 6).Value
                Price_Change = Stock_Closing_Price - Stock_Open_Price
                Percent_Change = Price_Change / Stock_Open_Price
    
                ' Print out  to leaderboard
                ws.Cells(Leader_Board_Row, 12).Value = Stock_Volume_Total
                ws.Cells(Leader_Board_Row, 11).Value = FormatPercent(Percent_Change)
                ws.Cells(Leader_Board_Row, 10).Value = Price_Change
                ws.Cells(Leader_Board_Row, 9).Value = Ticker_Stock_Name
                
                ' Conditional Formatting
                If (Price_Change > 0) Then
                    ws.Cells(Leader_Board_Row, 11).Interior.ColorIndex = 4
                ElseIf (Price_Change < 0) Then
                    ws.Cells(Leader_Board_Row, 11).Interior.ColorIndex = 3
                Else
                    ' Nothing left to do, the row will remain without color
                End If
    
                ' Reset total, Leader_Board_Row
                Stock_Volume_Total = 0
                Leader_Board_Row = Leader_Board_Row + 1
                ' Set the opening price of the NEXT Ticker
                Stock_Open_Price = ws.Cells(i + 1, 3).Value
            Else
                'Calculate total stock volume
                Stock_Volume_Total = Stock_Volume_Total + Stock_Volume
            End If
        Next i
        
        ' Second Loop for Second Leaderboard
        Dim Max_Percent_Increase As Double
        Dim Min_Percent_Decrease As Double
        Dim Max_Total_Stock_Volume As LongLong
        Dim Stock_Price_Max As String
        Dim Stock_Price_Min As String
        Dim Stock_Volume_Max As String
        
        Dim j As Integer
        
        ' Set to first row of the first leaderboard for comparison
        Max_Percent_Increase = ws.Cells(2, 11).Value
        Min_Percent_Decrease = ws.Cells(2, 11).Value
        Max_Total_Stock_Volume = ws.Cells(2, 12).Value
        Stock_Price_Max = ws.Cells(2, 9).Value
        Stock_Price_Min = ws.Cells(2, 9).Value
        Stock_Volume_Max = ws.Cells(2, 9).Value
        
        For j = 2 To Leader_Board_Row - 1
            ' Compare current row to the inits (first row)
            If (ws.Cells(j, 11).Value > Max_Percent_Increase) Then
                ' Max Percent Increase Change
                Max_Percent_Increase = ws.Cells(j, 11).Value
                Stock_Price_Max = ws.Cells(j, 9).Value
            End If
            
            If (Cells(j, 11).Value < Min_Percent_Decrease) Then
                ' New Min Percent Decrease Change
                Min_Percent_Decrease = ws.Cells(j, 11).Value
                Stock_Price_Min = ws.Cells(j, 9).Value
            End If
            
            If (Cells(j, 12).Value > Max_Total_Stock_Volume) Then
                ' New Stock Max Volume Change
                Max_Total_Stock_Volume = ws.Cells(j, 12).Value
                Stock_Volume_Max = ws.Cells(j, 9).Value
            End If
        Next j
        
        ' Write out to Excel Workbook
        ws.Range("O2").Value = Stock_Price_Max
        ws.Range("O3").Value = Stock_Price_Min
        ws.Range("O4").Value = Stock_Volume_Max
        
        ws.Range("P2").Value = FormatPercent(Max_Percent_Increase)
        ws.Range("P3").Value = FormatPercent(Min_Percent_Decrease)
        ws.Range("P4").Value = Max_Total_Stock_Volume
    Next ws
End Sub

