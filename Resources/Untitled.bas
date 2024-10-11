Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()

    ' Variables
    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim Quarterly_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Percent_Change As Double
    Dim Summary_Row As Integer
    Dim EndRow As Long
    Dim previous_i As Long
    Dim Increase As Double
    Dim Decrease As Double
    Dim Greatest As Double
    Dim increase_name As String
    Dim decrease_name As String
    Dim greatest_name As String
    
    '-----------------------------
    ' Loop through each worksheet
    For Each ws In Worksheets
    
        ' Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly_Change"
        ws.Range("K1").Value = "Percent_Change"
        ws.Range("L1").Value = "Total_Stock_Volume"
        
        ' Initialize variables
        Summary_Row = 2
        previous_i = 2
        Total_Stock_Volume = 0
        
        ' Last row to loop through
        EndRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To EndRow
            ' Check if new ticker is found
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                ' First open price
                OpenPrice = ws.Cells(previous_i, 3).Value
                ' Last close price
                ClosePrice = ws.Cells(i, 6).Value
                
                ' Calculate total stock volume
                For j = previous_i To i
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                Next j
                
                ' Calculate percentage change
                If OpenPrice <> 0 Then
                    Percent_Change = (ClosePrice - OpenPrice) / OpenPrice
                Else
                    Percent_Change = 0
                End If
                
                ' Calculate quarterly change
                Quarterly_Change = ClosePrice - OpenPrice
                
                ' Write values to summary table
                ws.Cells(Summary_Row, 9).Value = Ticker
                ws.Cells(Summary_Row, 10).Value = Quarterly_Change
                ws.Cells(Summary_Row, 11).Value = Percent_Change
                ws.Cells(Summary_Row, 11).NumberFormat = "0.00%"
                ws.Cells(Summary_Row, 12).Value = Total_Stock_Volume
                
                ' Reset for next ticker
                Summary_Row = Summary_Row + 1
                Total_Stock_Volume = 0
                previous_i = i + 1
            End If
        Next i
        
        '----------------------------------------------------------------------

        ' Find greatest increase, decrease, and total volume
        Increase = 0
        Decrease = 0
        Greatest = 0
        
        kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        For k = 2 To kEndRow
            current_k = ws.Cells(k, 11).Value
            volume = ws.Cells(k, 12).Value
            
            ' Find greatest percentage increase
            If current_k > Increase Then
                Increase = current_k
                increase_name = ws.Cells(k, 9).Value
            End If
            
            ' Find greatest percentage decrease
            If current_k < Decrease Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value
            End If
            
            ' Find greatest total volume
            If volume > Greatest Then
                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value
            End If
        Next k
        
        ' Output results
        ws.Range("N1").Value = "Column Name"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker Name"
        ws.Range("P1").Value = "Value"
        
        ws.Range("O2").Value = increase_name
        ws.Range("O3").Value = decrease_name
        ws.Range("O4").Value = greatest_name
        ws.Range("P2").Value = Increase
        ws.Range("P3").Value = Decrease
        ws.Range("P4").Value = Greatest
        
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        
        ' Conditional formatting for quarterly change
        jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        For j = 2 To jEndRow
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4 ' Green for positive
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red for negative
            End If
        Next j
    
    Next ws

End Sub



