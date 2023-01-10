Attribute VB_Name = "Yearly_Stock_Analysis"
Sub yearly_stock_analysis()

For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1") = "Ticker Symbol"
    ws.Range("J1") = "Yearly Difference"
    ws.Range("K1") = "Percent Difference"
    ws.Range("L1") = "Total Stock Volume"

    'Ticker Symbol variables
    Dim ticker_symbol As String
    Dim ticker_symbol_row As Integer

    ticker_symbol_row = 2

    'Yearly Difference variables
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_difference As Double
    Dim yearly_difference_row As Integer

    yearly_difference_row = 2

    'Perecent Difference variables
    Dim percent_difference As Double
    Dim percent_difference_row As Double

    percent_difference_row = 2

    'Total Volume variables
    Dim total_volume As Double
    Dim total_volume_row As Integer

    total_volume = 0
    total_volume_row = 2

    For i = 2 To lastrow

    'Ticker Symbol Column
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
            ticker_symbol = ws.Cells(i, 1).Value
    
            ws.Range("I" & ticker_symbol_row).Value = ticker_symbol
    
            ticker_symbol_row = ticker_symbol_row + 1
        
        End If
    
    'Yearly Difference and Percent Difference Columns
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Range("I" & yearly_difference_row).Value = ws.Cells(i, 1).Value Then
            opening_price = ws.Cells(i, 3).Value

        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Range("I" & yearly_difference_row).Value = ws.Cells(i, 1).Value Then
            closing_price = ws.Cells(i, 6).Value
     
            yearly_difference = closing_price - opening_price
            percent_difference = (yearly_difference / opening_price)
        
            ws.Range("J" & yearly_difference_row).Value = yearly_difference
            ws.Range("K" & percent_difference_row).Value = percent_difference
        
            ws.Range("K" & percent_difference_row).NumberFormat = "0.00%"
        
            yearly_difference_row = yearly_difference_row + 1
        
            percent_difference_row = percent_difference_row + 1
        
        End If

    'Total Volume Column
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            total_volume = total_volume + ws.Cells(i, 7).Value
    
            ws.Range("L" & total_volume_row).Value = total_volume
    
            total_volume_row = total_volume_row + 1
    
            total_volume = 0
    
        Else
    
            total_volume = total_volume + ws.Cells(i, 7).Value
    
        End If

    Next i

    'Conditional formatting for yearly difference and percent difference columns
    For i = 2 To lastrow

        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ws.Cells(i, 11).Interior.ColorIndex = 4
    
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
    
        End If

    Next i

    'Additional Calculations
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double

    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0

    For i = 2 To lastrow

        If ws.Cells(i, 11).Value > greatest_increase Then
            greatest_increase = ws.Cells(i, 11).Value
    
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
            ws.Range("Q2").NumberFormat = "0.00%"
    
        End If
    
        If ws.Cells(i, 11).Value < greatest_decrease Then
            greatest_decrease = ws.Cells(i, 11).Value
    
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
            ws.Range("Q3").NumberFormat = "0.00%"
    
        End If
    
        If ws.Cells(i, 12).Value > greatest_volume Then
            greatest_volume = ws.Cells(i, 12).Value
    
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
    
        End If

    Next i

    'Labeling Additional Calculations
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

Next ws

End Sub
