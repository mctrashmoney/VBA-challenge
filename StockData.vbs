Sub StockData()
    Dim ws As Worksheet
    Dim i As Long
    Dim ticker As String
    Dim opening_price As Double
    Dim closing_price As Double
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim lastrow As Long
    Dim summarytablerow As Long
    Dim first_row As Long
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_ticker As String
    Dim min_ticker As String
    Dim max_volume As Double
    Dim max_volume_ticker As String
    
    ' Loop through each worksheet in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Activate the worksheet
        ws.Activate
        
        ' Initialize variables
        quarterly_change = 0
        total_volume = 0
        summarytablerow = 2
        max_percent = 0
        min_percent = 0
        max_volume = 0
        
        ' Find the last row with data in column A
        lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set headers for the summary table
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Quarterly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        
        ' loop through each row to calculate quarterly change
        For i = 2 To lastrow
            ' Check if it's the first row for the ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                opening_price = ws.Cells(i, 3).Value
                first_row = i
            End If
            
            ' check if the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                
                ' closing price at the end of the quarter
                closing_price = ws.Cells(i, 6).Value
                
                ' quarterly change
                quarterly_change = closing_price - opening_price
                
                ' so the percent change doesnt error 
                If opening_price <> 0 Then
                    percent_change = quarterly_change / opening_price
                Else
                    percent_change = 0
                End If
                
                ' TOTAL volume
                total_volume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(first_row, 7), ws.Cells(i, 7)))
                
                ' ADD results to the summary table
                ws.Range("J" & summarytablerow).Value = ticker
                ws.Range("K" & summarytablerow).Value = quarterly_change
                ws.Range("L" & summarytablerow).Value = percent_change
                ws.Range("L" & summarytablerow).NumberFormat = "0.00%"
                ws.Range("M" & summarytablerow).Value = total_volume
                
                ' FORMATTING
                If quarterly_change > 0 Then
                    ws.Range("K" & summarytablerow).Interior.Color = RGB(0, 250, 0)
                ElseIf quarterly_change < 0 Then
                    ws.Range("K" & summarytablerow).Interior.Color = RGB(250, 0, 0)
                End If
                
                ' highest and lowest percent change
                If percent_change > max_percent Then
                    max_percent = percent_change
                    max_ticker = ticker
                End If
                
                If percent_change < min_percent Then
                    min_percent = percent_change
                    min_ticker = ticker
                End If
                
                ' check for maximum total volume
                If total_volume > max_volume Then
                    max_volume = total_volume
                    max_volume_ticker = ticker
                End If
                'next row
                summarytablerow = summarytablerow + 1
                
                ' reset variables
                quarterly_change = 0
                total_volume = 0
            End If
        Next i
        
        ' Percentages on table on the side
        ws.Range("Q2").Value = max_ticker
        ws.Range("R2").Value = max_percent
        ws.Range("R2").NumberFormat = "0.00%"
        
        ws.Range("Q3").Value = min_ticker
        ws.Range("R3").Value = min_percent
        ws.Range("R3").NumberFormat = "0.00%"
        
        ' Total stock bvolume on stock volume on the side
        ws.Range("Q4").Value = max_volume_ticker
        ws.Range("R4").Value = max_volume
    Next ws
End Sub

