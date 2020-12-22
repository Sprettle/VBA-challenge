Attribute VB_Name = "Module1"
Sub stock_data()
    
    'Declare Variable
    
    Dim Ticker      As String
    Dim Stock_Total As Double
    Dim Summary_Table As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As String
    Dim ws          As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        
        'Set Starting Values
        
        Ticker = 0
        Stock_Total = 0
        Yearly_Change = 0
        Percentage_Change = 0
        Summary_Table = 2
        
        'run forloop over tickers to get each type and quantity of stock
        
        Open_Balance = ws.Cells(2, 3).Value
        
        For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Get Types of Ticker
                Ticker = ws.Cells(i, 1).Value
                
                Close_Balance = ws.Cells(i, 6).Value
                
                'Get Yearly Change
                Yearly_Change = Close_Balance - Open_Balance
                
                'Get Percentage Change
                
                If Open_Balance <> 0 Then
                    Percentage_Change = (Yearly_Change / Open_Balance)
                    Percentage_Change = FormatPercent(Percentage_Change, 2)
                End If
                
                ' Add to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                
                ' Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table).Value = Ticker
                
                'Print the Stock total
                ws.Range("L" & Summary_Table).Value = Stock_Total
                
                'Print the Yearly Change
                ws.Range("J" & Summary_Table).Value = Yearly_Change
                
                'Print the Percentage Change
                ws.Range("K" & Summary_Table).Value = Percentage_Change
                
                ' Add one to the summary table row
                Summary_Table = Summary_Table + 1
                
                ' Reset the Stock Total
                Stock_Total = 0
                
                Open_Balance = ws.Cells(i + 1, 3).Value
                
                ' If the cell immediately following a row is the same Ticker...
            Else
                
                ' Add to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        For i = 2 To Cells(Rows.Count, 10).End(xlUp).Row
            
            ' Add Colour conditonal formating
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
            ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
            'ws.Range("Q2") = FormatPercent(Percentage_Change, 2)
            
            ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
            'ws.Range("Q3") = FormatPercent(Percentage_Change, 2)
            
            ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
            
            'For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row
            
            'If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
            'ws.Cells(i, 11).Value = ws.Cells(2, 16)
            'ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
            'ws.Cells(i, 11).Value = ws.Cells(2, 17)
            'End If
            
            'Next i
            
            'Input Headings
            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
            ' Auto Column Width
            ws.Columns("I").AutoFit
            ws.Columns("J").AutoFit
            ws.Columns("K").AutoFit
            ws.Columns("L").AutoFit
            ws.Columns("O").AutoFit
            
        Next ws
        
    End Sub
