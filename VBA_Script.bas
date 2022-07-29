Attribute VB_Name = "Module1"
Sub stock_summary()

For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

Dim Ticker_Symbol As String

Dim Stock_Volume As LongLong
Stock_Volume = 0

Dim Open_Price As Double
Open_Price = ws.Cells(2, 3).Value

Close_Price = 0

Dim Price_Change  As Double


Dim Percent_Change As Double


Dim Summary_Table_Row As Long
Summary_Table_Row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For I = 2 To LastRow
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
        Ticker_Symbol = ws.Cells(I, 1).Value
        Stock_Volume = Stock_Volume + ws.Cells(I, 7).Value
        
        Close_Price = ws.Cells(I, 6).Value
        Price_Change = Close_Price - Open_Price
        Percent_Change = Price_Change / Open_Price
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
        ws.Range("J" & Summary_Table_Row).Value = Price_Change
        ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
         ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
        

        Summary_Table_Row = Summary_Table_Row + 1
        
        Stock_Volume = 0
        Open_Price = ws.Cells(I + 1, 3)
        
        Else
            Stock_Volume = Stock_Volume + ws.Cells(I, 7).Value
        
    End If
        
    Next I
    
LastRow_Summary = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For I = 2 To LastRow_Summary
        If ws.Cells(I, 10).Value > 0 Then
            ws.Cells(I, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(I, 10).Interior.ColorIndex = 3
        End If
    Next I
    
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Columns("O:Q").AutoFit
    
    For I = 2 To LastRow_Summary
        If ws.Cells(I, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow_Summary)) Then
            ws.Cells(2, 16).Value = ws.Cells(I, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(I, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
        ElseIf ws.Cells(I, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow_Summary)) Then
            ws.Cells(3, 16).Value = ws.Cells(I, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(I, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
        ElseIf ws.Cells(I, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow_Summary)) Then
            ws.Cells(4, 16).Value = ws.Cells(I, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(I, 12).Value
        End If
    Next I
    

Next ws
    
End Sub
