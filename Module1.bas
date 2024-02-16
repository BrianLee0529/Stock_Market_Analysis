Attribute VB_Name = "Module1"
Sub CreateScript()
    Dim Last_Row As Long
    Dim Total_Stock_Volume As LongLong
    Dim Summary_Table_Row As Integer
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Yearly_Change As Double
    Dim Ticker As String
    Dim WS As Worksheet
    
    For Each WS In ThisWorkbook.Worksheets
    
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
    WS.Cells(1, 9).Value = "Tickers"
    WS.Cells(1, 10).Value = "YearlyChange"
    WS.Cells(1, 11).Value = "PercentChange"
    WS.Cells(1, 12).Value = "TotalStockVoume"
    
    Total_Stock_Volume = "0"
    Yearly_Change = "0"
    Percent_Change = "0"

    Summary_Table_Row = 2
    
    For I = 2 To Last_Row
        Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
        If Cells(I, 1).Value = Cells(I + 1, 1).Value Then
            Closing_Price = Cells(I, 6).Value
        If Opening_Price <> 0 Then
            Percent_Change = (Yearly_Change / Opening_Price) * 100
            ' msgbox("Same Ticker")
        Else
            ' msgbox("Diff Ticker")
            Cells(Summary_Table_Row, 9).Value = Cells(I, 1).Value
            Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
            Cells(Summary_Table_Row, 10).Value = Opening_Price - Cells(I, 6).Value
            Summary_Table_Row = Summary_Table_Row + 1
            ' reset Total_Stock_Volume = 0
            Opening_Price = Cells(I + 1, 3).Value
        End If
Next I
 End Sub
