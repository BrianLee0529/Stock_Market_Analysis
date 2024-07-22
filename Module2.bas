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
    
    Total_Stock_Volume = "0"
    Yearly_Change = "0"
    Percent_Change = "0"
    Max_Volume = "0"
    Max_Increase = "0"
    Max_decrease = "0"
    Summary_Table_Row = "2"
    
    WS.Cells(1, 9).Value = "Tickers"
    WS.Cells(1, 10).Value = "YearlyChange"
    WS.Cells(1, 11).Value = "PercentChange"
    WS.Cells(1, 12).Value = "TotalStockVoume"
    WS.Cells(2, 15).Value = "Greatest%Increase"
    WS.Cells(3, 15).Value = "Greatest%Decrease"
    WS.Cells(4, 15).Value = "Greatest Total Volume"
    WS.Cells(1, 16).Value = "Ticker"
    WS.Cells(1, 17).Value = "Value"

   
    Last_Row = WS.Cells(Rows.Count, 1).End(xlUp).Row
    Open_Price = WS.Cells(2, 3).Value
    
    For I = 2 To Last_Row
       
        If WS.Cells(I, 1).Value <> WS.Cells(I + 1, 1).Value Then
            WS.Cells(Summer_Table_Row, 9).Value = WS.Cells(I, 1).Value
            Closing_Price = WS.Cells(I, 6).Value
            Total_Stock_Volume = Total_Stock_Volume + WS.Cells(I, 7).Value
            WS.Cells(Summary_Table_Row, 9).Value = WS.Cells(I, 1).Value
            WS.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
            Yearly_Change = WS.Cells(I, 6).Value - Opening_Price
            WS.Cells(Summary_Table_Row, 10).Value = Yearly_Change
            WS.Cells(Summary_Table_Row, 11).Value = Yearly_Change / Open_Price
            WS.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
            
        If Total_Stock_Volume > Max_Volume Then
            Max_Volume = Total_Stock_Volume
            Max_Volume_Ticker = WS.Cells(I, 1).Value

        End If
        
        If Yearly_Change > Max_Increase Then
            Max_Increase = Yearly_Change
            Max_Increase_Ticker = WS.Cells(I, 1).Value
        End If
        If Yearly_Change < Max_decrease Then
            Max_decrease = Yearly_Change
            Max_decrease_Ticker = WS.Cells(I, 1).Value
        End If
        
        Summary_Table_Row = Summary_Table_Row + 1
        Total_Stock_Volume = 0
        
        If I < Last_Row Then
            Opening_Price = WS.Cells(I + 1, 3).Value
        End If
    Else
        Total_Stock_Volume = Total_Stock_Volume + WS.Cells(I, 7).Value
    End If
    Next I
    
    ' Find the highest and lowest percentage change
    Dim maxPercentageChange As Double
    Dim minPercentageChange As Double
    Dim maxPercentageChangeTicker As String
    Dim minPercentageChangeTicker As String
    
    maxPercentageChange = Application.WorksheetFunction.Max(Range("K2:K" & Summary_Table_Row - 1))
    minPercentageChange = Application.WorksheetFunction.Min(Range("K2:K" & Summary_Table_Row - 1))
    
    maxPercentageChangeTicker = Cells(Application.WorksheetFunction.Match(maxPercentageChange, Range("K2:K" & Summary_Table_Row - 1), 0) + 1, 9).Value
    minPercentageChangeTicker = Cells(Application.WorksheetFunction.Match(minPercentageChange, Range("K2:K" & Summary_Table_Row - 1), 0) + 1, 9).Value
    
    ' Output to cells
    Range("P2").Value = maxPercentageChangeTicker
    Range("Q2").Value = maxPercentageChange
    Range("P3").Value = minPercentageChangeTicker
    Range("Q3").Value = minPercentageChange
    
    ' Find the highest volume
    Dim Max_Volume As Double
    Dim maxVolumeTicker As String
    
    Max_Volume = Application.WorksheetFunction.Max(Range("L2:L" & Summary_Table_Row - 1))
    maxVolumeTicker = Cells(Application.WorksheetFunction.Match(Max_Volume, Range("L2:L" & Summary_Table_Row - 1), 0) + 1, 9).Value
    
    ' Output to cells
    Range("P4").Value = maxVolumeTicker
    Range("Q4").Value = Max_Volume 
    Next WS
 
 End Sub
