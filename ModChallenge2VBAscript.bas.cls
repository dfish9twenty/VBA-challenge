VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Module2Challenge()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim Ticker As String
Dim PercentChange As Double
Dim Volume As Double
Volume = 0
Dim Row As Double
Row = 2
Dim Column As Integer
Column = 1
Dim i As Long


Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"

OpenPrice = Cells(2, 3).Value


For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        TickerSymbol = Cells(i, 1).Value
        Cells(Row, 9).Value = TickerSymbol
        
        ClosePrice = Cells(i, 6).Value
        
        YearlyChange = ClosePrice - OpenPrice
        Cells(Row, 10).Value = YearlyChange
        
    If (OpenPrice = 0 And ClosePrice = 0) Then
        PercentChange = 0
        
        ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
            PercentChange = 1
        Else
            PercentChange = YearlyChange / OpenPrice
            Cells(Row, 11).Value = PercentChange
            Cells(Row, 11).NumberFormant = "0.00%"
        End If
  
    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    Cells(Row, 12).Value = TotalStockVolume
    
    Row = Row + 1
    
    OpenPrice = Cells(i + 1, 3)
    
    TotalStockVolume = 0
    
    Else
        TotalStockVolume = TotalStockVolume + Cells(1, 7).Value
    End If

Next i

YearChangeLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row

For j = 2 To YearChangeLastRow
    If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
    Cells(j, 10).InteriorColorIndex = 4
    ElseIf Cells(j, 10).Value < 0 Then
        Cells(j, 10).InteriorColorIndex = 3
    End If

Next j

Cells(2, 15).Value = "Greatest Percent Increase"
Cells(3, 15).Value = "Greatest Percent Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For x = 2 To YearChangeLastRow
    If Cells(x, 11).Value = Applicatrion.WorksheetFuncton.Max(WS.Range("K2:K" & YearChangeLastRow)) Then
        Cells(2, 16).Value = Cells(x, 9).Value
        Cells(2, 17).Value = Cells(x, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"
    
    ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearChangeLastRow)) Then
        Cells(3, 16).Value = Cells(x, 9).Value
        Cells(3, 17).Value = Cells(x, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"
    
    ElseIf Cells(x, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearChangeLastRow)) Then
        Cells(4, 16).Value = Cells(x, 9).Value
        Cells(4, 17).Value = Cells(x, 12).Value
        
        End If
        
        Next x
        
        Next WS

End Sub
