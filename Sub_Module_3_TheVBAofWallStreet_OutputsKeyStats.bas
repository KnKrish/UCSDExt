Attribute VB_Name = "Module3"
Sub LargestStats()
For Index = 1 To ActiveWorkbook.Worksheets.Count
worksheetName = ActiveWorkbook.Worksheets(Index).Name
Sheets(worksheetName).Activate
Dim ChangePer As Double
Dim Volume As Double
Dim LowestPer As Double
Dim HighestPer As Double

Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greates Total Volume"
HighestVolume = Cells(2, 11).Value
LowestPer = Cells(2, 13).Value
HighestPer = Cells(2, 13).Value

sCount = Sheets(worksheetName).Range("I" & Rows.Count).End(xlUp).Row
For k = 2 To sCount
Volume = Cells(k, 11).Value
PercentChange = Cells(k, 13).Value
TickerName = Cells(k, 9).Value

If (HighestVolume < Volume) Then
    HighestVolume = Volume
    Ticker = TickerName
    Cells(4, 16).Value = HighestVolume
    Cells(4, 15).Value = Ticker
End If

If (LowestPer > PercentChange) Then
    LowestPer = PercentChange
    Ticker = TickerName
    Cells(3, 16).Value = LowestPer
    Cells(3, 16).NumberFormat = "00.0%"
    Cells(3, 15).Value = Ticker
End If
 
If (HighestPer < PercentChange) Then
    HighestPer = PercentChange
    Ticker = TickerName
    Cells(2, 16).Value = HighestPer
    Cells(2, 16).NumberFormat = "00.0%"
    Cells(2, 15).Value = Ticker
End If
 
Next k
Range("R:S").ClearContents
Next Index
End Sub


