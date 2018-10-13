Attribute VB_Name = "Module1"
'Variable Declarations:

Dim TickerName As String
Dim YearValue As String
Dim TotalVolume As Double
Dim VolumeCount As Double
Dim TargetYear As String
Dim Change As Currency
Dim sTickerOpen As Currency
Dim sTickerClose As Currency
Dim tTickerOpenningPrice As Currency
Dim tTickerClosingPrice As Currency

Sub Summarize_Ticker()
For Index = 1 To ActiveWorkbook.Worksheets.Count
worksheetName = ActiveWorkbook.Worksheets(Index).Name
tickercount = 0

Sheets(worksheetName).Activate
Range("A1").Activate
DataRowCount = Sheets(worksheetName).Range("A" & Rows.Count).End(xlUp).Row
Range("I:S").ClearContents

'Add title to summarized table
ActiveCell.Offset(0, 8).Value = "Ticker Name"
ActiveCell.Offset(0, 9).Value = "Year"
ActiveCell.Offset(0, 10).Value = "Volume"
ActiveCell.Offset(0, 11).Value = "Yearly Change"
ActiveCell.Offset(0, 17).Value = "Yearly Open"
ActiveCell.Offset(0, 18).Value = "Yearly Close"

TargetTicker = ActiveCell.Offset(tickercount, 8).Value
TargetYear = ActiveCell.Offset(tickercount, 9).Value
TargetVolume = ActiveCell.Offset(tickercount, 10).Value
YearlyChange = ActiveCell.Offset(tickercount, 11).Value

For i = 1 To DataRowCount

sTickerName = ActiveCell.Offset(i, 0)
sYearValue = Left(ActiveCell.Offset(i, 1), 4)
sVolumeCount = ActiveCell.Offset(i, 6)
sTickerOpen = ActiveCell.Offset(i, 2)
sTickerClose = ActiveCell.Offset(i, 5)

TargetTicker = ActiveCell.Offset(tickercount, 8).Value
TargetYear = ActiveCell.Offset(tickercount, 9).Value
TargetVolume = ActiveCell.Offset(tickercount, 10).Value
tTickerOpen = ActiveCell.Offset(tickercount, 11).Value


If ((TargetTicker <> sTickerName And TargetYear <> sYearValue) Or (TargetTicker = sTickerName And TargetYear <> sYearValue) Or (TargetTicker <> sTickerName And TargetYear = sYearValue)) Then
tickercount = tickercount + 1
    ActiveCell.Offset(tickercount, 8).Value = sTickerName
    ActiveCell.Offset(tickercount, 9).Value = sYearValue
    ActiveCell.Offset(tickercount, 10).Value = sVolumeCount
    
    tTickerOpenningPrice = sTickerOpen
    tTickerClosingPrice = sTickerClose
    
    ActiveCell.Offset(tickercount, 17).Value = tTickerOpenningPrice
    ActiveCell.Offset(tickercount, 18).Value = tTickerClosingPrice
        
    Change = ActiveCell.Offset(tickercount, 18).Value - ActiveCell.Offset(tickercount, 17).Value
    ActiveCell.Offset(tickercount, 11).Value = Change
   
    TargetTicker = ActiveCell.Offset(tickercount, 8).Value
    TargetYear = ActiveCell.Offset(tickercount, 9).Value
    TargetVolume = ActiveCell.Offset(tickercount, 10).Value
    

ElseIf TargetTicker = sTickerName And TargetYear = sYearValue Then
    
    TotalVolume = TargetVolume + sVolumeCount
    ActiveCell.Offset(tickercount, 10).Value = TotalVolume
       
    tTickerClosingPrice = sTickerClose
    ActiveCell.Offset(tickercount, 18).Value = tTickerClosingPrice
    
    Change = ActiveCell.Offset(tickercount, 18).Value - ActiveCell.Offset(tickercount, 17).Value
    ActiveCell.Offset(tickercount, 11).Value = Change
  
End If

Next i

Next Index
End Sub

