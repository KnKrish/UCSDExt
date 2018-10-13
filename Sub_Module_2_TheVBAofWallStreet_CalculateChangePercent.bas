Attribute VB_Name = "Module2"
Sub CalculateChangePer()
For Index = 1 To ActiveWorkbook.Worksheets.Count
worksheetName = ActiveWorkbook.Worksheets(Index).Name
Sheets(worksheetName).Activate
Dim Change As Currency
Dim OpenPrice As Currency
Dim ChangePer As Double

Cells(1, 13).Value = "Change Percent"
Count = Sheets(worksheetName).Range("I" & Rows.Count).End(xlUp).Row
'MsgBox (Count)

For k = 2 To Count - 30
'MsgBox (Cells(k, 12).Value)
Change = Cells(k, 12).Value
OpenPrice = Cells(k, 18).Value
'MsgBox (OpenPrice)
If OpenPrice <> 0 Then
ChangePer = Change / OpenPrice
End If
Cells(k, 13).Value = ChangePer
Cells(k, 13).NumberFormat = "00.0%"
Next k
Next Index

End Sub

