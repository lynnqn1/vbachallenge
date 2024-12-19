Attribute VB_Name = "Module1"
Sub StockAnalysis()

' Identify variables
Dim ws As Worksheet
Dim I As Integer
Dim lastrow As Long
Dim ticker As String
Dim openingPrice As Double
Dim closingPrice As Double
Dim quarterlyCh
ange As Double
Dim percentageChange As Double

' Setting worksheet (with looping through other worksheets)
Set ws = ThisWorkbook.Sheets("A")

' Last row with data
lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Quarterly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    
' Create loop that inputs data into new columns
For I = 2 To lastrow
    ticker = ws.Cells(I, 1).Value
    openingPrice = ws.Cells(I, 2).Value
    closingPrice = ws.Cells(I, 3).Value
    volume = ws.Cells(I, 4).Value
    
    '   Calculating Quarterly Change
    quarterlyChange = closingPrice - openingPrice

    ' Calculating Percentage change
    If openingPrice <> 0 Then
        percentageChange = (quarterlyChange / openingPrice) * 100
    Else
        percentageChange = 0
    End If

    ' Results of calculations
    ws.Cells(I, 9).Value = ticker
    ws.Cells(I, 10).Value = quarterlyChange
    ws.Cells(I, 11).Value = percentageChange
    ws.Cells(I, 12).Value = volume

    '  Conditional formatting of quarterly change
    If Cells(2, 10) > 0 Then
        ws.Cells(I, 10).Interior.ColorIndex = 4
    ElseIf Cells(2, 10) < 0 Then
        ws.Cells(I, 10).Interior.ColorIndex = 3
    End If
Next I
    
End Sub
