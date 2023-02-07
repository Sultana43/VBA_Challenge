Attribute VB_Name = "Module1"
Sub Stockbroker()

Dim ws As Worksheet

For Each ws In Worksheets

'setting title

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

'ticker output
'yearly output
'percentage change
'total stock volume

Dim i As Double
Dim ticker As String
Dim TotalVolume As Double
Dim j As Double
Dim Change As Double
Dim PercentChange As Double
Dim x As Double


TotalVolume = 0
j = 0
Change = 0
PercentChange = 0
x = 2


LastRowIndex = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To LastRowIndex

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

TotalVolume = TotalVolume + ws.Cells(i, 7).Value

Change = ws.Cells(i, 6).Value - ws.Cells(x, 3).Value

PercentChange = Change / ws.Cells(x, 3).Value






ws.Range("I" & j + 2).Value = ws.Cells(i, 1).Value

ws.Range("J" & j + 2).Value = Change

ws.Range("J" & j + 2).NumberFormat = "0.00"

ws.Range("K" & j + 2).Value = PercentChange

ws.Range("K" & j + 2).NumberFormat = "0.00%"


ws.Range("L" & j + 2).Value = TotalVolume

If ws.Range("J" & j + 2).Value >= 0 Then
ws.Range("J" & j + 2).Interior.Color = vbGreen
ElseIf ws.Range("J" & j + 2).Value < 0 Then
ws.Range("J" & j + 2).Interior.Color = vbRed

    End If


TotalVolume = 0

j = j + 1

x = i + 1

Change = 0






Else

TotalVolume = TotalVolume + ws.Cells(i, 7).Value

End If




' ticker = ws.Cells.ws.Range("A" & 1 + 1) = "AAB"
' ws.Range("I" & i + 1) = ticker
' ws.Range("J" & i + 1) = ws.Range("C2" & i + 1) - ws.Range("F" & i + 1)
' ws.Range("K" & i + 1) = ws.Range("C2" & i + 1) / ws.Range("F" & i + 1) * 10
' ws.Range("L" & i + 1) = ws.Range("G" & i + 1)

Next i

Next ws

End Sub

