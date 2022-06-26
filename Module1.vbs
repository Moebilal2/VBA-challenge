Attribute VB_Name = "Module1"

Sub stock():
For Each ws In Worksheets
Dim opn As Double
Dim diff As Double
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
rownum = 2
totalv = 0
ws.Range("k1").Value = "Ticker"
ws.Range("n1").Value = "Yearly Change"
ws.Range("o1").Value = "Percent Change"
ws.Range("p1").Value = "Total Stock Volume"
ws.Range("l1").Value = "Open price"
ws.Range("m1").Value = "Last Close Price"
ws.Range("q2").Value = "Greatest % Increase"
ws.Range("q3").Value = "Greatest % Decrease"
ws.Range("q4").Value = "Greatest Total Volume"
ws.Range("r1").Value = "Ticker"
ws.Range("s1").Value = "Value"

stockticker = " "
yearlychange = 0
lastprice = 0
percentchange = 0
opn = ws.Range("C" & 2).Value
For i = 2 To lastrow



 If ws.Range("a" & i).Value = ws.Range("a" & i + 1).Value Then
 
ws.Range("l" & rownum).Value = opn
 brandname = ws.Range("a" & i).Value
 ws.Range("k" & rownum).Value = brandname
 totalv = totalv + ws.Range("g" & i).Value
 ws.Range("p" & rownum).Value = totalv
 
 Else
 
 totalv = totalv + ws.Range("g" & i).Value
 ws.Range("p" & rownum).Value = totalv
 lastprice = ws.Range("f" & i).Value
 ws.Range("m" & rownum).Value = lastprice
 totalv = 0
 lastprice = 0
 opn = ws.Range("c" & i + 1).Value
 yearlychange = ws.Range("m" & rownum).Value - ws.Range("l" & rownum).Value
 ws.Range("n" & rownum).Value = yearlychange
 percentchange = ws.Range("n" & rownum).Value / ws.Range("l" & rownum).Value
 ws.Range("o" & rownum).Value = percentchange
 ws.Range("o" & rownum).Style = "percent"
 rownum = rownum + 1
 
 End If
 
 Next i
 lastcolr = ws.Cells(Rows.Count, 15).End(xlUp).Row
 For r = 2 To lastcolr
 If ws.Range("o" & r).Value >= 0 Then
 ws.Range("o" & r).Interior.ColorIndex = 4
 Else
 ws.Range("o" & r).Interior.ColorIndex = 3
 End If
 Next r
 
 '---------------------BONUS-----------------

 
Dim MaxInc As Double
Dim MaxDec As Double
Dim IncIndex As Integer
Dim DecIndex As Integer
Dim totalindex As Integer


MaxInc = WorksheetFunction.Max(ws.Range("o1:o91"))
IncIndex = WorksheetFunction.Match(MaxInc, ws.Range("o1:o91"), 0)
ws.Range("s2").Value = MaxInc
ws.Range("r2").Value = ws.Range("K" & IncIndex)

MaxDec = WorksheetFunction.Min(ws.Range("o1:o91"))
DecIndex = WorksheetFunction.Match(MaxDec, ws.Range("o1:o91"), 0)
ws.Range("s3").Value = MaxDec
ws.Range("r3").Value = ws.Range("K" & DecIndex)

maxtotal = 999999999999#
maxtotal = WorksheetFunction.Max(ws.Range("p1:p91"))
totalindex = WorksheetFunction.Match(maxtotal, ws.Range("p1:p91"), 0)
ws.Range("s4").Value = maxtotal
ws.Range("r4").Value = ws.Range("K" & totalindex)

Next ws
 

End Sub


