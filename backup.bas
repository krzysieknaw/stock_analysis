Attribute VB_Name = "Module1"
Option Explicit

Sub MacroCheck()
    Dim testMessage As String
    testMessage = "hello world"
    MsgBox (testMessage)

   MsgBox ("hello worlds")
    
End Sub

Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'create a header row using cells/integer values
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'create a header row using range/string values
    Range("A4").Value = "Year"
    Range("B4").Value = "Total Daily Volume"
    Range("C4").Value = "Return"
    

End Sub
'how to find end row in large data sheet
Sub CountRows1()
'activate sheet, name is a string in this instance
Worksheets("2018").Activate
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox (last_row)
End Sub
'for loop to return column name values

Sub ForLoop()

'activate sheet, name is a string in this instance
Worksheets("2018").Activate

'establish variable
Dim i As Integer
'loop using cell integer value to return values
For i = 1 To 8
    MsgBox (Cells(1, i))
'close loop
Next i

End Sub

Sub totalVolume()


Worksheets("2018").Activate
'declaring variable generally frees up more memory than default

Dim totalVolume As Integer
totalVolume = 0

Dim rowStart As Long
rowStart = 2

Dim rowEnd As Long
rowEnd = 3013

Dim i As Integer

For i = rowStart To rowEnd
    'increase totalVolume
    'totalVolume = totalVolume + Cells(i, 8).Value

Next i

End Sub


