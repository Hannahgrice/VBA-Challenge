Attribute VB_Name = "Module1"
Sub remove_ticker_duplicates()

    Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    

End Sub

Sub move_ticker()

    Columns("A:A").Copy

    Columns("I:I").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
     Application.CutCopyMode = False
    Columns("I:I").RemoveDuplicates Columns:=1, Header:=xlNo

End Sub


Sub Yearly_change()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("A")

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    For i = 2 To lastRow
        ws.Cells(i, "J").Value = ws.Cells(i, "C").Value - ws.Cells(i, "F").Value
        
    Next i

End Sub

Sub Percent_change()

  Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("A")
    
    columnC = "C"
    columnJ = "J"
    resultColumn = "K"
    
    ' Loop through these rows to multiply values from "c" and "j"
    
lastRow = ws.Cells(ws.Rows.Count, columnC).End(xlUp).row

Dim row
For row = 1 To lastRow
    Dim valueC, valueJ, resultPercentage
    valueC = ws.Cells(row, columnC).Value
    valueJ = ws.Cells(row, columnJ).Value
    
    If IsNumeric(valueC) And IsNumeric(valueJ) Then
        resultPercentage = (valueC * valueJ) * 100 ' Multiplying values and making to percentage
        ws.Cells(row, resultColumn).Value = resultPercentage & "%" ' puting result as a percentage
    Else
        ws.Cells(row, resultColumn).Value = "N/A"
    End If
Next
    

    
End Sub

Sub Total_stock()

 Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("A")
    
    Dim volumeColumn, timeColumn, resultColumn
    
volumeColumn = "G"
timeColumn = "B"
resultColumn = "L"

lastRow = ws.Cells(ws.Rows.Count, volumeColumn).End(xlUp).row

Dim row
For row = 1 To lastRow
    Dim volumeValue, timeValue, resultValue
    volumeValue = ws.Cells(row, volumeColumn).Value
    timeValue = ws.Cells(row, timeColumn).Value
    
    If IsNumeric(volumeValue) And IsNumeric(timeValue) And timeValue <> 0 Then
        resultValue = volumeValue / timeValue ' Dividing
        ws.Cells(row, resultColumn).Value = resultValue ' Place result
    Else
        ws.Cells(row, resultColumn).Value = "N/A"
    End If
Next


End Sub
