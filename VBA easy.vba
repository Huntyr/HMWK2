Sub doubleloop()

Dim firstofyear As Double
Dim lastofyear As Double
Dim totalvolind As Double
Dim ticker As String
Dim i As Long
Dim lastrow As Long
Dim row As Long

'add labels to row 1'

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

'set variables to account for the length of the range and row counter for later loop

row = 2
lastrow = Range("A1").End(xlDown).row
totalvolind = 0

'loop through the records until the next row is not the same as the current
'when that occurs print variables to given cells
For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        totalvolind = totalvolind + Cells(i, 7)
        Cells(row, 10).Value = totalvolind
        Cells(row, 9).Value = ticker
        
        row = row + 1
        totalvolind = 0
    Else
        totalvolind = totalvolind + Cells(i, 7)
        
    End If
Next i

        
        
        



End Sub
