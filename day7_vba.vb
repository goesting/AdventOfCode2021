'Needs cleanup
Sub day5()
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
bestresult = 99999999

inp = Split(Cells(1, 1).Value, ",")
With WorksheetFunction
low = 0
high = 2000
End With

For i = low To high
result = 0
For j = LBound(inp) To UBound(inp)
'part 1 implemnetation -------------
'    m = Math.Abs(inp(j) - i)     '-
'    result = result + (m)        '-
'-----------------------------------

'part 2 implemnetation -------------
    n = Math.Abs(inp(j) - i)      '-
    m = n * ((n + 1) / 2)         '-
    result = result + (m)         '-
'-----------------------------------

Next j
If result < bestresult Then bestresult = result

Next i

result = bestresult
'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
