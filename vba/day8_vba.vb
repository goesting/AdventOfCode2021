Sub day8a()
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To lastrow
    inp = Split(Split(Cells(i, 1).Value, "| ")(1), " ")
    For j = LBound(inp) To UBound(inp)
        l = Len(inp(j))
        If l = 2 Or l = 4 Or l = 3 Or l = 7 Then
            result = result + 1
        End If
    Next j
Next i
'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
Sub day8b()
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To lastrow
    s = Split(Cells(i, 1).Value, "| ")(0)
    todecode = Split(Split(Cells(i, 1).Value, " | ")(1), " ")
    For j = LBound(todecode) To UBound(todecode)
        st = todecode(j)
        som = 0
        For c = 1 To Len(st)
            som = som + Len(s) - Len(Replace(s, Mid(st, c, 1), ""))
        Next c
        Select Case som                                'do magic :3
            Case 42: T = 0
            Case 17: T = 1
            Case 34: T = 2
            Case 39: T = 3
            Case 30: T = 4
            Case 37: T = 5
            Case 41: T = 6
            Case 25: T = 7
            Case 49: T = 8
            Case 45: T = 9
        End Select
        result = result * 10 + T
    Next j
    bigresult = bigresult + result
    result = 0
Next i
'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 2: " & bigresult
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
