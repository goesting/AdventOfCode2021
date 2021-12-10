Sub day8a()
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

valido = Array("(", "[", "{", "<")                  'VALID Open characters
validc = Array(")", "]", "}", ">")                  'VALID Close character
punten = Array(3, 57, 1197, 25137)                  'punten per symbool

Dim stack As Object
Set stack = CreateObject("System.Collections.Stack") 'Create Stack

For i = 1 To lastrow
inp = Cells(i, 1).Value
    For j = 1 To Len(inp)
        c = Mid(inp, j, 1)
        If UBound(Filter(valido, c)) > -1 Then      '(if c is in array valido) c==open bracket
                stack.push c                        'put on stack
        Else                                        'c==closebracket
            d = stack.pop                           'pop last from stack and check if this closes
            With Application
                If Not (.Match(d, valido, False) = .Match(c, validc, False)) Then
                    result = result + punten(.Match(c, validc, False) - 1) '-1 to get correct 0 starting array index
                End If
            End With
        End If
    Next j
Next i
'------ Output results and time taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
Sub day8b()
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
valido = Array("(", "[", "{", "<")                  'VALID Open characters
validc = Array(")", "]", "}", ">")                  'VALID Close character
punten = Array(1, 2, 3, 4)                          'punten per symbool

Dim results() As Variant
ReDim results(1 To lastrow)

Dim stack As Object
Set stack = CreateObject("System.Collections.Stack") 'Create Stack

For i = 1 To lastrow
inp = Cells(i, 1).Value
    For j = 1 To Len(inp)
        c = Mid(inp, j, 1)
        If UBound(Filter(valido, c)) > -1 Then  'open bracket
                stack.push c
        Else                                    'closebracket
            d = stack.pop
            With Application
                If Not (.Match(d, valido, False) = .Match(c, validc, False)) Then
                    stack.Clear                 'ignore invalid lines, clear stack and do next
                    GoTo nextline
                End If
            End With
        End If
    Next j
    
    line_result = 0
    Do While stack.Count > 0
        c_score = punten(Application.Match(stack.pop, valido, False) - 1)
        line_result = (line_result * 5) + c_score
    Loop
    results(i) = line_result
nextline:
Next i

'-----write results and sort them, then get middle element -------------------------------------
ActiveSheet.Range("H1").Resize(UBound(results, 1)).Value = WorksheetFunction.Transpose(results) 'put results into Range
Dim rngTOSort As Range
Set rngTOSort = ActiveWorkbook.Sheets("day10").Range("H1:H" & lastrow)
With rngTOSort                                                                                  'sort the Range
    .Sort Key1:=Sheets("day10").Range("H1"), Order1:=xlDescending, Header:=xlNo, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
End With
lastHrow = ActiveSheet.Range("H" & Rows.Count).End(xlUp).Row                                    'find amount of output results
result = Cells(Int((lastHrow + 1) / 2), 8).Value                                                'get value in halfway down

'------ Output results and time taken ---------------------
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
