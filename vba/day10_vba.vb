Sub day8a()
On Error GoTo errorhands
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim valido(4) As String
Dim validc(4) As String
Dim c As String
valido(1) = "["
valido(2) = "{"
valido(3) = "("
valido(4) = "<"
validc(1) = "]"
validc(2) = "}"
validc(3) = ")"
validc(4) = ">"
Dim punten(4) As Integer
punten(1) = 57
punten(2) = 1197
punten(3) = 3
punten(4) = 25137

Dim stack As Object
Set stack = CreateObject("System.Collections.Stack") 'Create Stack

For i = 1 To lastrow
inp = Cells(i, 1).Value
    For j = 1 To Len(inp)
        c = Mid(inp, j, 1)
        If IsInArray(c, valido) Then 'open bracket
                stack.push c
        
        Else 'closebracket
            d = stack.pop
            With Application
            If Not (.Match(d, valido, False) = .Match(c, validc, False)) Then
            result = result + punten(.Match(c, validc, False) - 1)
            End If
            End With
            
            
'): 3 points.
']: 57 points.
'}: 1197 points.
'>: 25137 points.
            
            
        End If

    Next j
Next i
'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
errorhands:
Stop
Resume
End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub day8b()
On Error GoTo errorhands
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim valido(4) As String
Dim validc(4) As String
    Dim results() As Variant
    ReDim results(1 To lastrow)
Dim c As String
valido(1) = "["
valido(2) = "{"
valido(3) = "("
valido(4) = "<"
validc(1) = "]"
validc(2) = "}"
validc(3) = ")"
validc(4) = ">"
Dim punten(4) As Integer
punten(1) = 2
punten(2) = 3
punten(3) = 1
punten(4) = 4

Dim stack As Object
Set stack = CreateObject("System.Collections.Stack") 'Create Stack

For i = 1 To lastrow
inp = Cells(i, 1).Value
    For j = 1 To Len(inp)
        c = Mid(inp, j, 1)
        If IsInArray(c, valido) Then 'open bracket
                stack.push c
        
        Else 'closebracket
            d = stack.pop
            With Application
            If Not (.Match(d, valido, False) = .Match(c, validc, False)) Then
            stack.Clear
            GoTo nextline
            End If
            End With
            
            

            
            
        End If

    Next j
    line_result = 0
    Do While stack.Count > 0
        pop = stack.pop
        ind = Application.Match(pop, valido, False) - 1
        c_score = punten(ind)
        line_result = (line_result * 5) + c_score
    
    Loop

    results(i) = line_result
    'result = result + line_result
    
nextline:
DoEvents
If i Mod 10 = 0 Then
Debug.Print "On line " & i
End If
Next i

ActiveSheet.Range("H1").Resize(UBound(results, 1)).Value = WorksheetFunction.Transpose(results) 'put results into Range
Dim rngTOSort As Range
Set rngTOSort = ActiveWorkbook.Sheets("day10").Range("H1:H" & lastrow)
With rngTOSort                                                                          'sort the Range
    .Sort Key1:=Sheets("day10").Range("H1"), Order1:=xlDescending, Header:=xlNo, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
End With
lastHrow = ActiveSheet.Range("H" & Rows.Count).End(xlUp).Row
result = Cells(Int((lastHrow + 1) / 2), 8).Value

'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
errorhands:
Stop
Resume
End Sub
