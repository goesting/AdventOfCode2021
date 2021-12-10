'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

Sub main()
'alt methods are a lot quicker
Call day10a
Call day10a_alt
Call day10b
Call day10b_alt
End Sub

Sub day10a()
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
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
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
Sub day10b()
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
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
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
Sub day10a_alt()
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim results() As Variant
ReDim results(1 To lastrow)

For i = 1 To lastrow
    inp = Cells(i, 1).Value
    l = Len(inp)
    new_l = 0
    Do While l <> new_l
        l = new_l
        inp = Replace(Replace(Replace(Replace(inp, "()", ""), "[]", ""), "{}", ""), "<>", "")
        new_l = Len(inp)
    Loop
    For j = 1 To l
        Select Case Mid(inp, j, 1)
            Case ")"
                result = result + 3
                GoTo nextline
            Case "]"
                result = result + 57
                GoTo nextline
            Case "}"
                result = result + 1197
                GoTo nextline
            Case ">"
                result = result + 25137
                GoTo nextline
            Case Else: 'do nothing
        End Select
    Next j
nextline:
Next i

'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken (alt method) = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub

Sub day10b_alt()
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Counter = 1 'index for saving the results to array results

Dim results() As Variant
ReDim results(1 To 51) '51 is the remaining lines that are not invalid, from part 1

For i = 1 To lastrow
    inp = Cells(i, 1).Value
    l = Len(inp)
    new_l = 0
    Do While l <> new_l
        l = new_l
        inp = Replace(Replace(Replace(Replace(inp, "()", ""), "[]", ""), "{}", ""), "<>", "")
        new_l = Len(inp)
    Loop
    result = 0
    For j = l To 1 Step -1
        result = result * 5
        Select Case Mid(inp, j, 1)
            Case "(": result = result + 1
            Case "[": result = result + 2
            Case "{": result = result + 3
            Case "<": result = result + 4
            Case Else
                result = 0
                GoTo nextline
        End Select
    Next j
    results(Counter) = result
    Counter = Counter + 1
nextline:
Next i
Call QuickSort(results, 1, UBound(results))
result = results(26)
'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken (alt method) = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
Dim pivot   As Variant
Dim tmpSwap As Variant
Dim tmpLow  As Long
Dim tmpHi   As Long

tmpLow = inLow
tmpHi = inHi

pivot = vArray((inLow + inHi) \ 2)

While (tmpLow <= tmpHi)
   While (vArray(tmpLow) < pivot And tmpLow < inHi)
      tmpLow = tmpLow + 1
   Wend

   While (pivot < vArray(tmpHi) And tmpHi > inLow)
      tmpHi = tmpHi - 1
   Wend

   If (tmpLow <= tmpHi) Then
      tmpSwap = vArray(tmpLow)
      vArray(tmpLow) = vArray(tmpHi)
      vArray(tmpHi) = tmpSwap
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
   End If
Wend

If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
