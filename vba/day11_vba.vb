'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long
Sub day11a()
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim grid() As Variant
ReDim grid(1 To 10, 1 To 10)
Dim result As Long
ticks = 100                         'nr of loops to do
'----load data into matrix
For i = 1 To lastrow
    inp = Cells(i, 1).Value
    For j = 1 To Len(inp)
        c = Mid(inp, j, 1)
        grid(i, j) = Int(c)
    Next j
Next i
'----do actual work
For t = 1 To ticks
    Call doTick(grid, result)
Next t
'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
Sub doTick(ByRef grid As Variant, ByRef result)
'----add 1 to all squid-------------
For x = 1 To 10
    For y = 1 To 10
        grid(x, y) = grid(x, y) + 1
    Next y
Next x

'make all 10s flash, and propagate to neighbours
'loop and do again, untill no new flashes are produced
Do
flashfound = False
For x = 1 To 10
    For y = 1 To 10
        v = grid(x, y)
        If v > 9 Then
            flashfound = True
            result = result + 1   'flash counter ++
            
            'update neighbours
            For v = -1 To 1
                For h = -1 To 1
                If x + h > 0 And x + h < 11 And y + v > 0 And y + v < 11 Then   'check for edge cases
                    If Not grid(x + h, y + v) = 0 Then                          'check if neighbour has already flashed, do not update those
                        grid(x + h, y + v) = grid(x + h, y + v) + 1
                    End If
                End If
                Next h
            Next v
            grid(x, y) = 0                                                      'set flasher to 0
        End If
    Next y
Next x
Loop While flashfound

End Sub

Sub day11b()
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
Dim grid() As Variant
ReDim grid(1 To 10, 1 To 10)
Dim result As Long
tick = 0                'nr of loops done
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

'load data into matrix
For i = 1 To lastrow
    inp = Cells(i, 1).Value
    For j = 1 To Len(inp)
        c = Mid(inp, j, 1)
        grid(i, j) = Int(c)
    Next j
Next i

'loop until everything flashes at once
Do While True
    tick = tick + 1     'loop counter++
    allflash = doTick_part2(grid) 'return == true if everything flashes
    
    If allflash Then
        GoTo finish 'exit infinite loop
    End If
Loop
finish:
result = tick
'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
Function doTick_part2(ByRef grid As Variant)
'increment all elements by 1
For x = 1 To 10
    For y = 1 To 10
        grid(x, y) = grid(x, y) + 1
    Next y
Next x

'make all 10s flash, and propagate to neighbours
'loop and do again, untill no new flashes are produced
Do
flashfound = False
For x = 1 To 10
    For y = 1 To 10
        v = grid(x, y)
        If v > 9 Then
            flashfound = True
            local_result = local_result + 1  'number of flashes in this tick
            
            'update neighbours of flasher
            For v = -1 To 1
                For h = -1 To 1
                If x + h > 0 And x + h < 11 And y + v > 0 And y + v < 11 Then   'check edge cases
                    If Not grid(x + h, y + v) = 0 Then                          'check if already flashes
                        grid(x + h, y + v) = grid(x + h, y + v) + 1
                    End If
                End If
                Next h
            Next v
            grid(x, y) = 0              'set flasher to 0
        End If
    Next y
Next x
Loop While flashfound

If local_result = 100 Then  'if everything flashes set return value to true
    doTick_part2 = True
Else
    doTick_part2 = False
End If

End Function

