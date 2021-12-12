'ROUGH DRAFT, NEEDS CLEANUP AND OPTIMALIZATION
'0.1 bit of cleanup done, but till slow

'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

Sub day12a()
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
ReDim grid(1 To lastrow, 0 To 1)
Dim result As Long
'----load data into matrix
For i = 1 To lastrow
    inp = Split(Cells(i, 1).Value, "-")
    from = inp(0)
    toward = inp(1)
    grid(i, 0) = from
    grid(i, 1) = toward
Next i
currenpath = "start"
Path = getNextNode(grid, currenpath, result)

'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
Function getNextNode(cavedata, currentpath, ByRef result, Optional part2 = False, Optional ByVal twicevisitedsmall = False)
currentnode = Split(currentpath, "-")(UBound(Split(currentpath, "-")))
If currentnode = "end" Then
    'path done
    getNextNode = currenpath
    DoEvents 'prevent excel from being unresponsive
    result = result + 1
Else
'For all possible next nodes
    For i = LBound(cavedata, 1) To UBound(cavedata, 1)          'check all possible paths for ones matching thisnode
        For j = 0 To 1                                          'nodes can go left->right as right -> left, so check both
            If (cavedata(i, j) = currentnode) Then              'current node found in data
                NextNode = cavedata(i, 1 - j)                   'nextnode is the other one in the list
                If NextNode = "start" Then                      'if it connects to start, ignore
                ElseIf Not NextNode = UCase(NextNode) Then      'if lowercase, do checks for times visited etc
                    If part2 And Not twicevisitedsmall Then                                                                         'no small cave has been visited twice
                        If Len(currentpath) < Len(Replace(currentpath, NextNode, "")) + Len(NextNode) * 2 Then                      'if visited small cave 0 or 1 times
                            If Len(currentpath) > Len(Replace(currentpath, NextNode, "")) Then                                      'if its 1 times, do the double visit and set to true
                                getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, True)
                            Else                                                                                                    'else its an unvisited node, so go for it
                                getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, twicevisitedsmall)
                            End If
                        End If
                    ElseIf Len(currentpath) = Len(Replace(currentpath, NextNode, "")) Then                                          'either part 1, or part 2 but we have already visited a small thing twice
                        getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, twicevisitedsmall)         'do next only if small cave wasnt visited before, by checking if it exists in currentpath
                    End If
                Else                                                                                                                'if uppercase, just do next, always allowed
                    getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, twicevisitedsmall)
                End If
            End If
        Next j
    Next i
End If

End Function

Sub day12b()
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
ReDim grid(1 To lastrow, 0 To 1)
Dim result As Long
'----load data into matrix
For i = 1 To lastrow
    inp = Split(Cells(i, 1).Value, "-")
    from = inp(0)
    toward = inp(1)
    grid(i, 0) = from
    grid(i, 1) = toward
Next i
currenpath = "start"
Path = getNextNode(grid, currenpath, result, True)
'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
End Sub
