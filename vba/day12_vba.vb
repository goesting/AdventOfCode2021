'ROUGH DRAFT, NEEDS CLEANUP AND OPTIMALIZATION


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
On Error GoTo errhand
'If part2 Then twicevisitedsmall = False


currentnode = Split(currentpath, "-")(UBound(Split(currentpath, "-")))
If currentnode = "end" Then
    'path done
    getNextNode = currenpath
    'Debug.Print currentpath
    DoEvents
    result = result + 1
Else
'For all possible next nodes
For i = LBound(cavedata, 1) To UBound(cavedata, 1)
    For j = 0 To 1
    If (cavedata(i, j) = currentnode) Then
        NextNode = cavedata(i, 1 - j)
        If NextNode = "start" Then
        ElseIf Not NextNode = UCase(NextNode) Then
        
            If part2 And Not twicevisitedsmall Then
                If Len(currentpath) < Len(Replace(currentpath, NextNode, "")) + Len(NextNode) * 2 Then
                    If Len(currentpath) > Len(Replace(currentpath, NextNode, "")) Then
                    getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, True)
                    Else
                    getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, twicevisitedsmall)
                    End If
                End If
                
        
            ElseIf Len(currentpath) = Len(Replace(currentpath, NextNode, "")) Then
            
                getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, twicevisitedsmall)
            End If
        
        Else
        
        getNextNode = getNextNode(cavedata, currentpath & "-" & NextNode, result, part2, twicevisitedsmall)
        End If
    End If
    Next j
Next i
'next node
End If

GoTo contin
errhand:
Stop
Resume
contin:

End Function



Sub day12b()
On Error GoTo errhand
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
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
errhand:
Stop
Resume
End Sub
