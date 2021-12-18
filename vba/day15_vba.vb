'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

Sub day15a()
On Error GoTo errhand:
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim grid() As Long
rowlength = Len(Cells(1, 1).Value)
ReDim grid(1 To lastrow * rowlength, 1 To rowlength * rowlength)

For i = 1 To lastrow
    For j = 1 To rowlength
        v = Mid(Cells(i, 1).Value, j, 1)
        gridelement = (i - 1) * rowlength + j
        
        If Not (i = 1) Then grid(gridelement - rowlength, gridelement) = v
        If Not (i = lastrow) Then grid(gridelement + rowlength, gridelement) = v
        If Not (j = 1) Then grid(gridelement - 1, gridelement) = v
        If Not (j = rowlength) Then grid(gridelement + 1, gridelement) = v
    Next j
Next i

'For t = 1 To UBound(grid, 1)
'drow = ""
'    For u = 1 To UBound(grid, 2)
'     drow = drow & grid(t, u)
'    Next u
'Debug.Print drow
'Next t

Call dijkstra(grid, 1)





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
Sub createInputForPart2()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

For i = 1 To lastrow
    inp = Cells(i, 1).Value
    newinp = ""
    For e = 1 To 4
        For j = 1 To Len(inp)
            c = Mid(inp, j, 1)
            nextchar = ((c - 1) + e) Mod 9 + 1
            newinp = newinp & nextchar


        Next j
    Next e
    inp = inp & newinp
    Cells(i, 20).Value = inp
Next i

For i = 1 To lastrow
    inp = Cells(i, 20).Value
    
    For e = 1 To 4
    newinp = ""
        For j = 1 To Len(inp)
            c = Mid(inp, j, 1)
            nextchar = ((c - 1) + e) Mod 9 + 1
            newinp = newinp & nextchar
        Next j
    Cells(i + e * lastrow, 20).Value = newinp
    Next e
Next i

End Sub

Sub day15b()

'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim grid() As Byte
rowlength = Len(Cells(1, 1).Value)
ReDim grid(0 To lastrow + 1, 0 To rowlength + 1)

For i = 0 To lastrow + 1
    grid(i, 0) = 15
    grid(i, rowlength + 1) = 15
    For j = 1 To rowlength
        If i = 0 Then
            grid(i, j) = 15
        ElseIf i = rowlength + 1 Then
            grid(i, j) = 15
        Else
            v = Mid(Cells(i, 1).Value, j, 1)
            grid(i, j) = v
        End If
    Next j
Next i

'For t = 1 To UBound(grid, 1)
'drow = ""
'    For u = 1 To UBound(grid, 2)
'     drow = drow & grid(t, u)
'    Next u
'Debug.Print drow
'Next t

Call dijkstra(grid, 1)
'Call homebrew(grid)




'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------


End Sub


Sub dijkstra(grid, src)
On Error GoTo errhand:
Dim dist() As Long
ReDim dist(1 To UBound(grid, 1) * UBound(grid, 2))

Dim shortest() As Boolean
ReDim shortest(1 To UBound(grid, 1) * UBound(grid, 2))

For i = 1 To UBound(dist)
    dist(i) = 999999
Next i

dist(src) = 0

Dim gridpointer As Integer
Dim gridpointery As Integer

'TODO: make this go quicker pls
For i = 1 To UBound(shortest, 1)
    mdist = MinDist(dist, shortest)
    If mdist = 250000 Then GoTo solution
    shortest(mdist) = True
    
    gridpointerx = (mdist - 1) Mod (UBound(grid) - 1) + 1
    gridpointery = Int((mdist - 1) / (UBound(grid) - 1) + 1)
    'For j = 1 To UBound(shortest, 1)
   ' For j = 1 To 4
   '     Select Case j
   '         Case 1 'north
   '             If Not (gridpointerx = 1) Then 'if i< 501
   '                 nb = 1
   '         Case 2 'east
   '             If Not (gridpointery = 500) Then 'if i mod 500 = 0
   '
   '         Case 3 'south
   '             If Not (gridpointerx = 500) Then ' if i >249500
   '
   '         Case 4 'west
   '             If Not (gridpointery = 1) Then ' if (i-1) mod 500 = 0
    
    
    
    
    For j = -2 To 2
        If j = 0 Then '
        Else
            If gridpointerx = 1 And j = -2 Then GoTo nextloop
            If gridpointery = 1 And j = -1 Then GoTo nextloop
            If gridpointery = 500 And j = 1 Then GoTo nextloop
            If gridpointerx = 500 And j = 2 Then GoTo nextloop
        
            neighbourindex = mdist + (j - Sgn(j)) + (j Mod 2) * 500 'i ipv mdist
            
                
            If (shortest(neighbourindex) = False) Then
                If (Not dist(mdist) = 9999999) Then
                    
                    jumptoj = grid(gridpointerx + (j - Sgn(j)), gridpointery + (j Mod 2)) 'edit j values to get all 4 neighbours
                    If (dist(mdist) + jumptoj < dist(neighbourindex)) Then
                        dist(neighbourindex) = dist(mdist) + jumptoj
                    End If
                End If
            End If
            
            'If (shortest(j) = False) Then
            'If (Not grid(mdist, j) = 0) Then ' if j has connection to current point, =next to
            'If (Not dist(mdist) = 9999999) Then 'distance to mdist has not been set yet
            'If (dist(mdist) + grid(mdist, j) < dist(j)) Then  'if distance to mdist and then to j is less then currently in j
            '    dist(j) = dist(mdist) + grid(mdist, j)
            'End If
        End If
nextloop:
    Next j
    
    If i Mod 400 = 0 Then
        Percent = (i * 100) / UBound(dist)
        Debug.Print Percent & " -- " & gridpointerx & " - " & gridpointery
        DoEvents
    End If
Next i

solution:
Debug.Print dist(UBound(dist) - 1)
errhand:


End Sub
Sub homebrew(grid)
On Error GoTo errhands
Dim d_tocheck As Object
Set d_tocheck = CreateObject("Scripting.Dictionary")

Dim dist_from_origin() As Integer
ReDim dist_from_origin(0 To UBound(grid, 1), 0 To UBound(grid, 2))

Dim local_dists As Object
Set local_dists = CreateObject("Scripting.Dictionary")

Dim lowest_found() As Boolean
ReDim lowest_found(1 To UBound(grid, 1) - 1, 1 To UBound(grid, 2) - 1)


For i = LBound(dist_from_origin, 1) To UBound(dist_from_origin, 1)
    For j = LBound(dist_from_origin, 2) To UBound(dist_from_origin, 2)
        dist_from_origin(i, j) = 10999
    Next j
Next i

dist_from_origin(1, 1) = 0
d_tocheck.Add 1 & "-" & 2, 1
d_tocheck.Add 2 & "-" & 1, 1

'local_dists.RemoveAll
For pp = 1 To 250000
For i = 0 To d_tocheck.Count - 1
xy = Split(d_tocheck.keys()(i), "-")
x = xy(0)
y = xy(1)

'find best neighbour
xtop = dist_from_origin(x, y - 1)
xright = dist_from_origin(x + 1, y)
xbot = dist_from_origin(x, y + 1)
xleft = dist_from_origin(x - 1, y)

best = WorksheetFunction.Min(xtop, xleft, xbot, xright)


'add own distance to best neighbour and update dist array
new_dist = best + grid(x, y)
If new_dist < dist_from_origin(x, y) Then
    dist_from_origin(x, y) = new_dist
    local_dists.Add x & "-" & y, dist_from_origin(x, y)
End If



'assuming theres no backtravel OK, otherwise need to keep seperate array, and in each loop set the one who is lowest to true
Next i

'dic_min = Application.Min(local_dists)
dic_min = 9999
For Each Key In d_tocheck.keys
    If local_dists(Key) < dic_min Then
        dic_min = local_dists(Key)
        dic_min_coords = Key
    End If
Next
min_coords = Split(dic_min_coords, "-")
min_coords_x = min_coords(0)
min_coords_y = min_coords(1)

If min_coords_x = 500 And min_coords_y = 500 Then
    Debug.Print dic_min
    Stop
End If

lowest_found(min_coords_x, min_coords_y) = True
d_tocheck.Remove (min_coords_x & "-" & min_coords_y)

'add non occupied neighbours into dict to be checked
If Not d_tocheck.Exists(min_coords_x + 1 & "-" & min_coords_y) Then
    If Not min_coords_x = 500 Then
        If Not lowest_found(min_coords_x + 1, min_coords_y) Then d_tocheck.Add (min_coords_x + 1 & "-" & min_coords_y), 1
    End If
End If
If Not d_tocheck.Exists(min_coords_x - 1 & "-" & min_coords_y) Then
    If Not min_coords_x = 1 Then
        If Not lowest_found(min_coords_x - 1, min_coords_y) Then d_tocheck.Add (min_coords_x - 1 & "-" & min_coords_y), 1
    End If
End If
If Not d_tocheck.Exists(min_coords_x & "-" & min_coords_y + 1) Then
    If Not min_coords_y = 500 Then
        If Not lowest_found(min_coords_x, min_coords_y + 1) Then d_tocheck.Add (min_coords_x & "-" & min_coords_y + 1), 1
    End If
End If
If Not d_tocheck.Exists(min_coords_x & "-" & min_coords_y - 1) Then
    If Not min_coords_y = 1 Then
        If Not lowest_found(min_coords_x, min_coords_y - 1) Then d_tocheck.Add (min_coords_x & "-" & min_coords_y - 1), 1
    End If
End If


    If pp Mod 100 = 0 Then Debug.Print pp / 250000
    DoEvents
Next pp
errhands:
Stop
Resume

End Sub
Function MinDist(dist, shortest)
i_min = 9999999
index_min = -1

For m = 1 To UBound(dist)
    If (shortest(m) = False) Then
        If (dist(m) < i_min) Then
            i_min = dist(m)
            index_min = m
        End If
    End If
Next m
MinDist = index_min
    
End Function
Sub test()
mdist = 5
j = 1


neighbourindex = mdist + (j - Sgn(j)) + (j Mod 2)
nb = 1

End Sub
