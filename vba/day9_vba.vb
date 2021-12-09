Sub day8a()
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
lastcol = Len(Cells(1, 1).Value)
Dim hmap() As Variant
ReDim hmap(0 To lastrow + 1, 0 To lastcol + 1)

For i = LBound(hmap, 1) To UBound(hmap, 1)
    For j = LBound(hmap, 2) To UBound(hmap, 2)
        If i = 0 Or i = UBound(hmap, 1) Or j = 0 Or j = UBound(hmap, 2) Then
            hmap(i, j) = 9
        Else
            hmap(i, j) = CInt(Mid(Cells(i, 1).Value, j, 1))
        End If
    
    Next j
Next i

For i = LBound(hmap, 1) + 1 To UBound(hmap, 1) - 1
    For j = LBound(hmap, 2) + 1 To UBound(hmap, 2) - 1
        v = hmap(i, j)
        u = hmap(i - 1, j)
        d = hmap(i + 1, j)
        le = hmap(i, j - 1)
        ri = hmap(i, j + 1)
        If (u > v And d > v And le > v And ri > v) Then
            risk = v + 1
            totalrisk = totalrisk + risk
        End If
    Next j
Next i

result = totalrisk
'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub

Sub day8b()
Dim bs() As Variant                                                             'array to fill with all the bassin sizes
Dim hmap() As Variant                                                           'input, height map
Dim chk_q As Object                                                             'queue of all the points to check if they are in basin
Dim todo_q As Object                                                            'queue of all points that have been checked
Dim points_around() As Variant

Set chk_q = CreateObject("System.Collections.Queue")
Set todo_q = CreateObject("System.Collections.Queue")

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
lastcol = Len(Cells(1, 1).Value)

ReDim hmap(0 To lastrow + 1, 0 To lastcol + 1)                                  'add 0 and +1 and fill with all 9s
ReDim bs(1 To 423)                                                              'number of basin low points, from part 1

Start = Timer

For i = LBound(hmap, 1) To UBound(hmap, 1)                                      'traverse height map and fill it with input data
    For j = LBound(hmap, 2) To UBound(hmap, 2)
        If i = 0 Or i = UBound(hmap, 1) Or j = 0 Or j = UBound(hmap, 2) Then    'if edges fill with 9s to bound in box
            hmap(i, j) = 9
        Else
            hmap(i, j) = CInt(Mid(Cells(i, 1).Value, j, 1))
        End If
    Next j
Next i

For i = LBound(hmap, 1) + 1 To UBound(hmap, 1) - 1                              'Locate low points
    For j = LBound(hmap, 2) + 1 To UBound(hmap, 2) - 1
        v = hmap(i, j)                                                          'value of point
        u = hmap(i - 1, j)                                                      'value point up, down, left, and right
        d = hmap(i + 1, j)
        le = hmap(i, j - 1)
        ri = hmap(i, j + 1)
        If (u > v And d > v And le > v And ri > v) Then                         'if new low found -> start calculating the basin size
            bsize = 0                                                           'reset size of current basin
            todo_q.enqueue i & "," & j                                          'place point in the todo queue
            Do While todo_q.Count > 0                                           'if queue is not empty,
                pop = todo_q.dequeue                                            'get item and check
                x = Split(pop, ",")(0)
                y = Split(pop, ",")(1)
                bsize = bsize + 1                                               'increase basin size for every item gotten
                chk_q.enqueue x & "," & y                                       'place item in the queue of items done (for checking purpose)
                
                ReDim points_around(1 To 4)                                     'reset points around and fill
                points_around(1) = Array(x - 1, y)
                points_around(2) = Array(x + 1, y)
                points_around(3) = Array(x, y - 1)
                points_around(4) = Array(x, y + 1)
                
                For k = LBound(points_around) To UBound(points_around)          'more performance maybe? 2.5s->1.5s
                    q_s = points_around(k)(0) & "," & points_around(k)(1)       'nested ifs instead of 1 big if, much quicker
                    If Not todo_q.contains(q_s) Then                            'if point is not in checked queue or todo queue, and it
                        If Not chk_q.contains(q_s) Then                         'is not 9, then put it in todo queue
                            If Not hmap(points_around(k)(0), points_around(k)(1)) = 9 Then
                                todo_q.enqueue q_s
                            End If
                        End If
                    End If
                Next k
            Loop 'until todo queue is empty
            bsc = bsc + 1                                                       'index for array holding all the basin sizes
            bs(bsc) = bsize                                                     'add total size to the array
        End If
    Next j
Next i
ActiveSheet.Range("H1").Resize(UBound(bs, 1)).Value = WorksheetFunction.Transpose(bs) 'put results into Range
Dim rngTOSort As Range
Set rngTOSort = ActiveWorkbook.Sheets("day9").Range("H1:H423")
With rngTOSort                                                                          'sort the Range
    .Sort Key1:=Sheets("day9").Range("H1"), Order1:=xlDescending, Header:=xlNo, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
End With
result = [H1] * [H2] * [H3]                                                             'Multiply top 3
'------ Output results and time taken ---------------------
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
