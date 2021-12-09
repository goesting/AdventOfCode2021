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
x = 1

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
Start = Timer
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
lastcol = Len(Cells(1, 1).Value)
Dim hmap() As Variant
ReDim hmap(0 To lastrow + 1, 0 To lastcol + 1)
Dim chk_q As Object
Dim todo_q As Object
Set chk_q = CreateObject("System.Collections.Queue")
Set todo_q = CreateObject("System.Collections.Queue")
Dim bs() As Variant
ReDim bs(1 To 423)



For i = LBound(hmap, 1) To UBound(hmap, 1)
    For j = LBound(hmap, 2) To UBound(hmap, 2)
        If i = 0 Or i = UBound(hmap, 1) Or j = 0 Or j = UBound(hmap, 2) Then
            hmap(i, j) = 9
        Else
            hmap(i, j) = CInt(Mid(Cells(i, 1).Value, j, 1))
        End If
    
    Next j
Next i
x = 1

For i = LBound(hmap, 1) + 1 To UBound(hmap, 1) - 1
    For j = LBound(hmap, 2) + 1 To UBound(hmap, 2) - 1
        v = hmap(i, j)
        u = hmap(i - 1, j)
        d = hmap(i + 1, j)
        le = hmap(i, j - 1)
        ri = hmap(i, j + 1)
        If (u > v And d > v And le > v And ri > v) Then
            'new low found
            bsize = 0
            todo_q.enqueue i & "," & j
            Do While todo_q.Count > 0
            pop = todo_q.dequeue
            x = Split(pop, ",")(0)
            y = Split(pop, ",")(1)
            'Call getBasinSize(x, y, hmap, chk_q, todo_q, bsize)
            
            bsize = bsize + 1
            chk_q.enqueue x & "," & y
            
            For hor = -1 To 1
                q_s = x + hor & "," & y
                If (Not (chk_q.contains(q_s) Or todo_q.contains(q_s))) And Not (hmap(x + hor, y) = 9) Then
                    todo_q.enqueue q_s
                End If
            Next hor
        
            For ver = -1 To 1
                q_s = x & "," & y + ver
                If (Not (chk_q.contains(q_s) Or todo_q.contains(q_s))) And Not (hmap(x, y + ver) = 9) Then
                    todo_q.enqueue q_s
                End If
            Next ver

            
            
            
            
            
            Loop
            
            'save size somewhere
            bsc = bsc + 1
            bs(bsc) = bsize
            
            
        End If


    Next j



Next i



ActiveSheet.Range("H1").Resize(UBound(bs, 1)).Value = WorksheetFunction.Transpose(bs)
'sort an mulitply in excel worksheet.  works and quicker than writing a sort myself. will do for now



'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub
