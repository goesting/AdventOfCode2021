Sub day5a()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim grid() As Variant
xmax = 999
ymax = 999
ReDim Preserve grid(0 To xmax, 0 To ymax)

For i = 1 To lastrow
    'read data into coords----------------------
    inp_s = Split(Cells(i, 1).Value, " -> ")
    x = Split(inp_s(0), ",")
    y = Split(inp_s(1), ",")
    x1 = CInt(x(0))
    y1 = CInt(x(1))
    x2 = CInt(y(0))
    y2 = CInt(y(1))
    If (Not (x1 = x2 Or y1 = y2)) Then 'diag line
        GoTo notstraigth:
    End If
    
    'add counter to line segment
    If x1 = x2 Then
        If y1 > y2 Then Call swapPos(y1, y2) 'swap pos to not mess up for loop
        For j = y1 To y2
            grid(x1, j) = grid(x1, j) + 1
        Next j
    ElseIf y1 = y2 Then
        If x1 > x2 Then Call swapPos(x1, x2) 'swap pos
        For j = x1 To x2
            grid(j, y1) = grid(j, y1) + 1
        Next j
    End If
notstraigth:
Next i

'count all cells bigger then 1
For a = LBound(grid, 1) To UBound(grid, 1)
    For b = LBound(grid, 2) To UBound(grid, 2)
        If grid(a, b) > 1 Then
            result = result + 1
        End If
    Next b
Next a

Debug.Print result
End Sub

Sub day5b()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim grid() As Variant
xmax = 999
ymax = 999
ReDim Preserve grid(0 To xmax, 0 To ymax)

For i = 1 To lastrow
    'read data into coords----------------------
    inp_s = Split(Cells(i, 1).Value, " -> ")
    x = Split(inp_s(0), ",")
    y = Split(inp_s(1), ",")
    x1 = CInt(x(0))
    y1 = CInt(x(1))
    x2 = CInt(y(0))
    y2 = CInt(y(1))
    If (Not (x1 = x2 Or y1 = y2)) Then 'diag line
        If x2 > x1 Then
            xstep = 1
        Else
            xstep = -1
        End If
        If y2 > y1 Then
            ystep = 1
        Else
            ystep = -1
        End If
        j = y1
        For f = x1 To x2 Step xstep
            grid(f, j) = grid(f, j) + 1
            j = j + ystep
        Next f
    'add counter to line segment
    ElseIf x1 = x2 Then
        If y1 > y2 Then Call swapPos(y1, y2) 'swap pos to not mess up for loop
        For j = y1 To y2
            grid(x1, j) = grid(x1, j) + 1
        Next j
    ElseIf y1 = y2 Then
        If x1 > x2 Then Call swapPos(x1, x2) 'swap pos
        For j = x1 To x2
            grid(j, y1) = grid(j, y1) + 1
        Next j
    End If
Next i

'count all cells bigger then 1
For a = LBound(grid, 1) To UBound(grid, 1)
    For b = LBound(grid, 2) To UBound(grid, 2)
        If grid(a, b) > 1 Then
            result = result + 1
        End If
    Next b
Next a

Debug.Print result
End Sub
Sub swapPos(ByRef number1, ByRef number2)
    temp = number1
    number1 = number2
    number2 = temp
End Sub
