'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

Sub day13a()
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
Dim dots() As Variant
Dim fold_ins() As Variant
Dim grid() As Variant
'find size of dot list and fold instructions
For i = 1 To lastrow
If Cells(i, 1).Value = "" Then
    dots_size = i - 1
    foldins_size = lastrow - i
    Exit For
End If
Next i

ReDim fold_ins(1 To foldins_size)
ReDim dots(1 To dots_size, 1 To 2)

'read data into arrays
foldinstr = False
For i = 1 To lastrow
inp = Cells(i, 1).Value
If inp = "" Then
    foldinstr = True
    GoTo nextloop
End If
If foldinstr Then
    inp_s = Split(inp, " ")
    fold_ins(i - dots_size - 1) = inp_s(2)
    
Else
    inp_s = Split(inp, ",")
    dots(i, 1) = inp_s(0)
    dots(i, 2) = inp_s(1)

End If
nextloop:
Next i
x = 1
'actusl code

'size of paper after 1 fold = 0 to 654    0 to 893
ReDim grid(0 To 654, 0 To 893)
xsize = 1310
ysize = 894

For i = LBound(dots, 1) To UBound(dots, 1)
    xsize = 1310
    ysize = 894
    x = CInt(dots(i, 1))
    y = CInt(dots(i, 2))
    x_new = x
    y_new = y
    
    
    
    Do While x > 654
        x_new = xsize - x
        xsize = xsize / 2 - 1
        x = x_new
    Loop
    Do While y > 893
        y_new = ysize - y
        ysize = ysize / 2 - 1
        y = y_new
    
    Loop

grid(x_new, y_new) = 1
Next i

For i = LBound(grid, 1) To UBound(grid, 1)
    For j = LBound(grid, 2) To UBound(grid, 2)
        If grid(i, j) = 1 Then result = result + 1
    Next j
Next i

Debug.Print result

errhand:
Stop
Resume
End Sub

Sub day13b()
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
Dim dots() As Variant
Dim fold_ins() As Variant
Dim grid() As Variant
'find size of dot list and fold instructions
For i = 1 To lastrow
If Cells(i, 1).Value = "" Then
    dots_size = i - 1
    foldins_size = lastrow - i
    Exit For
End If
Next i

ReDim fold_ins(1 To foldins_size)
ReDim dots(1 To dots_size, 1 To 2)

'read data into arrays
foldinstr = False
For i = 1 To lastrow
inp = Cells(i, 1).Value
If inp = "" Then
    foldinstr = True
    GoTo nextloop
End If
If foldinstr Then
    inp_s = Split(inp, " ")
    fold_ins(i - dots_size - 1) = inp_s(2)
    
Else
    inp_s = Split(inp, ",")
    dots(i, 1) = inp_s(0)
    dots(i, 2) = inp_s(1)

End If
nextloop:
Next i
x = 1
'actusl code

'size of paper after 1 fold = 0 to 654    0 to 893
ReDim grid(0 To 654, 0 To 893)
xsize = 1310
ysize = 894

For i = LBound(dots, 1) To UBound(dots, 1)
    xsize = 1310
    ysize = 894
    x = CInt(dots(i, 1))
    y = CInt(dots(i, 2))
    x_new = x
    y_new = y
    
    
    
    Do While x > 39
        If x > xsize / 2 Then
            x_new = xsize - x
        End If
        xsize = xsize / 2 - 1
        x = x_new
    Loop
    Do While y > 5
        If y > ysize / 2 Then
            y_new = ysize - y
        End If
        ysize = ysize / 2 - 1
        y = y_new
    
    Loop

grid(x_new, y_new) = 1
Next i

For i = LBound(grid, 1) To UBound(grid, 1)
    For j = LBound(grid, 2) To UBound(grid, 2)
        Cells(j + 1, i + 4).Value = grid(i, j)
    Next j
Next i

errhand:
Stop
Resume
End Sub

