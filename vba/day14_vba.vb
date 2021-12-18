'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

Sub day14a()
On Error GoTo errhand:
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
Dim charoccuring As Object
Set charoccuring = CreateObject("Scripting.Dictionary")


Dim nextstr() As Variant

lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row

inp = Cells(1, 1).Value
ReDim nextstr(1 To Len(inp) - 1)
'Dim insert_commands() As Variant
'ReDim insert_commands(1 To lastrow - 2, 1 To 2)
'read input data
For i = 3 To lastrow
    v = Split(Cells(i, 1).Value, " -> ")
    'insert_commands(i - 2, 1) = v(0)
    'insert_commands(i - 2, 1) = v(1)
    dict.Add v(0), v(1)
Next i


steps = 10
For i = 1 To steps
'Debug.Print inp
    For j = 1 To Len(inp) - 1
        firstchar = Left(inp, 1)
        substr2 = Mid(inp, j, 2)
        If dict.Exists(substr2) Then
            nextstr(j) = dict(substr2) & Right(substr2, 1)
        Else
            nextstr(j) = Right(substr2, 1)
        End If
    Next j
    
    inp = firstchar & Join(nextstr, "")
    ReDim nextstr(1 To Len(inp) - 1)
Next i


For i = 1 To Len(inp)
    c = Mid(inp, i, 1)
    If charoccuring.Exists(c) Then
        charoccuring(c) = charoccuring(c) + 1
    Else
        charoccuring.Add c, 1
    End If
Next i

Max = Application.Max(charoccuring.Items)
Min = Application.Min(charoccuring.Items)
result = Max - Min





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
Sub day14b() 'brute force will not work
On Error GoTo errhand:
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
Dim charoccuring As Object
Set charoccuring = CreateObject("Scripting.Dictionary")
Dim charoccuring_tmp As Object
Set charoccuring_tmp = CreateObject("Scripting.Dictionary")
Dim d_result As Object
Set d_result = CreateObject("Scripting.Dictionary")
'init constant
inp = Cells(1, 1).Value
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
steps = 40
firstchar = Left(inp, 1)
lastchar = Right(inp, 1)


'load input
For i = 3 To lastrow
    v = Split(Cells(i, 1).Value, " -> ")
    Item1 = v(0)
    item2 = Left(Item1, 1) & v(1) & Right(Item1, 1)
    dict.Add Item1, item2
    charoccuring.Add Item1, 0
Next i

For i = 1 To Len(inp) - 1
    string2 = Mid(inp, i, 2)
    charoccuring(string2) = charoccuring(string2) + 1
Next i

'Dim temp() As Variant
'ReDim temp(0 To charoccuring.Count - 1)
Dim outp As String, outp1 As String, outp2 As String

'do actual code
For Step = 1 To steps

For i = 0 To charoccuring.Count - 1
    v = charoccuring.Items()(i)
    k = charoccuring.keys()(i)
    outp = dict(k)
    outp1 = Left(outp, 2)
    outp2 = Right(outp, 2)
    
    If charoccuring_tmp.Exists(outp1) Then
        charoccuring_tmp(outp1) = charoccuring_tmp(outp1) + v
    Else
        charoccuring_tmp.Add outp1, v
    End If
    If charoccuring_tmp.Exists(outp2) Then
        charoccuring_tmp(outp2) = charoccuring_tmp(outp2) + v
    Else
        charoccuring_tmp.Add outp2, v
    End If
    'temp(i) = temp(outp1) + k
    'temp(outp2) = temp(outp2) + k
Next i

'For i = LBound(temp) To UBound(temp)
'    k = charoccuring.Keys()(i)
'    charoccuring(k) = temp(i)
'Next i

'test dic
'For i = 0 To dict.Count - 1
'Debug.Print dict.Keys()(i), dict.Items()(i)
'Next i


'load tmp dict data into actual dict
charoccuring.RemoveAll
For i = 0 To charoccuring_tmp.Count - 1
charoccuring.Add charoccuring_tmp.keys()(i), charoccuring_tmp.Items()(i)
Next i
charoccuring_tmp.RemoveAll



'For i = 0 To charoccuring.Count - 1
'Debug.Print charoccuring.Keys()(i), charoccuring.Items()(i)
'Next i
'Debug.Print "------------------------------------"

Next Step

'getresult from dict

For i = 0 To charoccuring.Count - 1
letter = Left(charoccuring.keys()(i), 1)
v = charoccuring.Items()(i)

If d_result.Exists(letter) Then
    d_result(letter) = d_result(letter) + v
Else
    d_result.Add letter, v
End If
Next i

d_result(lastchar) = d_result(lastchar) + 1 ' account for the last char not being in the loop



'For i = 0 To d_result.Count - 1
'Debug.Print d_result.Keys()(i), d_result.Items()(i)
'Next i

Max = Application.Max(d_result.Items)
Min = Application.Min(d_result.Items)
result = Max - Min

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
