Sub day3a()
Dim secondsSince As Single
secondsSince = Timer()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim arr() As Variant
ReDim arr(1 To lastrow)
'Dim d As Object
'Set d = CreateObject("Scripting.Dictionary")
'inp = WorksheetFunction.Transpose(Range("A1:A1097"))

Dim coun() As Variant
ReDim coun(1 To 12)

For i = 1 To lastrow
inp = Cells(i, 1).Value
arr(i) = inp
Next i

For i = LBound(arr) To UBound(arr)
For j = 1 To Len(arr(i))
c = Mid(arr(i), j, 1)
If c = "1" Then
    coun(j) = coun(j) + 1
Else
    coun(j) = coun(j) - 1
End If




Next j
Next i


Gamma = 0
epsilon = 0

For k = 1 To 12
most_nr = coun(13 - k)
If most_nr > 0 Then
most_nr = 1
Gamma = Gamma + WorksheetFunction.Power(2, (k - 1))
Else
most_nr = 0
epsilon = epsilon + WorksheetFunction.Power(2, (k - 1))
End If




Next k
result = epsilon * Gamma
Debug.Print "solution for part a: " & result
Debug.Print "time for part a: " & Timer - secondsSince


End Sub
Sub day3b()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim common() As Variant
Dim secondsSince As Single
ReDim common(1 To 12)
Dim matchdict As Object
Set matchdict = CreateObject("Scripting.Dictionary")
Dim missdict As Object
Set missdict = CreateObject("Scripting.Dictionary")
secondsSince = Timer()

For i = 1 To lastrow
    inp = ActiveSheet.Cells(i, 1).Value
    matchdict.Add i, inp
    missdict.Add i, inp
Next i

For k = 1 To 12
    comm = 0
    For Each key In matchdict.Keys
        inp = matchdict(key)
        lett = Mid(inp, k, 1)
        If lett = "1" Then
            comm = comm + 1
        Else
            comm = comm - 1
        End If
    Next
    
    If comm >= 0 Then
        matchme = "1"
    Else
        matchme = "0"
    End If
    
    For Each key In matchdict.Keys
        inp = matchdict(key)
        lett = Mid(inp, k, 1)
        If lett <> matchme Then
            matchdict.Remove key
        End If
    Next
    Dim match_solution As String
    If matchdict.Count = 1 Then
        For Each key In matchdict.Keys
            match_solution = matchdict(key)
        Next
        Exit For
    End If

Next k
match_sol_dec = Bintodec(match_solution)
x = 1

' --- hieronder ne perfecte repeat van de code hierboven, maar dan voor least common. kan wrs efficienter :shrug: ----------

For k = 1 To 12
    comm = 0
    For Each key In missdict.Keys
        inp = missdict(key)
        lett = Mid(inp, k, 1)
        If lett = "1" Then
            comm = comm + 1
        Else
            comm = comm - 1
        End If
    Next
    
    If comm >= 0 Then
        matchme = "0"
    Else
        matchme = "1"
    End If
    
    For Each key In missdict.Keys
        inp = missdict(key)
        lett = Mid(inp, k, 1)
        If lett <> matchme Then
            missdict.Remove key
        End If
    Next
    Dim miss_solution As String
    If missdict.Count = 1 Then
        For Each key In missdict.Keys
            miss_solution = missdict(key)
        Next
        Exit For
    End If

Next k
miss_sol_Dec = Bintodec(miss_solution)



result = miss_sol_Dec * match_sol_dec

Debug.Print "solution for part b: " & result
Debug.Print "time for part b: " & Timer - secondsSince





End Sub
