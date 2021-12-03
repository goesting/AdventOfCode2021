Sub day4b()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
Dim common() As Variant
ReDim common(1 To 12)
Dim matchdict As Object
Set matchdict = CreateObject("Scripting.Dictionary")
Dim missdict As Object
Set missdict = CreateObject("Scripting.Dictionary")
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

Debug.Print result

End Sub

Function Bintodec(bin As String) As LongLong
    For i = 1 To Len(bin)
        Bintodec = Bintodec + WorksheetFunction.Power(2, Len(bin) - i) * Mid(bin, i, 1)
    Next i

End Function
