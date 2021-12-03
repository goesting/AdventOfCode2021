Sub day2a()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
f = 0
d = 0

For i = 1 To lastrow
inp = Cells(i, 1).Value
inp_s = Split(inp, " ")
c = inp_s(0)
a = inp_s(1)

If c = "forward" Then f = f + a
End If
If c = "down" Then d = d + a
If c = "up" Then d = d - a
Next i
result = f * d
Debug.Print result

End Sub
Sub day2b()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
f = 0
d = 0
aim = 0

For i = 1 To lastrow
inp = Cells(i, 1).Value
inp_s = Split(inp, " ")
c = inp_s(0)
a = inp_s(1)

If c = "forward" Then
f = f + a
d = d + (aim * a)
End If
If c = "down" Then aim = aim + a
If c = "up" Then aim = aim - a

Next i
result = f * d
Debug.Print result

End Sub


