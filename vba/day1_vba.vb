Sub day1a()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To lastrow
inp = Cells(i, 1).Value
If i = 1 Then
'nothing
Else
If inp > prev Then result = result + 1
End If
prev = inp
Next

Debug.Print result

End Sub

Sub day1b()
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
su_old = 99999
su = 99999
For i = 1 To lastrow
inp = Cells(i, 1).Value

If i = 1 Or i = 2 Then
'nothing
Else

su = inp + prev + prev2
If su > su_old Then result = result + 1

End If
prev2 = prev
prev = inp

su_old = su
Next

Debug.Print result

End Sub

