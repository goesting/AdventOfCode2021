
'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long
Sub day18a()
On Error GoTo errhand
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
snail = Cells(1, 1).Value

For i = 2 To lastrow
snail = "[" & snail & "," & Cells(i, 1).Value & "]"


Do 'nothing change
restart_all_check:
    hassplit = False
    Do  'nothing explodes
        'explode
restart_expl_check:
        exploded = False
        inbracks = 0
        nr1 = 0
        nr2 = 0
        nextnr = 0
        prevnr = 0
        preinbetweenstr = ""
        postinbetweenstr = ""
        For j = 1 To Len(snail)
            c = Mid(snail, j, 1)
            
            If inbracks = 5 Then 'go boom
            exploded = True
                'get two numbers
                fctr = j
                f = Mid(snail, fctr, 1)
                Do While Not f = ","
                    nr1 = nr1 * 10 + f
                    fctr = fctr + 1
                    f = Mid(snail, fctr, 1)
                Loop
                fctr = fctr + 1
                f = Mid(snail, fctr, 1)
                Do While Not f = "]"
                    nr2 = nr2 * 10 + f
                    fctr = fctr + 1
                    f = Mid(snail, fctr, 1)
                Loop
            
                'get previous number
                o = 0
                For k = j - 2 To 1 Step -1
                    d = Mid(snail, k, 1)
                   
                    If IsNumeric(d) Then
                        prevnr = WorksheetFunction.Power(10, o) * d + prevnr
                        o = o + 1
                    ElseIf o > 0 Then
                        Exit For
                    Else
                        preinbetweenstr = d & preinbetweenstr
                    End If
                Next k
                If k = 0 Then
                prevnrupd = ""
                Else
                prevnrupd = prevnr + nr1
                End If
                'get next number
                o = 0
                For l = fctr + 1 To Len(snail)
                    e = Mid(snail, l, 1)
                    
                    If IsNumeric(e) Then
                        nextnr = nextnr * 10 + e
                        o = o + 1
                    ElseIf o > 0 Then
                        Exit For
                    Else
                        postinbetweenstr = postinbetweenstr & e
                    End If
                Next l
                If l = Len(snail) + 1 Then
                nextnrupd = ""
                Else
                nextnrupd = nextnr + nr2
                End If
                wvgfdv = 1
                'create updated snail
                unchangepre = Left(snail, k)
                unchangepost = Right(snail, Len(snail) - l + 1)
                pre = unchangepre & prevnrupd
                Post = nextnrupd & unchangepost
                newstr = pre & preinbetweenstr & "0" & postinbetweenstr & Post
            
            
                snail = newstr
                GoTo restart_expl_check
            
            ElseIf c = "[" Then
                inbracks = inbracks + 1
            ElseIf c = "]" Then
                inbracks = inbracks - 1
            ElseIf c = "," Then
            
            Else 'number
            End If
        Next j
    Loop While exploded
    'split
    
    For j = 1 To Len(snail)
        c = Mid(snail, j, 1)
        
        If IsNumeric(c) Then
            nr_to_split = nr_to_split * 10 + c
            numbersadded = numbersadded + 1
        Else
            If nr_to_split > 9 Then
                prestr = Left(snail, j - numbersadded - 1)
                poststr = Right(snail, Len(snail) - j + 1)
                'split
                halfup = WorksheetFunction.Ceiling_Math(nr_to_split / 2)
                halfdown = WorksheetFunction.Floor_Math(nr_to_split / 2)
                
                newstr = prestr & "[" & halfdown & "," & halfup & "]" & poststr
                
                snail = newstr
                nr_to_split = 0
                numbersadded = 0
                GoTo restart_all_check
            End If
            nr_to_split = 0
            numbersadded = 0
        End If
    
    Next j
    

    
    
    
    
    
    
    
    
Loop While hassplit


x = 1
Next i
'do snail math
commasfound = True
Do While commasfound
commasfound = False
    lastcom = -1
    For fuckuvba = 1 To Len(snail)
        c = Mid(snail, fuckuvba, 1)
            If c = "[" Then
                lastinbr = fuckuvba
            ElseIf c = "]" Then
            
                If lastcom = -1 Then Exit For
                nr1 = Mid(snail, lastinbr + 1, lastcom - lastinbr - 1)
                nr2 = Mid(snail, lastcom + 1, fuckuvba - lastcom - 1)
                mag = 3 * nr1 + 2 * nr2
                
                newstr = Left(snail, lastinbr - 1) & mag & Right(snail, Len(snail) - fuckuvba)
                snail = newstr
                Exit For
            ElseIf c = "," Then
                lastcom = fuckuvba
                commasfound = True
            Else 'number
            End If


    Next fuckuvba
Loop
x = 2
Debug.Print snail

errhand:
Stop
Resume
End Sub















Sub day18b()
On Error GoTo errhand
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row


'make all unique snail babies
For firstsnail = 1 To lastrow
snail1 = Cells(firstsnail, 1).Value
    For secondsnail = 1 To lastrow
        If firstsnail = secondsnail Then
            'nothing
        Else
            snail = "[" & snail1 & "," & Cells(secondsnail, 1).Value & "]"

    'then do the same as part 1


Do 'nothing change
restart_all_check:
    hassplit = False
    Do  'nothing explodes
        'explode
restart_expl_check:
        exploded = False
        inbracks = 0
        nr1 = 0
        nr2 = 0
        nextnr = 0
        prevnr = 0
        preinbetweenstr = ""
        postinbetweenstr = ""
        For j = 1 To Len(snail)
            c = Mid(snail, j, 1)
            
            If inbracks = 5 Then 'go boom
            exploded = True
                'get two numbers
                fctr = j
                f = Mid(snail, fctr, 1)
                Do While Not f = ","
                    nr1 = nr1 * 10 + f
                    fctr = fctr + 1
                    f = Mid(snail, fctr, 1)
                Loop
                fctr = fctr + 1
                f = Mid(snail, fctr, 1)
                Do While Not f = "]"
                    nr2 = nr2 * 10 + f
                    fctr = fctr + 1
                    f = Mid(snail, fctr, 1)
                Loop
            
                'get previous number
                o = 0
                For k = j - 2 To 1 Step -1
                    d = Mid(snail, k, 1)
                   
                    If IsNumeric(d) Then
                        prevnr = WorksheetFunction.Power(10, o) * d + prevnr
                        o = o + 1
                    ElseIf o > 0 Then
                        Exit For
                    Else
                        preinbetweenstr = d & preinbetweenstr
                    End If
                Next k
                If k = 0 Then
                prevnrupd = ""
                Else
                prevnrupd = prevnr + nr1
                End If
                'get next number
                o = 0
                For l = fctr + 1 To Len(snail)
                    e = Mid(snail, l, 1)
                    
                    If IsNumeric(e) Then
                        nextnr = nextnr * 10 + e
                        o = o + 1
                    ElseIf o > 0 Then
                        Exit For
                    Else
                        postinbetweenstr = postinbetweenstr & e
                    End If
                Next l
                If l = Len(snail) + 1 Then
                nextnrupd = ""
                Else
                nextnrupd = nextnr + nr2
                End If
                wvgfdv = 1
                'create updated snail
                unchangepre = Left(snail, k)
                unchangepost = Right(snail, Len(snail) - l + 1)
                pre = unchangepre & prevnrupd
                Post = nextnrupd & unchangepost
                newstr = pre & preinbetweenstr & "0" & postinbetweenstr & Post
            
            
                snail = newstr
                GoTo restart_expl_check
            
            ElseIf c = "[" Then
                inbracks = inbracks + 1
            ElseIf c = "]" Then
                inbracks = inbracks - 1
            ElseIf c = "," Then
            
            Else 'number
            End If
        Next j
    Loop While exploded
    'split
    
    For j = 1 To Len(snail)
        c = Mid(snail, j, 1)
        
        If IsNumeric(c) Then
            nr_to_split = nr_to_split * 10 + c
            numbersadded = numbersadded + 1
        Else
            If nr_to_split > 9 Then
                prestr = Left(snail, j - numbersadded - 1)
                poststr = Right(snail, Len(snail) - j + 1)
                'split
                halfup = WorksheetFunction.Ceiling_Math(nr_to_split / 2)
                halfdown = WorksheetFunction.Floor_Math(nr_to_split / 2)
                
                newstr = prestr & "[" & halfdown & "," & halfup & "]" & poststr
                
                snail = newstr
                nr_to_split = 0
                numbersadded = 0
                GoTo restart_all_check
            End If
            nr_to_split = 0
            numbersadded = 0
        End If
    
    Next j
    

    
    
    
    
    
    
    
    
Loop While hassplit


x = 1
'Next i
'do snail math
commasfound = True
Do While commasfound
commasfound = False
    lastcom = -1
    For fuckuvba = 1 To Len(snail)
        c = Mid(snail, fuckuvba, 1)
            If c = "[" Then
                lastinbr = fuckuvba
            ElseIf c = "]" Then
            
                If lastcom = -1 Then Exit For
                nr1 = Mid(snail, lastinbr + 1, lastcom - lastinbr - 1)
                nr2 = Mid(snail, lastcom + 1, fuckuvba - lastcom - 1)
                mag = 3 * nr1 + 2 * nr2
                
                newstr = Left(snail, lastinbr - 1) & mag & Right(snail, Len(snail) - fuckuvba)
                snail = newstr
                Exit For
            ElseIf c = "," Then
                lastcom = fuckuvba
                commasfound = True
            Else 'number
            End If


    Next fuckuvba
Loop
x = 2
If CInt(snail) > CInt(snailmax) Then
snailmax = snail
End If



        End If
    Next secondsnail
Next firstsnail

Debug.Print snailmax

errhand:
Stop
Resume
End Sub

