Option Explicit
Sub day16a() 'add up all version numbers
heks = Cells(1, 1).Value

Dim htob() As String
ReDim htob(0 To 15)
htob(0) = "0000"
htob(1) = "0001"    'lookup quicker than actually converting
htob(2) = "0010"
htob(3) = "0011"
htob(4) = "0100"
htob(5) = "0101"
htob(6) = "0110"
htob(7) = "0111"
htob(8) = "1000"
htob(9) = "1001"
htob(10) = "1010"
htob(11) = "1011"
htob(12) = "1100"
htob(13) = "1101"
htob(14) = "1110"
htob(15) = "1111"

For i = 1 To Len(heks)
    c = Mid(heks, i, 1)
    If Asc(c) > 64 Then c = Asc(c) - 55
    binc = htob(c)
    binstring = binstring & binc
Next i

result = getVersionSum(binstring)
Debug.Print result
End Sub


Function getVersionSum(s)
On Error GoTo errhands
Dim binlength As String
Dim version As String
Dim binpackets As String
pointer = 1
Do While pointer < Len(s) - 9
'at version location
version = Mid(s, pointer, 3)


label = Mid(s, pointer + 3, 3)

pointer = pointer + 6
If label = "100" Then 'Literal value
    version = Bin2Dec(version)
    Do While True
    If Mid(s, pointer, 1) = 0 Then 'last digits of the literal
        pointer = pointer + 5
        Exit Do
    Else 'keep going
        pointer = pointer + 5
    End If
    Loop
Else ' operator
    lengthtypeID = Mid(s, pointer, 1)
    If lengthtypeID = 0 Then
        binlength = Mid(s, pointer + 1, 15)
        l = Bin2Dec(binlength)
        pointer = pointer + 16
        VersionSum = VersionSum + Bin2Dec(version)
        version = getVersionSum(Mid(s, pointer, l))
        pointer = pointer + l
    Else 'ID = 1
        binpackets = Mid(s, pointer + 1, 11)
        nr_p = Bin2Dec(binpackets)
        VersionSum = VersionSum + Bin2Dec(version)
        version = getVersionSum(Right(s, (Len(s) - (pointer + 11))))
        pointer = Len(s)


    End If
    
End If
VersionSum = CInt(version) + CInt(VersionSum)
Loop
finished:
getVersionSum = VersionSum
GoTo continu
errhands:
Stop
Resume
continu:
End Function
Function Bin2Dec(sMyBin As String) As Variant
    Dim x As Integer
    Dim iLen As Integer

    iLen = Len(sMyBin) - 1
    For x = 0 To iLen
        Bin2Dec = Bin2Dec + _
          Mid(sMyBin, iLen - x + 1, 1) * 2 ^ x
    Next
End Function

Sub day16b() 'add up all values depending on type ID
Static heks As String
Dim c As Variant
Dim binc As String
Dim binstring As String
Dim pointer As Integer
Dim result As Variant
Dim i As Integer

heks = Cells(1, 1).Value

Static htob() As String
ReDim htob(0 To 15)
htob(0) = "0000"
htob(1) = "0001"    'lookup quicker than actually converting
htob(2) = "0010"
htob(3) = "0011"
htob(4) = "0100"
htob(5) = "0101"
htob(6) = "0110"
htob(7) = "0111"
htob(8) = "1000"
htob(9) = "1001"
htob(10) = "1010"
htob(11) = "1011"
htob(12) = "1100"
htob(13) = "1101"
htob(14) = "1110"
htob(15) = "1111"

For i = 1 To Len(heks)
    c = Mid(heks, i, 1)
    If Asc(c) > 64 Then c = Asc(c) - 55
    binc = htob(c)
    binstring = binstring & binc
Next i
pointer = 1
result = getExpressionsum(binstring, pointer) ', blocks_done)
Debug.Print result
End Sub
Function getExpressionsum(s, ByRef pointer) ', ByVal blocksdone) ', Optional leng = 0, Optional subs = 0)
On Error GoTo errhands
'Dim version As String
Dim bitv As String, label As String, bit As String, lengthtypeID As String, included_packets As String

Dim i As Long, number_of_bits As Long, included_length As Long, startpoint As Long, blocks As Long
Dim exsumpart As Variant, exproductpart As Variant, min_compare As Variant, max_compare As Variant, exBigger1 As Variant, exBigger2 As Variant
Dim exEquals1 As Variant, exEquals2 As Variant, exSmaller1 As Variant, exSmaller2 As Variant

If pointer > 5550 Then

End If
'Do While pointer < Len(s) - 9 'needs to change for part 2 Do While block_complete = false? Do i even need a loop
    'each string passed in here starts with version and label
    'version = Mid(s, pointer, 3)
    label = Mid(s, pointer + 3, 3)
    'move pointer to behind version and label
    pointer = pointer + 6
    
    'check label to figure out what and how much data is behind
    If label = "100" Then 'Literal value = END NODE
        bitv = ""
        Do
            bit = Mid(s, pointer, 1)
            pointer = pointer + 1
            bitv = bitv & Mid(s, pointer, 4)
            pointer = pointer + 4
            
        
        Loop While bit = 1
        getExpressionsum = Bin2Dec(bitv)
    
    ElseIf label = "000" Then  'SUM
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            startpoint = pointer
            Do While pointer < startpoint + included_length
                exsumpart = getExpressionsum(s, pointer) ', blocksdone)
                getExpressionsum = getExpressionsum + exsumpart
            Loop
        Else 'ID = 1   - 11 bits is the amount of sub packets
            included_packets = Bin2Dec(Mid(s, pointer, 11))
            pointer = pointer + 11
            blocks = 0
            Do While blocks < CLng(included_packets)
                exsumpart = getExpressionsum(s, pointer) ', blocksdone)
                blocks = blocks + 1
                getExpressionsum = getExpressionsum + exsumpart
            Loop
        End If
    
    ElseIf label = "001" Then  'PRODUCT
        getExpressionsum = 1
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            startpoint = pointer
            Do While pointer < startpoint + included_length
                exproductpart = getExpressionsum(s, pointer) ', blocksdone)
                getExpressionsum = getExpressionsum * exproductpart
            Loop
        Else 'ID = 1   - 11 bits is the amount of sub packets
            included_packets = Bin2Dec(Mid(s, pointer, 11))
            pointer = pointer + 11
            blocks = 0
            Do While blocks < CLng(included_packets)
                exproductpart = getExpressionsum(s, pointer) ', blocksdone)
                blocks = blocks + 1
                getExpressionsum = getExpressionsum * exproductpart
            Loop
        End If
    
    ElseIf label = "010" Then  'MIN
        getExpressionsum = -1
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            startpoint = pointer
            Do While pointer < startpoint + included_length
                min_compare = getExpressionsum(s, pointer) ', blocksdone)
                If getExpressionsum = -1 Then
                    getExpressionsum = min_compare
                ElseIf min_compare < getExpressionsum Then
                    getExpressionsum = min_compare
                End If
            Loop
        Else 'ID = 1   - 11 bits is the amount of sub packets
            included_packets = Bin2Dec(Mid(s, pointer, 11))
            pointer = pointer + 11
            blocks = 0
            Do While blocks < CLng(included_packets)
                min_compare = getExpressionsum(s, pointer) ', blocksdone)
                blocks = blocks + 1
                If getExpressionsum = -1 Then
                    getExpressionsum = min_compare
                ElseIf min_compare < getExpressionsum Then
                    getExpressionsum = min_compare
                End If
            Loop
        End If
    
    ElseIf label = "011" Then  'MAX
        getExpressionsum = -1
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            startpoint = pointer
            Do While pointer < startpoint + included_length
                max_compare = getExpressionsum(s, pointer) ', blocksdone)
                If max_compare > getExpressionsum Then
                    getExpressionsum = max_compare
                End If
            Loop
        Else 'ID = 1   - 11 bits is the amount of sub packets
            included_packets = Bin2Dec(Mid(s, pointer, 11))
            pointer = pointer + 11
            blocks = 0
            Do While blocks < CLng(included_packets)
                max_compare = getExpressionsum(s, pointer) ', blocksdone)
                blocks = blocks + 1
                If getExpressionsum = -1 Then
                    getExpressionsum = max_compare
                ElseIf max_compare > getExpressionsum Then
                    getExpressionsum = max_compare
                End If
            Loop
        End If
    
    ElseIf label = "101" Then  '1 > 2
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            'included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            
            exBigger1 = getExpressionsum(s, pointer) ', blocksdone)
            exBigger2 = getExpressionsum(s, pointer) ', blocksdone)
            getExpressionsum = Abs(CInt(exBigger1 > exBigger2)) 'true = -1, false = 0, thanks vba
        Else 'ID = 1   - 11 bits is the amount of sub packets, should be 2
            'included_packets = Bin2Dec(Mid(s, pointer, 11)) ' = 2
            pointer = pointer + 11
            'blocks = 0

                exBigger1 = getExpressionsum(s, pointer) ', blocksdone)
                'blocks = blocks + 1
                exBigger2 = getExpressionsum(s, pointer) ', blocksdone)
                'blocks = blocks + 1
                getExpressionsum = Abs(CInt(exBigger1 > exBigger2)) 'true = -1, false = 0, thanks vba

        End If
    
    ElseIf label = "110" Then  '2 > 1
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            'included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            
            exSmaller1 = getExpressionsum(s, pointer) ', blocksdone)
            exSmaller2 = getExpressionsum(s, pointer) ', blocksdone)
            getExpressionsum = Abs(CInt(exSmaller1 < exSmaller2))
        Else 'ID = 1   - 11 bits is the amount of sub packets, should be 2
            'included_packets = Bin2Dec(Mid(s, pointer, 11)) ' = 2
            pointer = pointer + 11
            'blocks = 0

                exSmaller1 = getExpressionsum(s, pointer) ', blocksdone)
                'blocks = blocks + 1
                exSmaller2 = getExpressionsum(s, pointer) ', blocksdone)
                'blocks = blocks + 1
                getExpressionsum = Abs(CInt(exSmaller1 < exSmaller2)) 'true = -1, false = 0, thanks vba

        End If
    
    ElseIf label = "111" Then  '1 ==2
        lengthtypeID = Mid(s, pointer, 1)
        pointer = pointer + 1
        If lengthtypeID = 0 Then '15bits are the lenth of next part
            'included_length = Bin2Dec(Mid(s, pointer, 15))
            pointer = pointer + 15 'set pointer at first bit of subtype
            
            exEquals1 = getExpressionsum(s, pointer) ', blocksdone)
            exEquals2 = getExpressionsum(s, pointer) ', blocksdone)
            getExpressionsum = Abs(CInt(exEquals1 = exEquals2))
        Else 'ID = 1   - 11 bits is the amount of sub packets, should be 2
            included_packets = Bin2Dec(Mid(s, pointer, 11)) ' = 2
            pointer = pointer + 11
            'blocks = 0

                exEquals1 = getExpressionsum(s, pointer) ', blocksdone)
                'blocks = blocks + 1
                exEquals2 = getExpressionsum(s, pointer) ', blocksdone)
                'blocks = blocks + 1
                getExpressionsum = Abs(CInt(exEquals1 = exEquals2)) 'true = -1, false = 0, thanks vba

        End If
    
    End If
'Loop




finished:
'getExpressionsum = exSum 'send to caller function
GoTo continu
errhands:
Stop
Resume
continu:
End Function
Sub generatebincode()
inp = Cells(1, 1).Value




End Sub


Sub tests()
Dim varx As String
Debug.Print test_val(25)
End Sub
Function test_val(x)
Dim s As String
Dim digits4 As String
Dim binstr As String
Static htob() As String
Dim lastdigits As Boolean
ReDim htob(0 To 15)
htob(0) = "0000"
htob(1) = "0001"    'lookup quicker than actually converting
htob(2) = "0010"
htob(3) = "0011"
htob(4) = "0100"
htob(5) = "0101"
htob(6) = "0110"
htob(7) = "0111"
htob(8) = "1000"
htob(9) = "1001"
htob(10) = "1010"
htob(11) = "1011"
htob(12) = "1100"
htob(13) = "1101"
htob(14) = "1110"
htob(15) = "1111"

s = "111100"
lastdigits = True
Do While Not x = 0
digits4 = x Mod 16

x = (x - digits4) / 16
If lastdigits Then
    binstr = "0" & htob(digits4) & binstr
    lastdigits = False
Else
    binstr = "1" & htob(digits4) & binstr
End If
Loop
s = s & binstr
test_val = s
End Function
