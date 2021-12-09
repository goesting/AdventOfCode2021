Sub day6a()
Start = Timer                                               'Start the clock
inp = Split(Cells(1, 1).Value, ",")                         'read and split the input
ticks = 80                                                  'days we need to simulate

'------ Init the array holding the fishies ----------------
Dim arrFish() As Variant
ReDim arrFish(0 To 8)
For Each fish In inp
    arrFish(CInt(fish)) = arrFish(CInt(fish)) + 1
Next fish
'----------------------------------------------------------

'------ Run the simulation the required amount ------------
For i = 1 To ticks
    Call doTick(arrFish)
Next i
'----------------------------------------------------------

'------ Count all Fishy McFishersons ----------------------
For i = LBound(arrFish) To UBound(arrFish)
    result = result + arrFish(i)
Next i
'----------------------------------------------------------

'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub

Sub day6b() 'Exactly the same code, just more loops :3
Start = Timer                                               'Start the clock
inp = Split(Cells(1, 1).Value, ",")                         'read and split the input
ticks = 256                                                 'days we need to simulate

'------ Init the array holding the fishies ----------------
Dim arrFish() As Variant
ReDim arrFish(0 To 8)
For Each fish In inp
    arrFish(CInt(fish)) = arrFish(CInt(fish)) + 1
Next fish
'----------------------------------------------------------

'------ Run the simulation the required amount ------------
For i = 1 To ticks
    Call doTick(arrFish)
Next i
'----------------------------------------------------------

'------ Count all Fishy McFishersons ----------------------
For i = LBound(arrFish) To UBound(arrFish)
    result = result + arrFish(i)
Next i
'----------------------------------------------------------

'------ Output results and ti;e taken ---------------------
Debug.Print "Solution to part 2: " & result
Debug.Print "Time taken = " & Timer - Start & " seconds"
'----------------------------------------------------------
End Sub

Sub doTick(ByRef arr)
'----------------------------------------------------------
'- Run one day to the future for all fishes               -
'- Spawning new ones when a fish reaches 0                -
'----------------------------------------------------------

Dim temp() As Variant                   'array to hold new iteration
ReDim temp(0 To 8)

For i = LBound(temp) To UBound(temp)    'Loop through all 9 of the "day-populations"
    Select Case i
        Case 8:                         '8 gets filled by new babies from all the day-0
            temp(i) = arr(0)
        Case 6:                         '6 get both the babies coming from 7 and the resetting fish from 0
            temp(i) = arr(7) + arr(0)
        Case Else:                      'All other tick 1 down
            temp(i) = arr(i + 1)
    End Select
Next i

arr = temp
End Sub
