'Timing Functions
Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long
Sub day17a()
On Error GoTo errhand:
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
'no matter how high you launch, eventually it will hit the exact same height as where you started
'with the same speed as when you launched it (just in the opposite direction)
'so the highest speed it could arrive here, would be if the next step would take it all the way out to the minimum
'of the target range in one go
ymin = CInt(Split(Split(Split(Cells(1, 1).Value, ", ")(1), "=")(1), "..")(0)) ' get minimum of the target range
maxdownspeed_at_equal = Abs(ymin) - 1                                         ' maximum allowed speed to barely hit the target range
result = (maxdownspeed_at_equal + 1) * (maxdownspeed_at_equal / 2)            ' highest point is v + v-1 + v-2 + ... + 2 + 1 - > math - > sum = (max+min) * elements/2
                                                                              ' with v being the start speed (which is equal to the maxdownspeed at equal
'one line output for shits and giggles
Debug.Print (Abs(CInt(Split(Split(Split(Cells(1, 1).Value, ", ")(1), "=")(1), "..")(0)))) * (Abs(CInt(Split(Split(Split(Cells(1, 1).Value, ", ")(1), "=")(1), "..")(0))) - 1 / 2)

'------ Output results and time taken ---------------------
getTime endTime
timeElapsed = (endTime - startTime) / perSecond
Debug.Print "Solution to part 1: " & result
Debug.Print "Time taken = " & timeElapsed & " seconds"
Debug.Print ""
'----------------------------------------------------------
errhand: 'no real handling, just stopping to manually check it
Stop
Resume

End Sub

Sub day17b()
On Error GoTo errhand:
'variables for timing
   Dim startTime As Currency
   Dim endTime As Currency
   Dim perSecond As Currency
   Dim timeElapsed As Double
   getFrequency perSecond
   getTime startTime
'----start code------
x_str = Split(Cells(1, 1).Value, ", ")(0)
y_str = Split(Cells(1, 1).Value, ", ")(1)
xmin = CInt(Split(Split(x_str, "=")(1), "..")(0))
xmax = CInt(Split(Split(x_str, "=")(1), "..")(1))
ymin = CInt(Split(Split(y_str, "=")(1), "..")(0))
ymax = CInt(Split(Split(y_str, "=")(1), "..")(1))
vy_min = ymin
vy_max = Abs(ymin) - 1
For vy = vy_min To vy_max
    steps = 0
    nextspeed = vy
    If vy > 0 Then 'when launching up
        steps = steps + vy * 2 + 1 ' up to launch height
        nextspeed = vy + 1
        minsteps = getMinSteps(nextspeed * -1, ymax) + (vy * 2 + 1)
        maxsteps = getMaxSteps(nextspeed * -1, ymin) + (vy * 2 + 1)
    Else
    minsteps = getMinSteps(nextspeed, ymax)
    maxsteps = getMaxSteps(nextspeed, ymin)
    End If
    'get all vx that can land between xmin and xmax in a value between minsteps and maxsteps
    vx_min = 0                 'if vx = i then after x steps you are :   x*( i - ((x-1)/2))
    For s = minsteps To maxsteps
        min_vx = getMinVx(s, xmin)
        max_vx = getMaxVx(s, xmax)
        result = result + max_vx - min_vx + 1
    Next s
Next vy
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
Function getMinSteps(v0, ymax)  ' not ok
Sum = 0
i = 0
Do While Sum > ymax
Sum = Sum + v0 - i
i = i + 1
Loop
getMinSteps = i
End Function
Function getMaxSteps(v0, ymin) ' not ok
Sum = 0
i = 0
Do While Sum >= ymin
Sum = Sum + v0 - i
i = i + 1
Loop
getMaxSteps = i - 1
End Function
Function getMinVx(steps, xmin)
'land on (at least) xmin with 1 horizontal remaining
'  (minvx+1) * minvx =  2*xmin  ==> minvx = (-1 + sqrt( 1-8*xmin)
'getMinVx = WorksheetFunction.Ceiling_Math((-1 + Sqr((1 + 8 * xmin)) / 2)) incorrect, this is absolute min, not relative
avg_step = xmin / steps
getMinVx = WorksheetFunction.Ceiling_Math(avg_step + (steps - 1) / 2)
If getMinVx < steps Then ' horizontal speed will reach zero
    getMinVx = WorksheetFunction.Ceiling_Math((-1 + Sqr((1 + 8 * xmin)) / 2))
End If

End Function
Function getMaxVx(steps, xmax)
avg_step = xmax / steps
getMaxVx = WorksheetFunction.Floor_Math(avg_step + (steps - 1) / 2)
If getMaxVx < steps Then ' horizontal speed will reach zero
    getMaxVx = WorksheetFunction.Ceiling_Math((-1 + Sqr((1 + 8 * xmax)) / 2))
End If
End Function
Sub etst()
x = getMinVx(4, 18)
End Sub
