Sub day4()
'-------INIT ALL THE BOARDS------
lastrow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row 'size of data
nr_of_boards = (lastrow - 1) / 6                            'each board is 5 lines, plus one blank line
Dim arrBoard() As Variant                                   '-Create array holding all boards-
ReDim arrBoard(1 To nr_of_boards) ', 0 To 4, 0 To 4)        '1d array of 2d boards
Dim board() As Variant                                      'Create board variable, 5*5 array
ReDim board(0 To 4, 0 To 4)

board_index = 1                                                                 'keeps track of which board are we looking at
For cellrow = 3 To lastrow                                                      'go through all the input line by line, starting at 3
    board_row = (cellrow - 2) Mod 6                                             'calculate which row we are looking at of the board (1-5), 0 is the empty row
    If board_row <> 0 Then                                                      'if not empty row
        inp = Split(Replace(Trim(Cells(cellrow, 1).Value), "  ", " "), " ")     'get input, trim and replace extra spaces so we can split on single space
            
        For j = LBound(inp) To UBound(inp)                                      'go through all values of a single input line
            If inp(j) <> 0 Then
                board(board_row - 1, j) = inp(j)                                'store them in the board array variable
            Else
               board(board_row - 1, j) = "1000"                                 ' HACK because 0 can actually show up as a number, and I use 0 to show that a number has been used
            End If
            'arrBoard(board_index, board_row - 1, j) = inp(j)
        Next j
        
    Else
        arrBoard(board_index) = board                                           'if we are on the empty row, data of 1 board is finished, shove that 2d array, into the 1d array
        ReDim board(0 To 4, 0 To 4)                                             'reset the board array for the next one
        board_index = board_index + 1                                           'keep track of which board we are doing
    End If
Next cellrow
arrBoard(board_index) = board                                                   'doesnt get to empty line under last, so do it once more manually when finished


Dim haswon() As Variant                                                         'for PART 2: keep track of who has won and who has not, need last one not won
ReDim haswon(1 To UBound(arrBoard))

'------------END INIT---------------


'--------------DRAW-----------------
draw = Split(Cells(1, 1).Value, ",")                                            'read draw data
For d = LBound(draw) To UBound(draw)
    drawn = draw(d)                                                             'get the draw number, if 0 set to 1000(see above)
    If drawn = "0" Then drawn = "1000"
    
    For i = LBound(arrBoard, 1) To UBound(arrBoard, 1)
        For j = LBound(arrBoard(i), 1) To UBound(arrBoard(i), 1)
            For k = LBound(arrBoard(i), 2) To UBound(arrBoard(i), 2)            'triple nested for loop to check all boards for drawn number,
                nr = arrBoard(i)(j, k)
                If nr = drawn Then
                    arrBoard(i)(j, k) = 0                                       'if found, set to 0 and jump out of for loops
                    GoTo Nextboard
                End If
            Next k
        Next j
Nextboard:                                                                      'jumping goto point to skip board if nr has been found. no number appears twice
    Next i

    For i = LBound(arrBoard, 1) To UBound(arrBoard, 1)                          'Once all board have been checked for drawn number
        If haswon(i) <> True Then                                               'loop through all boards, and if they have not won yet, pass them to the checkwinner() function
            result = checkForWinner(arrBoard(i))                                'returns sum of remaining number if winner, or 0 if not winner
        End If
        
        If result <> 0 Then
            Debug.Print result * drawn                                          'PART1 SOL: first returned value
            haswon(i) = True
            result = 0
        End If
    Next i
Next d


Debug.Print "done"                                                              'PART2 SOL: last printed value before "done"
End Sub


'--------------------------------------------------------------------------------------------------------------------------------
Function checkForWinner(arr As Variant)         'checks all rows and columns for one that only contains zeros
                                                'berekent ook de totale waarde dat er nog in zit, en geeft die door indien winner
winner = False

'-----check rows-----
totalsum = 0

For i = LBound(arr, 1) To UBound(arr, 1)        'Loop the rows, then the each nr in a row, and add them all up
    rowsum = 0
    For j = LBound(arr, 2) To UBound(arr, 2)
        rowsum = rowsum + arr(i, j)
    Next j
    totalsum = totalsum + rowsum
    If rowsum = 0 Then                          'sum should be 0 to be a completed row, because thats how we set it
        winner = True
    End If
Next i

'-----check columns-----
For i = LBound(arr, 2) To UBound(arr, 2)        'Do the same but reverse order of rows and columns
    rowsum = 0
    For j = LBound(arr, 1) To UBound(arr, 1)
        rowsum = rowsum + arr(j, i)
    Next j
If rowsum = 0 Then
    winner = True
End If

Next i


'-----check diagonals-----                GVD tijd verkakt aan diagonals, why I no read gud?
'rowsum = 0
'For i = LBound(arr, 2) To UBound(arr, 2)
'rowsum = rowsum + arr(i, i)
'Next i
'If rowsum = 0 Then
'winner = True
'End If

'rowsum = 0
'For i = LBound(arr, 2) To UBound(arr, 2)
'rowsum = rowsum + arr(i, UBound(arr, 2) - i)
'Next i
'If rowsum = 0 Then
'winner = True
'End If


If winner Then                          'if this board is a winner, return the total sum of all remaining numbers
    checkForWinner = totalsum Mod 1000  'Mod 1000 omdat originele nummers 0 omgezet zijn naar nummer 1000, dus die er terug afhalen
End If


End Function
