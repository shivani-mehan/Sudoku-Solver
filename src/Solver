Option Explicit

' Global variable
Dim time As Single

Sub Sudoku()

Dim sudokuArray() As Variant
Dim counter As Integer
Dim i As Integer
Dim j As Integer
Dim difficulty As String

' Clear any previous statistics
With Worksheets("Solver")
    .Cells(12, 8).MergeArea.ClearContents
    .Cells(13, 4).MergeArea.ClearContents
    .Range("K13:L13").ClearContents
End With

' If empty board
If WorksheetFunction.CountA(Worksheets("Solver").Range("D3:L11")) = 1 Or WorksheetFunction.CountA(Range("D3:L11")) = 0 Then
    MsgBox "You can't solve an empty puzzle!", vbCritical + vbRetryCancel, "Oops!"
    Exit Sub
ElseIf WorksheetFunction.CountA(Worksheets("Solver").Range("D3:L11")) < 20 Then 'Invalid sudoku
    MsgBox "Invalid Sudoku", vbCritical + vbRetryCancel, "Oops!"
    Call Clear
    Exit Sub
End If

' Initiate time
time = 0
    
' Put board into array
sudokuArray = Worksheets("Solver").Range("D3:L11").Value
    
' Figure out how many numbers are filled in to determine difficulty
counter = 0
        
For i = 1 To 9
    For j = 1 To 9
        If sudokuArray(i, j) <> "" Then
            counter = counter + 1
        End If
    Next j
Next i
    
If counter > 30 Then
    difficulty = "Easy"
ElseIf 25 <= counter And counter <= 30 Then
    difficulty = "Medium"
Else
    difficulty = "Hard"
End If
    
' Solve Puzzle
Call SolvePuzzle(sudokuArray, 1, 1)
    
    
' Fill in completion details
With Worksheets("Solver")
    .Cells(12, 8).MergeArea.Value = difficulty
    .Cells(13, 4).MergeArea.Value = "Done!"
    .Range("K13").Value = "Time:"
    .Range("L13").Value = time
End With


' Update Database
' Call DatabaseUpdate
    
End Sub
Sub DatabaseUpdate()

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim connStr As String
Dim providerStr As String
Dim sqlStr As String
    
' Database is sitting in the current folder
connStr = "Data Source=" & ThisWorkbook.Path & "\shivaniMehan_a5.accdb"
providerStr = "Microsoft.ACE.OLEDB.12.0"
        
' Open the Connection string
With cn
    .ConnectionString = connStr
    .Provider = providerStr
    .Open
End With
        
' SQL Query access table
sqlStr = "SELECT * FROM SudokuSimulations"

' Upddate database with information
rs.Open sqlStr, cn, adOpenDynamic, adLockOptimistic

With rs
    .AddNew
    .Fields("Level") = Worksheets("Solver").Cells(12, 8).MergeArea.Value
    .Fields("Time") = Worksheets("Solver").Range("L13").Value
    .Update
End With

rs.Close
cn.Close

Set rs = Nothing

End Sub


Function CheckConditions(sudokuArray, row As Integer, col As Integer, num As Integer)

Dim i As Integer
Dim j As Integer
Dim rowBoundary
Dim colBoundary

' Check row for the number
For i = 1 To 9
    If sudokuArray(row, i) = num Then
        CheckConditions = False
        Exit Function
    End If
Next i

' If not in row, check column for number
For i = 1 To 9
    If sudokuArray(i, col) = num Then
        CheckConditions = False
        Exit Function
    End If
Next i

' If not in row or column, check 3x3 box for number
 
 ' Figure out column boundary
If col < 4 Then
    'rowBoundary = 1
   colBoundary = 1
ElseIf col < 7 Then
   ' rowBoundary = 4
   colBoundary = 4
Else
    colBoundary = 7
   ' rowBoundary = 7
End If
 
 ' Figure out location of cell and row boundary
 If row < 4 Then
     'colBoundary = 1
    rowBoundary = 1
ElseIf row < 7 Then
    'colBoundary = 4
    rowBoundary = 4
Else
    'colBoundary = 7
   rowBoundary = 7
End If

    
' Iterate through the 3x3 square
For i = rowBoundary To rowBoundary + 2
    For j = colBoundary To colBoundary + 2
        If sudokuArray(i, j) = num Then
            CheckConditions = False
            Exit Function
        End If
    Next j
Next i

' If number passed all of the conditions and is therefore valid
CheckConditions = True

End Function
Function SolvePuzzle(sudokuArray, row As Integer, col As Integer)

Dim num As Integer
Dim solved As Boolean
Dim tempRow As Integer
Dim tempCol As Integer
Dim startTime As Single
Dim elapsedTime As Single

startTime = Timer()

' Base Case: row > 9 or col > 9, update the puzzle
If row > 9 Or col > 9 Then
    solved = UpdatePuzzle(sudokuArray)
    Exit Function
Else

    ' For other cases, iterate left to right and top to bottom
    ' If col = 9, need to proceed to next row by setting col to 1 and row to +1

    tempRow = row
    tempCol = col

    If col = 9 Then
        tempCol = 1
        tempRow = row + 1
    Else
        tempCol = tempCol + 1
        tempRow = row
    End If

    ' Case 1: cell already has value in it
    If (sudokuArray(row, col) <> "") Then
        solved = SolvePuzzle(sudokuArray, tempRow, tempCol)
        
    Else ' Case 2: doesn't have value in it
        For num = 1 To 9
            If (CheckConditions(sudokuArray, row, col, num) = True) Then
                 sudokuArray(row, col) = num
                 solved = SolvePuzzle(sudokuArray, tempRow, tempCol)
            End If
            
           ' sudokuArray(row, col) = ""
            
        Next num
    
    ' Case 3: none of the values work, need to iterate through and fix a previous cell
    
    sudokuArray(row, col) = ""
    
    End If
    
End If
 
elapsedTime = elapsedTime + (Timer() - startTime)
time = elapsedTime

End Function
Function UpdatePuzzle(sudokuArray)

Dim i As Integer
Dim j As Integer

For i = 1 To 9
    For j = 1 To 9
       Worksheets("Solver").Cells(2 + i, 3 + j).Value = sudokuArray(i, j)
    Next j
Next i

End Function
Sub Clear() ' Clear results of the game and set back to default

Worksheets("Solver").Range("D3:L11").ClearContents

With Worksheets("Solver").Range("D3")
    .Value = "Paste Here!"
    .Font.Name = "Calibri"
    .Font.size = 8
End With

Worksheets("Solver").Cells(12, 8).MergeArea.ClearContents
