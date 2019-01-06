Option Explicit

Sub ChartInfo()

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim connStr As String
Dim providerStr As String
Dim sqlStr As String
Dim rowCount As Integer
    
' Clear previous reults
 Worksheets("Chart").Range(Range("N4"), Range("N4").End(xlDown)).ClearContents
 Worksheets("Chart").Range(Range("O4"), Range("O4").End(xlDown)).ClearContents
    
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

' Populate table
rs.Open sqlStr, cn
rowCount = 4

With rs
    Do Until .EOF
         Worksheets("Chart").Cells(rowCount, 14) = .Fields("Level")
        Worksheets("Chart").Cells(rowCount, 15) = .Fields("Time")
        rowCount = rowCount + 1
        .MoveNext
    Loop
End With


rs.Close

' ----- Queries to get stats ------


' SQL Query get stats for Easy
sqlStr = "SELECT COUNT(*) AS EasyCount, AVG(Time) AS EasyAverage, MIN(Time) AS EasyMin, MAX(Time) AS EasyMax "
sqlStr = sqlStr & "FROM SudokuSimulations "
sqlStr = sqlStr & "WHERE Level = 'Easy' "

' Fill in worksheet
rs.Open sqlStr, cn


With rs
       Worksheets("Chart").Range("C32").Value = .Fields("EasyCount")
       Worksheets("Chart").Range("C33").Value = .Fields("EasyAverage")
       Worksheets("Chart").Range("C34").Value = .Fields("EasyMax")
       Worksheets("Chart").Range("C35").Value = .Fields("EasyMin")
End With


rs.Close



' SQL Query get stats for Medium
sqlStr = "SELECT COUNT(*) AS MediumCount, AVG(Time) AS MediumAverage, MIN(Time) AS MediumMin, MAX(Time) AS MediumMax "
sqlStr = sqlStr & "FROM SudokuSimulations "
sqlStr = sqlStr & "WHERE Level = 'Medium' "

' Fill in worksheet
rs.Open sqlStr, cn


With rs
       Worksheets("Chart").Range("F32").Value = .Fields("MediumCount")
       Worksheets("Chart").Range("F33").Value = .Fields("MediumAverage")
       Worksheets("Chart").Range("F34").Value = .Fields("MediumMax")
       Worksheets("Chart").Range("F35").Value = .Fields("MediumMin")
End With


rs.Close


' SQL Query get stats for Hard
sqlStr = "SELECT COUNT(*) AS HardCount, AVG(Time) AS HardAverage, MIN(Time) AS HardMin, MAX(Time) AS HardMax "
sqlStr = sqlStr & "FROM SudokuSimulations "
sqlStr = sqlStr & "WHERE Level = 'Hard' "

' Fill in worksheet
rs.Open sqlStr, cn


With rs
       Worksheets("Chart").Range("I32").Value = .Fields("HardCount")
       Worksheets("Chart").Range("I33").Value = .Fields("HardAverage")
       Worksheets("Chart").Range("I34").Value = .Fields("HardMax")
       Worksheets("Chart").Range("I35").Value = .Fields("HardMin")
End With


rs.Close
cn.Close

Set rs = Nothing

Worksheets("Chart").PivotTables(1).PivotCache.Refresh

End Sub

Sub GoToChart() ' Takes user from solver page to chart page

ThisWorkbook.Sheets("Chart").Activate

End Sub


