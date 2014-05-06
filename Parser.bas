Attribute VB_Name = "Parser"
'mimoune.djouallah@gmail.com

Option Explicit

Global Const APPNAME = "XER READER"

Global sXerFileName     As String

Sub Loadxer()
Dim qt As QueryTable
Dim sht As Worksheet

For Each sht In ActiveWorkbook.Worksheets
If SheetExists("import") = True Then
Else
Worksheets.Add(After:=Worksheets("Dashboard")).Name = "import"
End If
Next

  
' bulk load the xer to the sheet "Import"


With ThisWorkbook.Sheets("import").QueryTables.Add(Connection:="TEXT;" & sXerFileName, Destination:=ThisWorkbook.Sheets("import").Range("$A$1"))
        .Name = "Capture"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
   
    
Application.StatusBar = "xer Loaded in Excel"
 DoEvents


For Each qt In ThisWorkbook.Sheets("import").QueryTables
qt.Delete
Next qt


    Exit Sub



End Sub

'load Tables position

Sub check_Tables()
Dim i As Long
Dim x As Long
Dim ds As Worksheet
Dim rs As Worksheet
Dim sht As Worksheet
Dim table As String

Set ds = Worksheets("import")
Set rs = Worksheets("Table_Summary")

rs.Cells(1, 1) = "XER"
rs.Cells(2, 1) = sXerFileName
rs.Columns("A:A").EntireColumn.AutoFit


rs.Cells(1, 5) = "Tables"
rs.Cells(1, 6) = "Row"
rs.Cells(1, 7) = " Nr record"

i = 2

For x = 1 To LastRow(ds)
If ds.Cells(x, 1) = "%T" Then

rs.Cells(i, 5) = ds.Cells(x, 2)

rs.Cells(i, 6) = x


i = i + 1
End If
Next x

' calculate th number of records
For i = 2 To LastRow(rs) - 1

rs.Cells(i, 7) = rs.Cells(i + 1, 6) - rs.Cells(i, 6) - 2


Next i
rs.Cells(LastRow(rs), 7) = LastRow(ds) - 2 - rs.Cells(i, 6)


'add sheets if they don't exist already

For i = 2 To LastRow(rs)
table = rs.Cells(i, 5)
For Each sht In ActiveWorkbook.Worksheets
If SheetExists(table) = True Then
Else
Worksheets.Add(After:=Worksheets("Table_Summary")).Name = table
Worksheets(table).Visible = False
End If
Next
Next i


For i = 2 To LastRow(rs)
Call rs.Hyperlinks.Add(rs.Cells(i, 5), "", rs.Cells(i, 5) & "!A1")
Next i

End Sub


Sub copy_Tables_New_sheets()

Dim i As Long
Dim x As Long
Dim ds As Worksheet
Dim ts As Worksheet
Dim NeSht As Worksheet
Dim namesht As String
Dim rowMn As Long
Dim rowMx As Long
Dim Nbrow As Long
Dim TASKP As Worksheet
Dim cell As Range
Dim tt As String

Worksheets("Table_summary").Visible = True

Set ds = Worksheets("import")
Set ts = Worksheets("Table_summary")
For i = 2 To LastRow(ts)
   rowMn = ts.Cells(i, 6)
   Nbrow = ts.Cells(i, 7)
   rowMx = rowMn + Nbrow
   namesht = ts.Cells(i, 5)
   
   
   Set NeSht = Worksheets(namesht)
   ds.Range("b" + CStr(rowMn + 1), ColumnLetter(LastCol(ds)) + CStr(rowMx + 1)).Copy NeSht.Range("a1", ColumnLetter(LastCol(ds) - 1) + CStr(Nbrow + 1))


Application.StatusBar = "Processing Table: " & namesht & "..."
 DoEvents
Next i
Application.DisplayAlerts = False

ds.Delete

For Each NeSht In ActiveWorkbook.Worksheets
If SheetExists("WBS") = True Then
Worksheets("WBS").Delete
End If
Next
Worksheets.Add(After:=Worksheets("Table_Summary")).Name = "WBS"



Worksheets("WBS").Visible = False

For Each NeSht In ActiveWorkbook.Worksheets
If SheetExists("Task_list") = True Then
Worksheets("Task_list").Delete
End If
Next
Worksheets.Add(After:=Worksheets("wbs")).Name = "Task_list"


Worksheets("Task_list").Visible = False

End Sub

