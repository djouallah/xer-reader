Attribute VB_Name = "functions"
Option Explicit


Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function
Function LastRow(sh As Worksheet)
    On Error Resume Next
    LastRow = sh.Cells.Find(what:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function

Function LastCol(sh As Worksheet)
    On Error Resume Next
    LastCol = sh.Cells.Find(what:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function
Function ColumnLetter(ColumnNumber As Integer) As String
    Dim n As Integer
    Dim c As Byte
    Dim s As String

    n = ColumnNumber
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColumnLetter = s
End Function

Function dhDaysInMonth(Optional dtmDate As Date = 0) As Integer
    ' Return the number of days in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhDaysInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 1) - _
     DateSerial(Year(dtmDate), Month(dtmDate), 1)
End Function

Function Yearday(dat As Date) As Integer

    Yearday = (DateSerial(Year(dat), 12, 31) - DateSerial(Year(dat), 1, 1)) + 1
End Function

Public Function MConcate(rCells As Range, stDelim As String) As String
    Dim vArray As Variant
    Dim lLoop As Long
     
    ReDim vArray(1 To rCells.Count)
    For lLoop = 1 To UBound(vArray)
        vArray(lLoop) = rCells(1, lLoop)
    Next lLoop
     
    MConcate = Join(vArray, stDelim)
End Function
Sub DeletSheets()
Dim ws As Worksheet
 
    Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = "Dashboard" Or ws.Name = "Details_report" Or ws.Name = "gantt" Or ws.Name = "WBS" Or ws.Name = "Table_Summary" Then
    Worksheets("Table_Summary").Visible = False
    Worksheets("Table_Summary").Cells.Clear
    
    
    Else
    ws.Delete
    End If
             
       
    Next ws
    Application.DisplayAlerts = True
    

End Sub
    
Sub clearDashboard()

With Worksheets("Dashboard")
 .Range("a4:j8").Cells.ClearContents
 .Range("b10:b16").Cells.ClearContents
 .Range("c24:r24").Cells.ClearContents
 .Range("a25:r25").Cells.ClearContents
 .Range("e10:e14").Cells.ClearContents
 .Range("h10:h15").Cells.ClearContents
 .Range("k10:k15").Cells.ClearContents
 .Columns("v").Hidden = True
 .Range("V2:V" & LastRow(Worksheets("dashboard"))).Cells.ClearContents
 .Rows("9:28").EntireRow.Hidden = True
End With

End Sub
Sub Cleardetails()
Dim ws As Worksheet
Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Worksheets
    If ws.Name = "Dashboard" Or ws.Name = "Table_Summary" Then
    Worksheets("Table_Summary").Visible = False
    Worksheets("Table_Summary").Cells.Clear
   
    Else
        ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True

End Sub

