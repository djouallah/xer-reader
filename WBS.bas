Attribute VB_Name = "WBS"
Option Explicit
Global LevelMax As Long

Sub Create_Temp_Tables()
Dim sht As Worksheet
Dim i As Integer
Dim cnn As Object
Dim rssql As Object
Dim strsql As String


Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")
rssql.CursorLocation = 3

cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"


i = 0
For Each sht In ActiveWorkbook.Worksheets
If SheetExists("projwbs") = True Then
i = i + 1
End If
Next

If i = 0 Then
Exit Sub
End If


For Each sht In ActiveWorkbook.Worksheets
If SheetExists("tempWBS") = True Then
Worksheets("tempWBS").Delete
End If
Next
Worksheets.Add(After:=Worksheets("Table_Summary")).Name = "tempWBS"


For Each sht In ActiveWorkbook.Worksheets
If SheetExists("WBS") = True Then
Worksheets("WBS").Delete
End If
Next
Worksheets.Add(After:=Worksheets("tempwbs")).Name = "WBS"

Worksheets("WBS").Visible = True



'......................................load task List
i = 0
For Each sht In ActiveWorkbook.Worksheets
If SheetExists("task") = True Then
i = i + 1
End If
Next

If i = 0 Then
Exit Sub
End If


With Worksheets("tempWBS")
.Cells(1, 1) = "Parent_wbs_id_New"
.Cells(1, 2) = "wbs_id"
.Cells(1, 3) = "parent_wbs_id"
.Cells(1, 4) = "proj_node_flag"
.Cells(1, 5) = "wbs_name"
.Cells(1, 6) = "wbs_short_name"
.Cells(1, 7) = "seq_num"




strsql = "(SELECT cstr([wbs_id]),cstr([parent_wbs_id]),[proj_node_flag],[wbs_name], [wbs_short_name],val([seq_num]) FROM [PROJWBS$] order by [wbs_id] ) "

cnn.Open
rssql.Open strsql, cnn

.Cells(2, 2).CopyFromRecordset rssql
.Range("a2:a" & LastRow(Worksheets("tempWBS"))).Formula = "= if(d2=""y"",""root"",c2)"
End With


cnn.Close
Set cnn = Nothing
Set rssql = Nothing
End Sub


Sub CreateParentChildSheet()
'  http://stackoverflow.com/questions/9821545/how-to-build-parent-child-data-table-in-excel
  Dim Child() As String
  Dim ChildCrnt As String
  Dim InxChildCrnt As Long
  Dim InxChildMax As Long
  Dim InxParentCrnt As Long
  Dim LevelCrnt As Long
  Dim Parent() As Long
  Dim ParentName() As String
  Dim ParentNameCrnt As String
  Dim ParentSplit() As String
  Dim RowCrnt As Long
  Dim RowLast As Long
  Dim MyTimer  As Double
  Dim cell As Range
  Dim Temp() As Variant

  With Worksheets("tempWBS")
    RowLast = .Cells(Rows.Count, 1).End(xlUp).Row
    ' If row 1 contains column headings, if every child has one parent
    ' and the ultimate ancester is recorded as having a parent of "Root",
    ' there will be one child per row
    ReDim Child(1 To RowLast - 1)

    InxChildMax = 0
    For RowCrnt = 2 To RowLast
      ChildCrnt = .Cells(RowCrnt, 1).Value
      If LCase(ChildCrnt) <> "root" Then
        Call AddKeyToArray(Child, ChildCrnt, InxChildMax)
      End If
      ChildCrnt = .Cells(RowCrnt, 2).Value
      If LCase(ChildCrnt) <> "root" Then
        Call AddKeyToArray(Child, ChildCrnt, InxChildMax)
      End If
    Next

    ' If this is not true, one of the assumptions about the
    ' child-parent table is false
    Debug.Assert InxChildMax = UBound(Child)

  

    ' Child() now contains every child plus the root in
    ' ascending sequence.

    ' Record parent of each child
      ReDim Parent(1 To UBound(Child))
      For RowCrnt = 2 To RowLast
        If LCase(.Cells(RowCrnt, 1).Value) = "root" Then
          ' This child has no parent
          Parent(InxForKey(Child, .Cells(RowCrnt, 2).Value)) = 0
        Else
          ' Record parent for child
          Parent(InxForKey(Child, .Cells(RowCrnt, 2).Value)) = _
                           InxForKey(Child, .Cells(RowCrnt, 1).Value)
        End If
      Next

  End With

  ' Build parent chain for each child and store in ParentName
  ReDim ParentName(1 To UBound(Child))

  LevelMax = 1

  For InxChildCrnt = 1 To UBound(Child)
    ParentNameCrnt = Child(InxChildCrnt)
    InxParentCrnt = Parent(InxChildCrnt)
    LevelCrnt = 1
    Do While InxParentCrnt <> 0
      ParentNameCrnt = Child(InxParentCrnt) & "|" & ParentNameCrnt
      InxParentCrnt = Parent(InxParentCrnt)
      LevelCrnt = LevelCrnt + 1
    Loop
    ParentName(InxChildCrnt) = ParentNameCrnt
     

    If LevelCrnt > LevelMax Then
      LevelMax = LevelCrnt
    End If
  Next

 

  With Worksheets("WBS")
    
    
    ReDim Temp(UBound(Child) + 1, LevelMax)
    

    
    For InxChildCrnt = 1 To UBound(Child)
      ParentSplit = Split(ParentName(InxChildCrnt), "|")
      For InxParentCrnt = 0 To UBound(ParentSplit)
        Temp(InxChildCrnt, InxParentCrnt) = "'" & ParentSplit(InxParentCrnt)
        
         Next
        
     Application.StatusBar = "Progress: " & InxChildCrnt & " of " & UBound(Child) & ", " & Format(InxChildCrnt / UBound(Child), "0%")
     DoEvents
    Next
    

    
.Range("f1").Resize(UBound(Temp), UBound(Temp, 2)).Offset(0, 0) = Temp
    Application.StatusBar = "Transfer WBS to spreadsheet done : "
       DoEvents





   For LevelCrnt = 1 To LevelMax
      .Cells(1, LevelCrnt + 5) = LevelCrnt
    Next


  End With

End Sub

Sub AddKeyToArray(ByRef Tgt() As String, ByVal Key As String, _
                                                  ByRef InxTgtMax As Long)

  ' Add Key to Tgt if it is not already there.

  Dim InxTgtCrnt As Long

  For InxTgtCrnt = LBound(Tgt) To InxTgtMax
    If Tgt(InxTgtCrnt) = Key Then
      ' Key already in array
      Exit Sub
    End If
  Next
  ' If get here, Key has not been found
  InxTgtMax = InxTgtMax + 1
  If InxTgtMax <= UBound(Tgt) Then
    ' There is room for Key
    Tgt(InxTgtMax) = Key
  End If

End Sub

Function InxForKey(ByRef Tgt() As String, ByVal Key As String) As Long

  ' Return index entry for Key within Tgt

  Dim InxTgtCrnt As Long

  For InxTgtCrnt = LBound(Tgt) To UBound(Tgt)
    If Tgt(InxTgtCrnt) = Key Then
      InxForKey = InxTgtCrnt
      Exit Function
    End If
  Next

  Debug.Assert False        ' Error

End Function

Sub printwbs()
Dim lastr As Integer
Dim i As Integer
Dim c As Integer
Dim myarray() As Variant
Dim text As String




 lastr = LastRow(Worksheets("wbs"))

With Worksheets("wbs")
    .Activate
    
     .Sort.SortFields.Clear
    
    
    .Cells(1, 1) = "wbs_id"
    .Cells(1, 2) = "name"
    .Cells(1, 3) = "Level"
    .Cells(1, 4) = "Order"
    .Cells(1, 5) = "Path"
    
    For i = 1 To LevelMax
            
                 
            '.............................Lookup the order of the wbs_id .....................
            
            
            .Cells(1, 5 + LevelMax + i) = "order" & i
            
            
            .Range(Cells(2, LevelMax + 5 + i), Cells(lastr, LevelMax + 5 + i)).FormulaR1C1 = "=iferror(VLOOKUP(c" & CStr(5 + i) & ",tempWBS!c2:c7,6,0),0)"
    
            
            '.......................lookup the wbs discription..................................
         
            
            .Cells(1, 5 + LevelMax * 2 + i) = "level" & i
            
            
            .Range(Cells(2, LevelMax * 2 + 5 + i), Cells(lastr, LevelMax * 2 + 5 + i)).FormulaR1C1 = "=iferror(VLOOKUP(c" & CStr(5 + i) & ",tempWBS!c2:c5,4,0),"""")"
           
           
             '.......................lookup the wbs Path..................................
         
            
            .Cells(1, 5 + LevelMax * 3 + i) = "PATH" & i
            
            
            .Range(Cells(2, LevelMax * 3 + 5 + i), Cells(lastr, LevelMax * 3 + 5 + i)).FormulaR1C1 = "=iferror(VLOOKUP(c" & CStr(5 + i) & ",tempWBS!c2:c6,5,0),"""")"
           
           '**********************************************Sort *************************
            
            .Sort.SortFields.Add Key:=Range(.Cells(2, LevelMax + i + 5), .Cells(lastr, LevelMax + i + 5)), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
          '********************************************** wbs hiearchy *************************
    
            .Range(Cells(2, 6 + i - 1), Cells(lastr, 6 + i - 1)).Copy
            .Range(Cells(2, 1), Cells(lastr, 1)).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, Transpose:=False
         
         
         Application.StatusBar = "Lookup the WBS order " & i
      
            
    Next i
      .Range(Columns(6), Columns(6 + LevelMax * 4)).Hidden = True
        
    
    .Range("b2:b" & CStr(lastr)).FormulaR1C1 = "=VLOOKUP(c1,tempWBS!c2:c5,4,0)"
    
    .Range("c2:c" & CStr(lastr)).Formula = "=MATCH(A2,e2:AM2,0)-1"
    .Range("d2:d" & CStr(lastr)).Formula = "=row()-1"
    
    
    
    
    
    
        With .Sort
            .SetRange Range(Cells(1, 1), Cells(lastr, LevelMax * 3 + 4))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
    For i = 2 To lastr
    
    .Cells(i, 5) = MConcate(.Range(Cells(i, LevelMax * 3 + 6), Cells(i, LevelMax * 3 + 4 + Cells(i, 3) + 1)), ".")
    Application.StatusBar = "Generating WBS path: " & i & " of " & lastr & ", " & Format(i / lastr, "0%")

    Next i
    .Range(Cells(2, 1), Cells(lastr, LevelMax * 4 + 5)).Copy
    .Range(Cells(2, 1), Cells(lastr, LevelMax * 4 + 5)).PasteSpecial Paste:=xlPasteValues
    
             Application.StatusBar = "WBS done "

.Columns("d").NumberFormat = "#,##0"

End With
Application.DisplayAlerts = False
Worksheets("tempWBS").Delete

End Sub


