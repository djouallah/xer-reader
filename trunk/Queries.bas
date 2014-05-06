Attribute VB_Name = "Queries"
'  mimoune.djouallah@gmail.com
Option Explicit
Global Project_End_Date As Date
Global Project_Start_Date As Date
Global DataDate As Date
Global MULTI_PROJECT As Boolean

Sub GenerateDatasets()
#If VBA7 And Win64 Then
Dim i As LongLong   ' counter
Dim InTask As LongLong ' count the Number of Incomplted Tasks considered for the Analysis
Dim InLink As LongLong ' count the Number of Incomplted Tasks considered for the Analysis
Dim MP As LongLong  ' count the Number of Incomplted Tasks without predecessor
Dim MS As LongLong  ' count the Number of Incomplted Tasks without succesor
Dim NL As LongLong  ' count the Number of Incomplted Tasks with negative lag
Dim PL As LongLong  ' count the Number of Incomplted Tasks with positive lag
Dim rs As LongLong  ' count the Number of Incomplted Tasks without resources
Dim NF As LongLong  ' count the Number of Incomplted Tasks Negative Float
Dim HF As LongLong  ' count the Number of Incomplted Tasks with HIGH Float
Dim Hd As LongLong  ' count the Number of Incomplted Tasks with HIGH Duration
Dim ES As LongLong  ' count the Number of Incomplted Tasks with invalide Early Start
Dim EF As LongLong  ' count the Number of Incomplted Tasks with invalide Early Finish
Dim IAS As LongLong ' count the Number of Incomplted Tasks with invalide Actual Start
Dim AF As LongLong  ' count the Number of Incomplted Tasks with invalide Actual Finish
Dim HC As LongLong  ' count the number of incompleted tasks with hard constraints
Dim FS As LongLong ' count the number of FS links
Dim SS As LongLong ' count the number of SS links
Dim FF As LongLong ' count the number of FF links
Dim SF As LongLong ' count the number of SF links
#Else
Dim i As Integer   ' counter
Dim InTask As Long ' count the Number of Incomplted Tasks considered for the Analysis
Dim InLink As Long ' count the Number of Incomplted Tasks considered for the Analysis
Dim MP As Long  ' count the Number of Incomplted Tasks without predecessor
Dim MS As Long  ' count the Number of Incomplted Tasks without succesor
Dim NL As Long  ' count the Number of Incomplted Tasks with negative lag
Dim PL As Long  ' count the Number of Incomplted Tasks with positive lag
Dim rs As Long  ' count the Number of Incomplted Tasks without resources
Dim NF As Long  ' count the Number of Incomplted Tasks Negative Float
Dim HF As Long  ' count the Number of Incomplted Tasks with HIGH Float
Dim Hd As Long  ' count the Number of Incomplted Tasks with HIGH Duration
Dim ES As Long  ' count the Number of Incomplted Tasks with invalide Early Start
Dim EF As Long  ' count the Number of Incomplted Tasks with invalide Early Finish
Dim IAS As Long ' count the Number of Incomplted Tasks with invalide Actual Start
Dim AF As Long  ' count the Number of Incomplted Tasks with invalide Actual Finish
Dim HC As Long  ' count the number of incompleted tasks with hard constraints
Dim FS As Long ' count the number of FS links
Dim SS As Long ' count the number of SS links
Dim FF As Long ' count the number of FF links
Dim SF As Long ' count the number of SF links

#End If

Dim sht As Worksheet
Dim x  As Integer
Dim D As Integer
Dim cnn As Object
Dim rssql As Object
Dim strsql As String
Dim strsql1 As String




Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")
rssql.CursorLocation = 3

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"
    cnn.Open

' print Project Information

i = 0
For x = 2 To LastRow(Worksheets("Table_summary"))
If Worksheets("table_summary").Cells(x, 5).Value = "PROJECT" Then
i = i + 1
End If
Next x

If i = 1 Then
    With Worksheets("PROJECT")
    DataDate = CDate(.Cells(2, WorksheetFunction.Match("last_recalc_date", .Range("1:1"), 0)))
    End With
    
    cnn.Close
    strsql = "SELECT * FROM [PROJECT$] "
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    If rssql.RecordCount > 1 Then
        MsgBox "XER has multiple Project, analysis not supported.", vbInformation, APPNAME
        MULTI_PROJECT = True
    Exit Sub
     End If
     
End If

i = 0
' check if Task sheet is not empty
For x = 2 To LastRow(Worksheets("Table_summary"))
If Worksheets("table_summary").Cells(x, 5).Value = "TASK" Then
i = i + 1
End If
Next x

If i = 0 Then
MsgBox "No Task Found", vbInformation, APPNAME
Exit Sub
End If


With Worksheets("Dashboard")
    .Cells(4, 1) = "Project"
    .Cells(5, 1) = "Data Date"
    .Cells(5, 2).NumberFormat = "dd-mm-yyyy h:mm AM/PM"
    .Cells(6, 1) = "Project Start Date"
    .Cells(6, 2).NumberFormat = "dd-mm-yyyy h:mm AM/PM"
    .Cells(7, 1) = "Project End Date"
    .Cells(7, 2).NumberFormat = "dd-mm-yyyy h:mm AM/PM"
    
    cnn.Close
    strsql = "SELECT  [wbs_name]  FROM [PROJWBS$] WHERE [proj_node_flag] = 'Y'"
    cnn.Open
    
    rssql.Open strsql, cnn
    .Cells(4, 2).CopyFromRecordset rssql
    
    
    .Cells(5, 2) = DataDate
    
    cnn.Close
    strsql1 = "SELECT cdate([act_start_date]) AS [start] FROM [TASK$]  WHERE [status_code] = 'TK_Complete'" & _
    "UNION ALL " & _
    "SELECT cdate([early_start_date]) AS [start] FROM [TASK$]  WHERE [status_code] = 'TK_NotStart'" & _
    "UNION ALL " & _
    "SELECT cdate([act_start_date]) AS [start] FROM [TASK$]  WHERE [status_code] = 'TK_Active'" & _
    "ORDER BY [start]"
    
    
    strsql = " select Min([start]) FROM (" & strsql1 & ")"
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    .Cells(6, 2).CopyFromRecordset rssql
    
    
    Project_Start_Date = .Cells(6, 2)
   
    cnn.Close
    strsql1 = "SELECT cdate([act_end_date]) AS [finish] FROM [TASK$]  WHERE [status_code] = 'TK_Complete'" & _
    "UNION ALL " & _
    "SELECT cdate([early_end_date]) AS [finish] FROM [TASK$]  WHERE [status_code] = 'TK_NotStart'" & _
    "UNION ALL " & _
    "SELECT cdate([early_end_date]) AS [finish] FROM [TASK$]  WHERE [status_code] = 'TK_Active'" & _
    "ORDER BY [finish]"
    
    
    strsql = " select Max([finish]) FROM (" & strsql1 & ")"
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    .Cells(7, 2).CopyFromRecordset rssql
    
    Project_End_Date = .Cells(7, 2)
   
    
' Print tasks Statistics
    
    cnn.Close
    strsql = "SELECT [task_id] FROM [task$]"
    cnn.Open
    
    rssql.Open strsql, cnn
    
    
    .Cells(10, 2) = rssql.RecordCount
    cnn.Close
    strsql = "SELECT [clndr_id] FROM [task$] WHERE [task_type] = 'TT_WBS'"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    .Cells(11, 2) = rssql.RecordCount
    
    cnn.Close
    strsql = "SELECT [clndr_id] FROM [task$] WHERE [task_type] = 'TT_LOE'"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    .Cells(12, 2) = rssql.RecordCount
    cnn.Close
    strsql = "SELECT [clndr_id] FROM [task$] WHERE [task_type] = 'TT_Mile'"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
   .Cells(13, 2) = rssql.RecordCount
    
  
    cnn.Close
    strsql = "SELECT [clndr_id] FROM [task$] WHERE [task_type] = 'TT_FinMile'"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    .Cells(14, 2) = rssql.RecordCount
    cnn.Close
    strsql = "SELECT [clndr_id] FROM [task$] WHERE [task_type] = 'TT_Task'"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
   .Cells(15, 2) = rssql.RecordCount
   
    cnn.Close
    strsql = "SELECT [clndr_id] FROM [task$] WHERE [task_type] = 'TT_Rsrc'"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
   .Cells(16, 2) = rssql.RecordCount
      

   ' print Float Metrics
   
     cnn.Close
     strsql = "SELECT cdbl([total_float_hr_cnt]/[day_hr_cnt]) FROM [TASK$] INNER JOIN [CALENDAR$] ON [TASK$].[clndr_id] = [CALENDAR$].[clndr_id]" & _
    "WHERE [status_code] <> 'TK_Complete'"

    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    .Cells(2, 22).CopyFromRecordset rssql
     

    .Cells(10, 5) = Application.CountIf(.Range("V2:V" & LastRow(Worksheets("dashboard"))), "<0")
    .Cells(11, 5) = Application.CountIf(.Range("V2:V" & LastRow(Worksheets("dashboard"))), 0)
    .Cells(12, 5) = Application.CountIf(.Range("V2:V" & LastRow(Worksheets("dashboard"))), "<44") - Application.CountIf(.Range("V2:V" & LastRow(Worksheets("dashboard"))), 0) - Application.CountIf(.Range("V2:V" & LastRow(Worksheets("dashboard"))), "<0")
    .Cells(13, 5) = Application.CountIf(.Range("V2:V" & LastRow(Worksheets("dashboard"))), ">44")
    
    
    
    ' print Budget Metrics *****************************************
    '************* Budget
     cnn.Close
     strsql = "SELECT Sum([target_work_qty]) FROM [TASK$]"
   
     cnn.Open
     
    
     rssql.Open strsql, cnn
    
    .Cells(10, 8).CopyFromRecordset rssql
     

    '************* Actual
     cnn.Close
     strsql = "SELECT Sum([act_work_qty]) FROM [TASK$]"
   
     cnn.Open
     
    
     rssql.Open strsql, cnn
    
    .Cells(11, 8).CopyFromRecordset rssql
    cnn.Close
    '************* Remaining
     strsql = "SELECT Sum([remain_work_qty]) FROM [TASK$]"
   
     cnn.Open
     
    
     rssql.Open strsql, cnn
    
    .Cells(12, 8).CopyFromRecordset rssql
   
    '*************************************************************
     '************* Remaining
     .Cells(13, 8) = .Cells(12, 8) + .Cells(11, 8)
     '***********************************
   
   
   
   ' print Task Status *****************************************
    '************* Completed
       cnn.Close
       strsql = "SELECT [TASK_ID] FROM [TASK$] WHERE [status_code] = 'TK_Complete'"
       cnn.Open
       

       rssql.Open strsql, cnn

      .Cells(16, 5) = rssql.RecordCount
      
    '************* On going
     cnn.Close
     strsql = "SELECT [TASK_ID] FROM [TASK$] WHERE [status_code] = 'TK_Active'"
       cnn.Open
       

       rssql.Open strsql, cnn

      .Cells(17, 5) = rssql.RecordCount
     
       '************* Non Started
     cnn.Close
     strsql = "SELECT [TASK_ID] FROM [TASK$] WHERE [status_code] = 'TK_NotStart'"
       cnn.Open
       

       rssql.Open strsql, cnn

      .Cells(18, 5) = rssql.RecordCount
      
      
   .Rows("8:28").EntireRow.Hidden = False
End With






i = 0
For x = 2 To LastRow(Worksheets("Table_summary"))
If Worksheets("table_summary").Cells(x, 5).Value = "PROJECT" Then
i = i + 1
End If
Next x

If i = 1 Then
    With Worksheets("PROJECT")
    DataDate = CDate(.Cells(2, WorksheetFunction.Match("last_recalc_date", .Range("1:1"), 0)))
    End With
    cnn.Close
    strsql = "SELECT * FROM [PROJECT$] "
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    If rssql.RecordCount > 1 Then
        MsgBox "XER has multiple Project, analysis not supported.", vbInformation, APPNAME
    Exit Sub
     End If
     
End If

i = 0
' check if Task sheet is not empty
For x = 2 To LastRow(Worksheets("Table_summary"))
If Worksheets("table_summary").Cells(x, 5).Value = "TASK" Then
i = i + 1
End If
Next x

If i = 0 Then
MsgBox "No Task Found", vbInformation, APPNAME
Exit Sub
End If

 
For Each sht In ActiveWorkbook.Worksheets
    If SheetExists("Details_report") = True Then
       Worksheets("Details_report").Delete
    End If
Next


     Worksheets.Add(After:=Worksheets("Dashboard")).Name = "Details_report"
     
     With Worksheets("Details_report")
        .Cells(1, 1) = "Metrics"
        .Cells(1, 2) = "Internal ID"
        .Cells(1, 3) = "Activity ID"
        .Cells(1, 4) = "Activity Name "
        .Range("A1:D1").AutoFilter
        .Columns("A:D").AutoFit
        .Range("A1:D1").Interior.ColorIndex = 37
    End With







i = 2

'Metrics for Links
D = 0
For x = 2 To LastRow(Worksheets("Table_summary"))
If Worksheets("table_summary").Cells(x, 5) = "TASKPRED" Then

D = D + 1
End If
Next x

If D = 1 Then
cnn.Close
strsql = "SELECT [TASKPRED$].[TASK_ID] FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete'" & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open


rssql.Open strsql, cnn

InLink = rssql.RecordCount

Application.StatusBar = "Number of links calculated: "
 DoEvents

cnn.Close
strsql = "SELECT * FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete' AND [pred_type]= 'PR_FS'" & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open


rssql.Open strsql, cnn

FS = rssql.RecordCount

Application.StatusBar = "Number of FS calculated: "
 DoEvents
 
cnn.Close
strsql = "SELECT * FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete' AND [pred_type]= 'PR_SS'" & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open


rssql.Open strsql, cnn

SS = rssql.RecordCount
cnn.Close
Application.StatusBar = "Number of SS calculated: "
 DoEvents
 
 
 
strsql = "SELECT * FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete' AND [pred_type]= 'PR_FF'" & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open


rssql.Open strsql, cnn

FF = rssql.RecordCount
cnn.Close
 Application.StatusBar = "Number of FF calculated: "
 DoEvents
   
    
strsql = "SELECT * FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete' AND [pred_type]= 'PR_SF'" & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open


rssql.Open strsql, cnn

SF = rssql.RecordCount


Application.StatusBar = "Number of SF calculated: "
 DoEvents





   
cnn.Close
strsql = "SELECT distinct [TASKPRED$].[TASK_ID], [TASK$].[task_code],[TASK$].[task_name] FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete' AND iif(isnull([lag_hr_cnt]),0,clng([lag_hr_cnt])) < 0 " & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open

rssql.Open strsql, cnn

NL = rssql.RecordCount
If NL <> 0 Then
Worksheets("Details_report").Range("A" & i & ":" & "A" & CStr(NL + i - 1)) = " Negative Lag"
Worksheets("Details_report").Cells(i, 2).CopyFromRecordset rssql
i = i + NL
End If
cnn.Close

Application.StatusBar = "Number of Negative Lag calculated: "
 DoEvents

strsql = "SELECT distinct [TASKPRED$].[TASK_ID], [TASK$].[task_code],[TASK$].[task_name] FROM [TASKPRED$] INNER JOIN [TASK$] ON [TASKPRED$].[TASK_ID] = [TASK$].[TASK_ID]" & _
"WHERE [status_code] <> 'TK_Complete' AND iif(isnull([lag_hr_cnt]),0,clng([lag_hr_cnt])) > 0 " & _
"AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"

cnn.Open


rssql.Open strsql, cnn

PL = rssql.RecordCount

Application.StatusBar = "Number of Lag calculated: "
 DoEvents


If PL <> 0 Then
Worksheets("Details_report").Range("A" & i & ":" & "A" & CStr(PL + i - 1)) = " Have Lag"
Worksheets("Details_report").Cells(i, 2).CopyFromRecordset rssql
i = i + PL
End If
End If


If D = 0 Then
NL = 0
PL = 0
FS = 0
SS = 0
SF = 0
FF = 0
End If





' metrics for the tasks-------------------------------------------------------

With Worksheets("Details_report")

    cnn.Close
    strsql = "SELECT [task_id], [task_code],[task_name] FROM [task$] WHERE [status_code] <> 'TK_Complete' AND ( [task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    InTask = rssql.RecordCount
    
    
    Application.StatusBar = "Number of incompleted calculated: "
     DoEvents
    
    If InTask <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(InTask + i - 1)) = " Included Tasks"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + InTask
    End If
    
    
    D = 0
    For x = 2 To LastRow(Worksheets("Table_summary"))
    If Worksheets("table_summary").Cells(x, 5) = "TASKPRED" Then
    
    D = D + 1
    End If
    Next x
    
    If D = 1 Then
    cnn.Close
    strsql = "SELECT distinct [TASK$].[TASK_ID],[task_code],[task_name] FROM [TASK$] LEFT JOIN [TASKPRED$] ON [TASK$].[TASK_ID] = [TASKPRED$].[TASK_ID]" & _
    "WHERE [TASKPRED$].[TASK_ID] is Null AND [status_code] <> 'TK_Complete'" & _
    "AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    
    
    MP = rssql.RecordCount
    
    
    Application.StatusBar = "Number of missing Predecessor: "
     DoEvents
    
    If MP <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(MP + i - 1)) = "Missing Predecessor"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + MS
    End If
    End If
    
    If D = 0 Then
    cnn.Close
    strsql = "SELECT [task_id], [task_code],[task_name] FROM [task$] WHERE [status_code] <> 'TK_Complete' AND ( [task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    InTask = rssql.RecordCount
    MP = InTask
    
    
    If InTask <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(InTask + i - 1)) = " Missing Predecessor"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + InTask
    End If
    End If
    
    
    D = 0
    For x = 2 To LastRow(Worksheets("Table_summary"))
    If Worksheets("table_summary").Cells(x, 5) = "TASKPRED" Then
    
    D = D + 1
    End If
    Next x
    
    If D = 1 Then
    cnn.Close
    strsql = "SELECT distinct [TASK$].[TASK_ID],[task_code],[task_name] FROM [TASK$] LEFT JOIN [TASKPRED$] ON [TASK$].[TASK_ID] = [TASKPRED$].[pred_task_id]" & _
    "WHERE [TASKPRED$].[pred_task_id] is Null AND [status_code] <> 'TK_Complete'" & _
    "AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    
    MS = rssql.RecordCount
    
    
    Application.StatusBar = "Number of missing Successor: "
     DoEvents
    
    If MS <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(MS + i - 1)) = "Missing Successor"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + MS
    End If
    End If
    
    If D = 0 Then
    cnn.Close
    strsql = "SELECT [task_id], [task_code],[task_name] FROM [task$] WHERE [status_code] <> 'TK_Complete' AND ( [task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    MS = rssql.RecordCount
    If MS <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(InTask + i - 1)) = " Missing Successor"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + InTask
    End If
    End If
    
     cnn.Close
     strsql = "SELECT  [task_id], [task_code],[task_name] FROM [task$] WHERE [status_code] <> 'TK_Complete'" & _
     "AND ( [task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')" & _
     "AND ( [cstr_type] = 'CS_MANDFIN' OR [cstr_type] = 'CS_MANDSTART' OR [cstr_type2] = 'CS_MANDFIN' OR [cstr_type2] = 'CS_MANDSTART')"
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    HC = rssql.RecordCount
    
    If HC <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(HC + i - 1)) = " Hard Constraints"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + HC
    End If
    
    
    
    Application.StatusBar = "Number of Hard Constraints: "
     DoEvents
    
    
    cnn.Close
    strsql = "SELECT [TASK$].[task_id] , [task_code],[task_name]  FROM [TASK$] INNER JOIN [CALENDAR$] ON [TASK$].[clndr_id] = [CALENDAR$].[clndr_id]" & _
    "WHERE [status_code] <> 'TK_Complete'" & _
    "AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')" & _
    "AND CDBL([target_drtn_hr_cnt] / [day_hr_cnt])> 44 "
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    
    Hd = rssql.RecordCount
    
    If Hd <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(Hd + i - 1)) = " high duration"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + Hd
    End If
    
    
    Application.StatusBar = "Number of High Duration: "
     DoEvents
    
    cnn.Close
    strsql = "SELECT [TASK$].[task_id] , [task_code],[task_name]  FROM [TASK$] INNER JOIN [CALENDAR$] ON [TASK$].[clndr_id] = [CALENDAR$].[clndr_id]" & _
    "WHERE [status_code] <> 'TK_Complete'" & _
    "AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')" & _
    "AND CDBL([total_float_hr_cnt] / [day_hr_cnt])> 44 "
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    
    HF = rssql.RecordCount
    
    If HF <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(HF + i - 1)) = " high Float"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + Hd
    End If
    
    
    Application.StatusBar = "Number of High Float: "
     DoEvents
    
    cnn.Close
    strsql = "SELECT  [task_id] , [task_code],[task_name] FROM" & _
    "(select *, IIf(IsNull([total_float_hr_cnt]),0,cdbl(val([total_float_hr_cnt]))) AS [New] from [task$])" & _
    "WHERE [status_code] <> 'TK_Complete' AND [new] < 0  AND ( [task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    cnn.Open
    
    
    
    rssql.Open strsql, cnn
    
    NF = rssql.RecordCount
    
    
    If NF <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(NF + i - 1)) = " Neagtive Float"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + NF
    End If
    cnn.Close
    strsql = "SELECT  [task_id] , [task_code],[task_name] FROM [task$] WHERE ([status_code] = 'TK_Complete'  AND " & CLng(DataDate) & " < clng([act_start_date]))OR ( [status_code] = 'TK_Active' AND " & CLng(DataDate) & " < clng([act_start_date]))"
    strsql1 = strsql
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    IAS = rssql.RecordCount
    
    
    Application.StatusBar = "Number of Invalide Actual Start: "
     DoEvents
    
    
    If IAS <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(IAS + i - 1)) = " Invalide Actual Start"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + IAS
    End If
    
    cnn.Close
    strsql = "SELECT  [task_id] , [task_code],[task_name] FROM [task$] WHERE [status_code] = 'TK_Complete' AND " & CLng(DataDate) & " < clng([act_end_date])"
    strsql1 = strsql
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    AF = rssql.RecordCount
      
    If AF <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(AF + i - 1)) = " Invalide Actual Finish"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + AF
    End If
     
     
     
    Application.StatusBar = "Invalide Actual Finish: "
     DoEvents
     cnn.Close
    strsql = "SELECT  [task_id] , [task_code],[task_name] FROM [task$] WHERE ([status_code] = 'TK_Active'  AND " & CLng(DataDate) & " > clng([early_start_date])) OR ([status_code] = 'TK_NotStart'  AND " & CLng(DataDate) & " > clng([early_start_date]))"
    strsql1 = strsql
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    ES = rssql.RecordCount
     
    If ES <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(ES + i - 1)) = " Invalide Forecast Start"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + ES
    End If
          
          
          
    Application.StatusBar = "Number of Invalide Forecast Start: "
     DoEvents
    cnn.Close
    strsql = "SELECT  [task_id] , [task_code],[task_name] FROM [task$] WHERE ([status_code] = 'TK_Active'  AND " & CLng(DataDate) & " > clng([early_end_date])) OR ([status_code] = 'TK_NotStart'  AND " & CLng(DataDate) & " > clng([early_end_date]))"
    strsql1 = strsql
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    EF = rssql.RecordCount
         
    If EF <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(EF + i - 1)) = " Invalide Forecast Finish"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + EF
    End If
         
         
     Application.StatusBar = "Number of Invalide Forecast Finish: "
     DoEvents
        
    D = 0
    For x = 2 To LastRow(Worksheets("Table_summary"))
    If Worksheets("table_summary").Cells(x, 5) = "TASKRSRC" Then
    D = D + 1
    End If
    Next x
    
    If D = 1 Then
     cnn.Close
    strsql = "SELECT distinct [TASK$].[TASK_ID],[task_code],[task_name] FROM [TASK$] LEFT JOIN [TASKRSRC$] ON [TASK$].[TASK_ID] = [TASKRSRC$].[TASK_ID]" & _
    "WHERE [TASKRSRC$].[TASK_ID] is Null AND [status_code] <> 'TK_Complete'" & _
    "AND ([task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    rs = rssql.RecordCount
    
    If rs <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(rs + i - 1)) = " Tasks without resources"
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + rs
    End If
    End If
    
    
    Application.StatusBar = "Number of Tasks without Resources: "
     DoEvents
    
    
    If D = 0 Then
    cnn.Close
    strsql = "SELECT [task_id], [task_code],[task_name] FROM [task$] WHERE [status_code] <> 'TK_Complete' AND ( [task_type] = 'TT_Task' OR [task_type] = 'TT_Rsrc')"
    cnn.Open
    
    
    rssql.Open strsql, cnn
    
    InTask = rssql.RecordCount
    rs = InTask
    If InTask <> 0 Then
    .Range("A" & i & ":" & "A" & CStr(InTask + i - 1)) = "Tasks without resources "
    .Cells(i, 2).CopyFromRecordset rssql
    i = i + InTask
    End If
    End If



    .Columns("A").AutoFit
    .Columns("c:d").AutoFit
    .Columns("b").Hidden = True
   
End With

'--------------------- print the report ----------------------------------------



' print the result in Dashboard

With Worksheets("dashboard")



' Print Metrics
    .Cells(25, 1) = InLink
    .Cells(25, 2) = InTask
    .Cells(24, 3) = MP
    .Cells(24, 4) = MS
    .Cells(24, 5) = NL
    .Cells(24, 6) = PL
    .Cells(24, 7) = FS
    .Cells(24, 8) = SS + FF
    .Cells(24, 9) = SF
    .Cells(24, 10) = HC
    .Cells(24, 11) = HF
    .Cells(24, 12) = NF
    .Cells(24, 13) = Hd
    .Cells(24, 14) = IAS
    .Cells(24, 15) = AF
    .Cells(24, 16) = ES
    .Cells(24, 17) = EF
    .Cells(24, 18) = rs
    If InTask <> 0 Then
        .Cells(25, 3) = CDbl(MP / InTask)
        .Cells(25, 4) = CDbl(MS / InTask)
        .Cells(25, 5) = CDbl(NL / InTask)
        .Cells(25, 6) = CDbl(PL / InTask)
        .Cells(25, 10) = CDbl(HC / InTask)
        .Cells(25, 11) = CDbl(HF / InTask)
        .Cells(25, 12) = CDbl(NF / InTask)
        .Cells(25, 13) = CDbl(Hd / InTask)
        .Cells(25, 14) = CDbl(IAS / InTask)
        .Cells(25, 15) = CDbl(AF / InTask)
        .Cells(25, 16) = CDbl(ES / InTask)
        .Cells(25, 17) = CDbl(EF / InTask)
        .Cells(25, 18) = CDbl(rs / InTask)
    End If
    
    If InLink <> 0 Then
        .Cells(25, 7) = CDbl(FS / InLink)
        .Cells(25, 8) = CDbl((SS + FF) / (2 * InLink))
        .Cells(25, 9) = CDbl(SF / InLink)
    End If
    .Activate
End With


Application.StatusBar = ""
 DoEvents
 
cnn.Close
Set cnn = Nothing
Set rssql = Nothing

End Sub

Sub Task_list()
Dim sht As Worksheet
Dim lastr As Long
Dim tempa As Variant
Dim NT As Integer ' number of tasks
Dim i As Integer

Dim cnn As Object
Dim rssql As Object
Dim strsql As String

Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"




For Each sht In ActiveWorkbook.Worksheets
If SheetExists("Task_list") = True Then
Worksheets("Task_list").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "Task_list"





With Worksheets("Task_list")

.Columns("n:p").NumberFormat = "d/m/yyyy;@"

.Activate
.Cells(1, 1) = "task_id"
.Cells(1, 2) = "task"
.Cells(1, 3) = "Order"
.Cells(1, 4) = "Level"
.Cells(1, 5) = "wbs_id"
.Cells(1, 6) = "PATH"
.Cells(1, 7) = "Activity ID"
.Cells(1, 8) = "Activity Name "
.Cells(1, 9) = "Budget_labor"
.Cells(1, 10) = "Actual_Labor"
.Cells(1, 11) = "Remaining"
.Cells(1, 12) = "At_Completion"
.Cells(1, 13) = "od "
.Cells(1, 14) = "Start "
.Cells(1, 15) = "restart "
.Cells(1, 16) = "Finish "
.Cells(1, 17) = "Total_Float "
.Cells(1, 18) = "Duration_Pct"
.Cells(1, 19) = "Activity_Pct"
.Cells(1, 20) = "Earned_Value"


strsql = "SELECT [task_id],1,[Order],'', [TASK$].[wbs_id],[Path],[task_code],[task_name],val([target_work_qty])," & _
"val([act_work_qty]),val([remain_work_qty]),([act_work_qty]+[remain_work_qty]),iif(val([target_drtn_hr_cnt])=0,0.0000001,val([target_drtn_hr_cnt]))," & _
"cdate([act_start_date]) AS [start], cdbl([restart_date]) as [restart],cdate([early_end_date]) AS [finish]," & _
"(([total_float_hr_cnt])/[day_hr_cnt]) AS [TF_days]," & _
"iif([target_drtn_hr_cnt]=0,0,iif(cdbl([target_drtn_hr_cnt])<cdbl([remain_drtn_hr_cnt]),0,cdbl(cdbl([target_drtn_hr_cnt]-[remain_drtn_hr_cnt])/cdbl([target_drtn_hr_cnt])))) as [D_percent]," & _
"iif([complete_pct_type]='CP_Drtn',[D_percent],iif([complete_pct_type]='CP_Phys',val([phys_complete_pct])/100,iif([remain_work_qty]+[act_work_qty]+[remain_equip_qty]+[act_equip_qty]=0,0,(([act_work_qty]+[act_equip_qty])/([remain_work_qty]+[act_work_qty]+[remain_equip_qty]+[act_equip_qty]))))) as [percent] " & _
" ,([act_work_qty]+[remain_work_qty])* [percent] FROM  ([TASK$] INNER JOIN [CALENDAR$] ON [TASK$].[clndr_id] = [CALENDAR$].[clndr_id]) " & _
"INNER JOIN [WBS$] on [TASK$].[wbs_id] = [WBS$].[wbs_id] WHERE [status_code] = 'TK_Active'" & _
"UNION ALL " & _
"SELECT [task_id],1,[Order],'',[TASK$].[wbs_id],[Path],[task_code],[task_name],val([target_work_qty]),val([act_work_qty]),val([remain_work_qty]),([act_work_qty]+[remain_work_qty]),iif(val([target_drtn_hr_cnt])=0,0.0000001,val([target_drtn_hr_cnt])) ," & _
"cdate([early_start_date]) AS [start],cdbl([restart_date]) as [restart],([early_end_date]) AS [finish],(([total_float_hr_cnt])/[day_hr_cnt]) AS [TF_days], 0 ,0,0 FROM  ([TASK$] INNER JOIN [CALENDAR$] ON [TASK$].[clndr_id] = [CALENDAR$].[clndr_id]) INNER JOIN [WBS$] on [TASK$].[wbs_id] = [WBS$].[wbs_id]  WHERE [status_code] = 'TK_NotStart'" & _
"UNION ALL " & _
"SELECT [task_id],1,[Order],'', [TASK$].[wbs_id],[Path],[task_code],[task_name]," & _
"val([target_work_qty]),val([act_work_qty]),val([remain_work_qty]),([act_work_qty]+[remain_work_qty]),iif(val([target_drtn_hr_cnt])=0,0.0000001,val([target_drtn_hr_cnt]))," & _
"cdate([act_start_date]) AS [start],'' as [restart]," & _
"cdate([act_end_date]) AS [finish],(([total_float_hr_cnt])/[day_hr_cnt]) AS [TF_days]," & _
"1 ,iif([complete_pct_type]='CP_Drtn',1,iif([complete_pct_type]='CP_Phys',val([phys_complete_pct])/100,iif([remain_work_qty]+[act_work_qty]+[remain_equip_qty]+[act_equip_qty]=0,0,(([act_work_qty]+[act_equip_qty])/([remain_work_qty]+[act_work_qty]+[remain_equip_qty]+[act_equip_qty]))))) as [percent]," & _
"([act_work_qty]+[remain_work_qty]) FROM ([TASK$] INNER JOIN [CALENDAR$] ON [TASK$].[clndr_id] = [CALENDAR$].[clndr_id]) INNER JOIN [WBS$] on [TASK$].[wbs_id] = [WBS$].[wbs_id] WHERE [status_code] = 'TK_Complete'"

'#" & Format(DataDate, "Medium Date") & "#

cnn.Open
rssql.Open strsql, cnn
.Cells(2, 1).CopyFromRecordset rssql



cnn.Close
Set cnn = Nothing
Set rssql = Nothing


'******** for some reason order is sometimes returned as text, Jet is weird :(
  Columns("c:c").Select
    Selection.TextToColumns Destination:=Range("c1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True




  
  


Worksheets("Task_list").Visible = False


End With
End Sub


Sub TASKRSRC_SUM()
Dim sht As Worksheet

Dim cnn As Object
Dim rssql As Object
Dim strsql As String


Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"

For Each sht In ActiveWorkbook.Worksheets
If SheetExists("TASKRSRC_SUM") = True Then
Worksheets("TASKRSRC_SUM").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "TASKRSRC_SUM"


With Worksheets("TASKRSRC_SUM")
.Cells(1, 1) = "task_id"
.Cells(1, 2) = "target_cost"
.Cells(1, 3) = "Actual_cost"
.Cells(1, 4) = "remain_cost"
.Cells(1, 5) = "At_Completion_cost"
.Cells(1, 6) = "target_qty"
.Cells(1, 7) = "act_qty"
.Cells(1, 8) = "remain_qty"


strsql = "SELECT [task_id],sum([target_cost]) AS [Sum_target_cost],sum([act_ot_cost]+ [act_reg_cost]) AS [Sum_Act_cost],sum([remain_cost]) AS [Sum_Remaining_cost],[Sum_Act_cost]+[Sum_Remaining_cost], sum([target_qty])," & _
"sum([act_ot_qty]+[act_reg_qty]),Sum([remain_qty])" & _
"from [TASKRSRC$] GROUP BY [task_id]"



cnn.Open
rssql.Open strsql, cnn
.Cells(2, 1).CopyFromRecordset rssql


.Columns("b:e").NumberFormat = "#,##0.00"


cnn.Close
Set cnn = Nothing
Set rssql = Nothing
Worksheets("TASKRSRC_SUM").Visible = False
End With

End Sub
Sub Task_list_cost0()
Dim sht As Worksheet
Dim i As Integer

Dim cnn As Object
Dim rssql As Object
Dim strsql As String
Dim strsql1 As String
Dim strsql2 As String
Dim strsql3 As String
Dim strsql4 As String
Dim strsql5 As String

Dim def_cost_per_qty As Double


Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"

For Each sht In ActiveWorkbook.Worksheets
If SheetExists("Task_list_cost0") = True Then
Worksheets("Task_list_cost0").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "Task_list_cost0"



Application.StatusBar = "Group by WBS "



With Worksheets("Task_list_cost0")
.Cells(1, 1) = "task_id"
.Cells(1, 2) = "target_cost"
.Cells(1, 3) = "Actual_cost"
.Cells(1, 4) = "remain_cost"
.Cells(1, 5) = "At_Completion_cost"



strsql1 = "SELECT [def_cost_per_qty] from [PROJECT$]"
cnn.Open
rssql.Open strsql1, cnn
def_cost_per_qty = rssql.fields(0)


cnn.Close

i = 0
For Each sht In ActiveWorkbook.Worksheets
If SheetExists("TASKRSRC_SUM") = True Then
i = i + 1
End If
Next

If i = 0 Then

strsql3 = "SELECT [task_id], sum(([target_work_qty]+[target_equip_qty])* " & def_cost_per_qty & " ) AS [Sum_target_cost]," & _
"sum(([act_work_qty]+[act_equip_qty])* " & def_cost_per_qty & " ) AS [Sum_Act_cost], " & _
"sum(([remain_work_qty]+[remain_equip_qty])* " & def_cost_per_qty & " ) AS [Sum_Remaining_cost], " & _
"[Sum_Act_cost]+[Sum_Remaining_cost] " & _
"from [TASK$] GROUP BY [TASK$].[task_id]"

strsql = strsql3

Else


strsql2 = "SELECT [task_id],[target_cost],[Actual_cost] ,[remain_cost], [At_Completion_cost] from [TASKRSRC_SUM$]"


strsql3 = "SELECT [TASK$].[task_id], sum(([target_work_qty]+[target_equip_qty])* " & def_cost_per_qty & " ) AS [Sum_target_cost]," & _
"sum(([act_work_qty]+[act_equip_qty])* " & def_cost_per_qty & " ) AS [Sum_Act_cost], " & _
"sum(([remain_work_qty]+[remain_equip_qty])* " & def_cost_per_qty & " ) AS [Sum_Remaining_cost], " & _
"[Sum_Act_cost]+[Sum_Remaining_cost] " & _
"from [TASK$] LEFT JOIN [TASKRSRC$] ON [TASK$].[task_id] = [TASKRSRC$].[task_id] WHERE isnull([TASKRSRC$].[task_id]) " & _
"GROUP BY [TASK$].[task_id]"

strsql4 = "SELECT   [TASK$].[task_id], ([target_work_qty]+[target_equip_qty]-[target_qty]) * " & def_cost_per_qty & "   AS [Sum_target_cost], " & _
"([act_work_qty]+[act_equip_qty]-[act_qty])* " & def_cost_per_qty & "  AS [Sum_Act_cost], " & _
"([remain_work_qty]+[remain_equip_qty]-[remain_qty])* " & def_cost_per_qty & "  AS [Sum_Remaining_cost], " & _
"[Sum_Act_cost]+[Sum_Remaining_cost] " & _
"from [TASK$] INNER JOIN [TASKRSRC_SUM$] ON [TASK$].[task_id] = [TASKRSRC_SUM$].[task_id] "

strsql = strsql2 & " UNION ALL " & strsql3 & " UNION ALL " & strsql4
End If



i = 0
For Each sht In ActiveWorkbook.Worksheets
If SheetExists("PROJCOST") = True Then
i = i + 1
End If
Next

If i = 0 Then
strsql = strsql

Else
strsql5 = "SELECT [task_id],[target_cost],[act_cost] ,[remain_cost], [act_cost] + [remain_cost] from [PROJCOST$]"

strsql = strsql & " UNION ALL " & strsql5
End If



cnn.Open
rssql.Open strsql, cnn
.Cells(2, 1).CopyFromRecordset rssql


.Columns("b:e").NumberFormat = "#,##0.00"


cnn.Close
Set cnn = Nothing
Set rssql = Nothing
Worksheets("Task_list_cost0").Visible = False
End With
Application.StatusBar = "Lookup Cost"

End Sub

Sub Task_list_cost()

Dim cnn As Object
Dim rssql As Object
Dim strsql As String
Dim sht As Worksheet



Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"

For Each sht In ActiveWorkbook.Worksheets
If SheetExists("Task_list_cost") = True Then
Worksheets("Task_list_cost").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "Task_list_cost"



With Worksheets("Task_list_cost")

.Cells(1, 1) = "task_id"
.Cells(1, 2) = "target_cost"
.Cells(1, 3) = "Actual_cost"
.Cells(1, 4) = "remain_cost"
.Cells(1, 5) = "At_Completion_cost"




strsql = "SELECT [task_id], sum([target_cost]), sum([Actual_cost]), sum([remain_cost]), sum([At_Completion_cost]) " & _
"from [Task_list_cost0$] group by [task_id] "
cnn.Open
rssql.Open strsql, cnn

.Cells(2, 1).CopyFromRecordset rssql


.Columns("t:w").NumberFormat = "#,##0.00"


cnn.Close
strsql = "SELECT  sum([target_cost]), sum([Actual_cost]), sum([remain_cost]), sum([At_Completion_cost]) " & _
"from [Task_list_cost0$] "
cnn.Open
rssql.Open strsql, cnn

Worksheets("dashboard").Cells(10, 12).CopyFromRecordset rssql

    Worksheets("dashboard").Range("L10:O10").Copy
    Worksheets("dashboard").Range("K10").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Worksheets("dashboard").Range("L10:O10").ClearContents
    Worksheets("dashboard").Range("k10:k14").NumberFormat = "#,##0"
    
cnn.Close
Set cnn = Nothing
Set rssql = Nothing
Worksheets("Task_list_cost").Visible = False
End With
Application.StatusBar = "Sum Cost"

End Sub




Sub Task_list_Merge()

Dim cnn As Object
Dim rssql As Object
Dim strsql As String
Dim sht As Worksheet



Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"

For Each sht In ActiveWorkbook.Worksheets
If SheetExists("Task_list_Merge") = True Then
Worksheets("Task_list_Merge").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "Task_list_Merge"



With Worksheets("Task_list_Merge")

.Cells(1, 1) = "task_id"
.Cells(1, 2) = "task"
.Cells(1, 3) = "Order"
.Cells(1, 4) = "Level"
.Cells(1, 5) = "wbs_id"
.Cells(1, 6) = "PATH"
.Cells(1, 7) = "Activity ID"
.Cells(1, 8) = "Activity Name"
.Cells(1, 9) = "Budget_labor"
.Cells(1, 10) = "Actual_Labor"
.Cells(1, 11) = "Remaining"
.Cells(1, 12) = "At_Completion"
.Cells(1, 13) = "od "
.Cells(1, 14) = "Start "
.Cells(1, 15) = "restart "
.Cells(1, 16) = "Finish "
.Cells(1, 17) = "Total_Float "
.Cells(1, 18) = "Duration_Pct"
.Cells(1, 19) = "Activity_Pct"
.Cells(1, 20) = "Labor_Earned_Value"
.Cells(1, 21) = "target_cost"
.Cells(1, 22) = "Actual_cost"
.Cells(1, 23) = "remain_cost"
.Cells(1, 24) = "At_Completion_cost"





strsql = "SELECT [Task_list$].[task_id],[task],[Order],[Level],[wbs_id],[PATH],[Activity ID],[Activity Name]," & _
"[Budget_labor],[Actual_Labor],[Remaining],[At_Completion],[od],[Start],iif([restart]='','',val([restart])),[Finish],[Total_Float]," & _
"[Duration_Pct],[Activity_Pct],[Earned_Value],[Task_list_cost$].[target_cost],[Task_list_cost$].[Actual_cost],[Task_list_cost$].[remain_cost],[Task_list_cost$].[At_Completion_cost]" & _
"from [Task_list$] LEFT JOIN [Task_list_cost$] on [Task_list$].[task_id] = [Task_list_cost$].[task_id]"
    
    cnn.Open
    rssql.Open strsql, cnn
    .Cells(2, 1).CopyFromRecordset rssql

.Cells(2, 1).CopyFromRecordset rssql




cnn.Close
Set cnn = Nothing
Set rssql = Nothing
Worksheets("Task_list_Merge").Visible = False
End With
Application.StatusBar = "Task Cost List Done"

End Sub

Sub WBS_list()
Dim sht As Worksheet

Dim cnn As Object
Dim rssql As Object
Dim strsql As String

Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"

For Each sht In ActiveWorkbook.Worksheets
If SheetExists("WBS_List") = True Then
Worksheets("WBS_List").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "WBS_List"



Application.StatusBar = "Group by WBS "

With Worksheets("WBS_List")

.Columns("b").NumberFormat = "#,##0"
.Columns("o").NumberFormat = "d/m/yyyy;@"

.Cells(1, 1) = "task_id"
.Cells(1, 2) = "task"
.Cells(1, 3) = "Order"
.Cells(1, 4) = "Level"
.Cells(1, 5) = "wbs_id"
.Cells(1, 6) = "PATH"
.Cells(1, 7) = "name"
.Cells(1, 8) = "Activity Name"
.Cells(1, 9) = "Budget_labor"
.Cells(1, 10) = "Actual_Labor"
.Cells(1, 11) = "Remaining"
.Cells(1, 12) = "At_Completion"
.Cells(1, 13) = "od "
.Cells(1, 14) = "Start "
.Cells(1, 15) = "restart "
.Cells(1, 16) = "Finish "
.Cells(1, 17) = "Total_Float "
.Cells(1, 18) = "Duration_Pct"
.Cells(1, 19) = "Activity_Pct"
.Cells(1, 20) = "Labor_Earned_Value"
.Cells(1, 21) = "target_cost"
.Cells(1, 22) = "Actual_cost"
.Cells(1, 23) = "remain_cost"
.Cells(1, 24) = "At_Completion_cost"


strsql = "SELECT 'NA',0,val([WBS$].[Order]),[WBS$].[Level],[WBS$].[wbs_id],[WBS$].[Path],[WBS$].[name],'',sum([Budget_labor]), sum([Actual_Labor]),sum([Remaining]),sum([At_Completion])" & _
",sum([od]) as [T_OD] ,min([Start])as [M_Start],min([restart]) as [M_restart],max([Finish]) as [M_Finish],min([Total_Float])," & _
"iif([M_Finish]-[M_start]=0,0,iif([M_Finish]<val([M_restart]),1,iif([M_Start]>=val([M_restart]),0,([M_Finish]-val([M_restart]))/([M_Finish]-[M_Start]))))," & _
"iif([T_AC]=0,0,sum([At_Completion_cost]*[Activity_Pct])/[T_AC]) AS [A_Pct], sum([Labor_Earned_Value]) ,sum([target_cost]),sum([Actual_cost]),sum([remain_cost]),sum([At_Completion_cost])as [T_AC] from [wbs$] " & _
"inner JOIN [Task_list_Merge$] ON [Task_list_Merge$].[wbs_id] =  [wbs$].[wbs_id] GROUP BY [WBS$].[Path],[WBS$].[Order],[WBS$].[Level],[WBS$].[wbs_id],[WBS$].[name]"

cnn.Open
rssql.Open strsql, cnn
.Cells(2, 1).CopyFromRecordset rssql




cnn.Close
Set cnn = Nothing
Set rssql = Nothing
Worksheets("WBS_List").Visible = False
End With
Application.StatusBar = "Completed"

End Sub

Sub WBS_list_Parent()
Dim sht As Worksheet
Dim cnn As Object
Dim rssql As Object
Dim strsql As String

Set cnn = CreateObject("ADODB.Connection")
Set rssql = CreateObject("ADODB.Recordset")

    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & ";"



For Each sht In ActiveWorkbook.Worksheets
If SheetExists("WBS_List_Parent") = True Then
Worksheets("WBS_List_Parent").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "WBS_List_Parent"



With Worksheets("WBS_List_Parent")

.Cells(1, 1) = "task_id"
.Cells(1, 2) = "task"
.Cells(1, 3) = "Order"
.Cells(1, 4) = "Level"
.Cells(1, 5) = "wbs_id"
.Cells(1, 6) = "PATH"
.Cells(1, 7) = "name"
.Cells(1, 8) = "Activity Name"
.Cells(1, 9) = "Budget_labor"
.Cells(1, 10) = "Actual_Labor"
.Cells(1, 11) = "Remaining"
.Cells(1, 12) = "At_Completion"
.Cells(1, 13) = "od "
.Cells(1, 14) = "Start "
.Cells(1, 15) = "restart "
.Cells(1, 16) = "Finish "
.Cells(1, 17) = "Total_Float "
.Cells(1, 18) = "Duration_Pct"
.Cells(1, 19) = "Activity_Pct"
.Cells(1, 20) = "Earned_Value"
.Cells(1, 21) = "target_cost"
.Cells(1, 22) = "Actual_cost"
.Cells(1, 23) = "remain_cost"
.Cells(1, 24) = "At_Completion_cost"

    
    strsql = "SELECT 'NA',0,val([WBS$].[Order]),[WBS$].[Level],[WBS$].[wbs_id],[WBS$].[Path],[WBS$].[name],'',sum([Budget_labor]), sum([Actual_Labor]),sum([Remaining]),sum([At_Completion])" & _
    ",sum([od]) as [T_OD] ,min([Start])as [M_Start],min([restart])as [M_restart],max([Finish]) as [M_Finish],min([Total_Float])," & _
    "iif([M_Finish]-[M_start]=0,0,iif([M_Finish]< val([M_restart]),1,iif([M_Start]>= val([M_restart]),0,([M_Finish]- val([M_restart]))/([M_Finish]-[M_Start]))))," & _
    "iif([T_AC]=0,0,sum([At_Completion_cost]*[Activity_Pct])/[T_AC]) AS [A_Pct], sum([Labor_Earned_Value]) ,sum([target_cost]),sum([Actual_cost]),sum([remain_cost]),sum([At_Completion_cost]) as [T_AC] from [wbs$]" & _
    "inner JOIN [WBS_List$] ON [WBS_List$].[Path] like [wbs$].[Path] & '.%'  GROUP BY [WBS$].[Path],[WBS$].[Order],[WBS$].[Level],[WBS$].[wbs_id],[WBS$].[name]"
    
    cnn.Open
    rssql.Open strsql, cnn
    .Cells(2, 1).CopyFromRecordset rssql
    
    
    .Columns("b").NumberFormat = "#,##0"
    
    
    cnn.Close
    Set cnn = Nothing
    Set rssql = Nothing
  
    

End With

Worksheets("WBS_List_Parent").Visible = False

Application.StatusBar = "Sum Value at Parent wbs"

End Sub
Sub WBS_List_Merge()
Dim sht As Worksheet
Dim lastr As Integer



For Each sht In ActiveWorkbook.Worksheets
If SheetExists("WBS_List_Merge") = True Then
Worksheets("WBS_List_Merge").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "WBS_List_Merge"



With Worksheets("WBS_List_Merge")

 Worksheets("WBS_List").Range("a1:X" & CStr(LastRow(Worksheets("WBS_List")))).Copy
.Range("a1:X" & CStr(LastRow(Worksheets("WBS_List")))).PasteSpecial Paste:=xlPasteValues

lastr = LastRow(Worksheets("WBS_List")) + 1

If LastRow(Worksheets("WBS_List_Parent")) > 1 Then
   
 Worksheets("WBS_List_Parent").Range("a2:X" & CStr(LastRow(Worksheets("WBS_List_Parent")))).Copy
.Range("a" & lastr & ":X" & CStr(LastRow(Worksheets("WBS_List_Parent"))) + lastr).PasteSpecial Paste:=xlPasteValues
  
End If

End With
Worksheets("WBS_List_Merge").Visible = False

Application.StatusBar = "done"

End Sub

