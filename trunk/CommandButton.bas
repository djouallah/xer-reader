Attribute VB_Name = "CommandButton"
Public Sub CommandButton1_Click()
Dim ntime As Double
Dim i As Integer
Dim sht1 As Worksheet

Application.ScreenUpdating = False
Application.DisplayAlerts = False
MULTI_PROJECT = False
    
    
ntime = Timer
    
       Call Read_Xer
    If sXerFileName <> "" Then
        
        Call DeletSheets
        Call clearDashboard
        Call Cleardetails
        Call Loadxer
        Call check_Tables
        Call copy_Tables_New_sheets
        ThisWorkbook.Save
        Call GenerateDatasets
        
          i = 0
        For Each sht1 In ActiveWorkbook.Worksheets
        If SheetExists("TASKRSRC") = True And MULTI_PROJECT = False Then
        i = i + 1
        End If
        Next
         If i <> 0 Then
        Call TASKRSRC_SUM
        End If
        
        
        
        
         i = 0
        For Each sht1 In ActiveWorkbook.Worksheets
        If SheetExists("projwbs") = True And MULTI_PROJECT = False Then
        i = i + 1
        End If
        Next
    
       If i <> 0 Then
        Call Create_Temp_Tables
        Call CreateParentChildSheet
        Call printwbs
        ThisWorkbook.Save
        Call Task_list
        ThisWorkbook.Save
        Call Task_list_cost0
        ThisWorkbook.Save
        Call Task_list_cost
        ThisWorkbook.Save
        Call Task_list_Merge
        ThisWorkbook.Save
        Call WBS_list
        ThisWorkbook.Save
        
        Worksheets("Dashboard").Shapes("CommandButton2").Visible = True
        Worksheets("Dashboard").Shapes("CommandButton5").Visible = True
       End If
       
        Worksheets("Dashboard").Shapes("CommandButton3").Visible = True
        Worksheets("dashboard").Activate
        MsgBox "xer Loaded in " & Format(Timer - ntime, "0.00 \s\ec"), vbInformation, APPNAME
        
      
      
 End If
      
Application.ScreenUpdating = True
Application.DisplayAlerts = True


End Sub
Public Sub CommandButton2_Click()
Call clearDashboard
Call Cleardetails
  Worksheets("Dashboard").Shapes("CommandButton2").Visible = False
  Worksheets("Dashboard").Shapes("CommandButton3").Visible = False
  Worksheets("Dashboard").Shapes("CommandButton5").Visible = False
End Sub

Public Sub CommandButton3_Click()
Call DeletSheets
sXerFileName = ""
Project_End_Date = Empty
Project_End_Date = Empty
DataDate = Empty
  Worksheets("Dashboard").Shapes("CommandButton3").Visible = False
  Worksheets("Dashboard").Shapes("CommandButton5").Visible = False
End Sub



Public Sub CommandButton4_Click()
ActiveWorkbook.FollowHyperlink Address:="http://code.google.com/p/xer-dcma14/"
End Sub

Public Sub CommandButton5_Click()
Dim ntime As Double
Dim i As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False

ntime = Timer
 

     i = 0
    For Each sht1 In ActiveWorkbook.Worksheets
    If SheetExists("projwbs") = True And SheetExists("task") = True Then
    i = i + 1
    End If
    Next
    
    If i <> 0 Then
        Call WBS_list_Parent
        ThisWorkbook.Save
        Call WBS_List_Merge
        ThisWorkbook.Save
        Call gantt_Data
        Call Gantt_Chart
    Else
GoTo fail
    End If
    
    MsgBox "Gantt Chart done in " & Format(Timer - ntime, "0.00 \s\ec"), vbInformation, APPNAME
    Exit Sub
    
fail:
    MsgBox "Can't print the Chart, No Task Founds", vbInformation, APPNAME
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


End Sub

