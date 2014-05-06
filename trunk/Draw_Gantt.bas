Attribute VB_Name = "Draw_Gantt"
'  mimoune.djouallah@gmail.com

Option Explicit

Sub gantt_Data()
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
If SheetExists("gantt") = True Then
Worksheets("gantt").Delete
End If
Next


Worksheets.Add(After:=Worksheets("wbs")).Name = "gantt"

With Worksheets("gantt")



    .Sort.SortFields.Clear

'********************** copy the tasks from Task List Merge
Worksheets("Task_list_Merge").Range("a1:X" & CStr(LastRow(Worksheets("Task_list_Merge")))).Copy
.Range("a4:X" & CStr(LastRow(Worksheets("Task_list_Merge")) + 4)).PasteSpecial Paste:=xlPasteValues

    
NT = LastRow(Worksheets("gantt"))
lastr = LastRow(Worksheets("gantt"))



'********************** Transfer WBS data from the Merge WBS
   
    strsql = "SELECT [task_id],[task],[Order],[Level],[wbs_id],[Path],[name],'',sum([Budget_labor]), sum([Actual_Labor]),sum([Remaining]),sum([At_Completion])as [T_AC]" & _
    ",sum([od]) as [T_OD] ," & _
    "min([Start])as [M_Start],min([restart])as [M_restart],max([Finish]) as [M_Finish],min([Total_Float])," & _
    "iif([M_Finish]-[M_start]=0,0,iif([M_Finish]<val([M_restart]),1,iif([M_Start]>=val([M_restart]),0,([M_Finish]-val([M_restart]))/([M_Finish]-[M_Start]))))," & _
    "iif([T_AC]=0,0,sum([At_Completion]*[Activity_Pct])/[T_AC]) AS [A_Pct], sum([Labor_Earned_Value]),SUM([target_cost]),SUM([Actual_cost]), SUM([remain_cost]),SUM([At_Completion_cost]) " & _
    "from [WBS_List_Merge$] GROUP BY [task],[Path],[Order],[Level],[wbs_id],[name],[task_id]"
    
    cnn.Open
    rssql.Open strsql, cnn
    .Cells(lastr + 1, 1).CopyFromRecordset rssql
    

    
    
   


'************************************************************************************


lastr = LastRow(Worksheets("gantt"))


'*************************** paint


lastr = LastRow(Worksheets("gantt"))

tempa = .Range("d1:d" & lastr).Value

For i = NT + 1 To lastr
      
      .Range("A" & i & ":X" & i).Interior.ColorIndex = 14 + tempa(i, 1)
      
      .Rows(i).Font.FontStyle = "Bold Italic"
      .Rows(i).Font.Size = 12
    
       Application.StatusBar = "PAINT WBS: " & i & "..."
       DoEvents

    
Next i
'************************************
'*******************************sort******************************
    .Sort.SortFields.Add Key:=Range("c5:c" & lastr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Sort.SortFields.Add Key:=Range("b5:b" & lastr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .Sort.SortFields.Add Key:=Range("l5:l" & lastr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With .Sort
        .SetRange Range("A4:X" & lastr)
        .Header = xlYes
       .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

      
   
    

'****************************************************************
.Range("a4:X4").AutoFilter
.Columns("h").AutoFit
.Columns("s").AutoFit
.Columns("x").AutoFit


.Columns("n:p").NumberFormat = "dd-mm-yyyy"
.Columns("i:l").NumberFormat = "#,##0"
.Columns("t").NumberFormat = "#,##0"
.Columns("r:s").NumberFormat = "0.00%"
.Columns("q").NumberFormat = "0.00"
.Columns("u:x").NumberFormat = "#,##0.0"

 
.Columns("m").Hidden = True
.Columns("o").Hidden = True

.Columns("a:c").Hidden = True
.Columns("e:f").Hidden = True
.Rows(1).RowHeight = 16
.Rows(2).RowHeight = 16
.Rows(3).RowHeight = 16

.Range("a4:X4").Interior.ColorIndex = 37

'******************copy  percent complete to dashboard*************
Worksheets("dashboard").Cells(8, 2) = .Cells(5, 18)
Worksheets("dashboard").Cells(8, 1) = "Duration % Complete"
Worksheets("dashboard").Cells(8, 2).NumberFormat = "0.00%"

Worksheets("dashboard").Cells(14, 8).NumberFormat = "#,##0"
Worksheets("dashboard").Cells(14, 8) = .Cells(5, 20)



Worksheets("dashboard").Cells(15, 11) = .Cells(5, 19)
Worksheets("dashboard").Cells(14, 11) = Worksheets("dashboard").Cells(15, 11).Value * Worksheets("dashboard").Cells(13, 11).Value
Worksheets("dashboard").Cells(14, 11).NumberFormat = "#,##0"

If Worksheets("dashboard").Cells(13, 8).Value <> 0 Then
Worksheets("dashboard").Cells(15, 8) = Worksheets("dashboard").Cells(14, 8).Value / Worksheets("dashboard").Cells(13, 8).Value
Worksheets("dashboard").Cells(15, 8).NumberFormat = "0.00%"
End If


'********************************
 cnn.Close
 Set cnn = Nothing
 Set rssql = Nothing

End With

End Sub
Sub Gantt_Chart()

Dim d1 As Range
Dim d2 As Range
Dim start_cell As String

Dim Startx As Double
Dim Starty As Double
Dim Starty_end As Double
Dim Calen_Month As Date
Dim Calen_Year As Date

Dim counter As Double
Dim Fa As Double
Dim start As Integer
Dim finish As Integer
Dim restart As Integer
Dim TF As Integer
Dim lastr As Integer
Dim tempa As Variant


Dim x As Integer

Dim A_Duration As Double
Dim R_Duration As Double
Dim Duration As Double
Dim Lenght As Double
Dim GAP As Double
Dim GAP1 As Double ' from the project start to the first friday
Dim GAP2 As Double  ' from the last friday to the project end
Dim GAP3 As Double  ' from the first friday to the last friday
Dim shp As Shape



On Error Resume Next

Project_Start_Date = Worksheets("dashboard").Cells(6, 2)
Project_End_Date = Worksheets("dashboard").Cells(7, 2)
DataDate = Worksheets("dashboard").Cells(5, 2)

Fa = 3 ' a dummy Factor to increase the size of the weekly textbox.

start = 14
restart = 15
finish = 16
TF = 17
start_cell = "AA"

lastr = LastRow(Worksheets("gantt"))

With Worksheets("gantt")


    .Columns("y:" & start_cell).ColumnWidth = 0.5

    .Columns("ab:BFU").ColumnWidth = 1

    .Activate
      
    
    '........................................Draw Timescale Grid
    
    ' .......................... draw the the first year
    Set d1 = .Range(start_cell & 1)
    
    Startx = d1.Left + d1.Width
    Starty = d1.Top
    
    
    
    If Project_End_Date > DateSerial(Year(Project_Start_Date), 12, 31) Then
    
    counter = (DateSerial(Year(Project_Start_Date), 12, 31) - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date)))
    Else
     counter = Project_End_Date - Project_Start_Date

    End If
    
    
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx, Starty, counter * Fa, d1.Height)
    
    .TextFrame.Characters.text = CStr(Format(Project_Start_Date, "yyyy"))
    .TextFrame.Characters.Font.Size = 10
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    
    
    ' .......................... draw the middle years
    
    
    Calen_Year = DateSerial(Year(Project_Start_Date), 12, 31) + 1
    
    Do While Calen_Year + Yearday(Calen_Year) <= Project_End_Date
    
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx + counter * Fa, Starty, Yearday(Calen_Year) * Fa, d1.Height)
    .TextFrame.Characters.text = CStr(Format(Calen_Year, "yyyy"))
    .TextFrame.Characters.Font.Size = 10
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    
    counter = counter + Yearday(Calen_Year)
    Calen_Year = Calen_Year + Yearday(Calen_Year)
    
    
    Loop
    '................................. draw the last year
    
    If (Project_End_Date - Project_Start_Date - counter) > 0 Then
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx + counter * Fa, Starty, (DateSerial(Year(Project_End_Date), Month(Project_End_Date), Day(Project_End_Date)) + 1 - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date)) - counter) * Fa, d1.Height)
    .TextFrame.Characters.text = CStr(Format(Project_End_Date, "yyyy"))
    .TextFrame.Characters.Font.Size = 10
    .TextFrame.HorizontalAlignment = xlCenter
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    End If
    
    
    
    
    
    
    ' .......................... draw the months
    Set d1 = .Range(start_cell & 2)
    
    Startx = d1.Left + d1.Width
    Starty = d1.Top
    
    ' .......................... draw the first month
    
    
    
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx, Starty, (dhDaysInMonth(Project_Start_Date) - Day(Project_Start_Date)) * Fa, d1.Height)
    .TextFrame.Characters.text = CStr(Format(Project_Start_Date, "mmm"))
    .TextFrame.Characters.Font.Size = 8
    .TextFrame.HorizontalAlignment = xlLeft
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    
    
    ' .......................... draw the middle months
    
    
    Calen_Month = Project_Start_Date - Day(Project_Start_Date) + dhDaysInMonth(Project_Start_Date) + 1
    
    counter = (dhDaysInMonth(Project_Start_Date) - Day(Project_Start_Date))
    
    Do While Calen_Month <= Project_End_Date - Day(Project_End_Date)
    
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx + counter * Fa, Starty, dhDaysInMonth(Calen_Month) * Fa, d1.Height)
    .TextFrame.Characters.text = CStr(Format(Calen_Month, "mmm"))
    .TextFrame.Characters.Font.Size = 8
    .TextFrame.HorizontalAlignment = xlLeft
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    
    counter = counter + dhDaysInMonth(Calen_Month)
    Calen_Month = Calen_Month + dhDaysInMonth(Calen_Month)
    
    
    Loop
    '................................. draw the last month
    
    If (Project_End_Date - Project_Start_Date - counter) > 0 Then
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx + counter * Fa, Starty, (DateSerial(Year(Project_End_Date), Month(Project_End_Date), Day(Project_End_Date)) + 1 - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date)) - counter) * Fa, d1.Height)
    .TextFrame.Characters.text = CStr(Format(Project_End_Date, "mmm"))
    .TextFrame.Characters.Font.Size = 8
    .TextFrame.HorizontalAlignment = xlLeft
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    End If
    
    
    ' .......................... draw the first friday after the start day
    
    Set d1 = .Range(start_cell & 3)
    
    Startx = d1.Left + d1.Width
    Starty = d1.Top
    
    Select Case Weekday(Project_Start_Date)
    Case 7
    GAP1 = 6
    Case 6
    GAP1 = 0
    Case 5
    GAP1 = 1
    Case 4
    GAP1 = 2
    Case 3
    GAP1 = 3
    Case 2
    GAP1 = 4
    Case 1
    GAP1 = 5
    End Select
    
    If GAP1 > 0 Then
    
    
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx, Starty, GAP1 * Fa, d1.Height)
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    End If
    
    '..................................Draw the weekly vertical grid
    Select Case Weekday(Project_End_Date)
    Case 7
    GAP2 = 6
    Case 6
    GAP2 = 0
    Case 5
    GAP2 = 1
    Case 4
    GAP2 = 2
    Case 3
    GAP2 = 3
    Case 2
    GAP2 = 4
    Case 1
    GAP2 = 5
    End Select
    
    GAP3 = GAP1
    Do While GAP3 <= (DateSerial(Year(Project_End_Date), Month(Project_End_Date), Day(Project_End_Date)) + 1 - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date)) - 7)
    
    With .Shapes.AddTextbox(msoTextOrientationHorizontal, Startx + GAP3 * Fa, Starty, 7 * Fa, d1.Height)
    .TextFrame.Characters.text = CStr(Day(Project_Start_Date + GAP3 + 1))
    .TextFrame.Characters.Font.Size = 6
    .TextFrame.HorizontalAlignment = xlLeft
    .TextFrame.VerticalAlignment = xlCenter
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(255, 0, 18)
    .Fill.ForeColor.RGB = RGB(255, 255, 204)
    .Line.Weight = 1
    End With
    GAP3 = GAP3 + 7
    
     Application.StatusBar = "Draw Timescale: "
           DoEvents
    
    Loop
    
    
    
    
    
    '.......................................................................................................
    
    
    '...........................DRAW THE VERTICAL GRID .......................................
    '............................Draw Start Date
    
     With .Range("ab4:ab" & lastr)
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        
        
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        
    End With
    
    
    '...................................Draw DataDate
    Set d1 = .Range(start_cell & 3)
    Set d2 = .Range("u" & lastr)
    
    
    Startx = d1.Left + d1.Width
    Starty = d1.Top + d1.Height
    Starty_end = d2.Top + d2.Height
    
    
    
    GAP = DataDate - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date))
    
    If GAP > 0 Then
    
    With .Shapes.AddLine(Startx + GAP * Fa, Starty, Startx + GAP * Fa, Starty_end)
    .Line.Weight = 1#
    .Line.DashStyle = msoLineLongDash
    .Line.ForeColor.RGB = RGB(255, 0, 0)
    End With
    End If
    
    '................................................Draw End Date
    
    GAP = (DateSerial(Year(Project_End_Date), Month(Project_End_Date), Day(Project_End_Date)) + 1 - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date)))
    If GAP > 0 Then
    
    With .Shapes.AddLine(Startx + GAP * Fa, d1.Top, Startx + GAP * Fa, Starty_end)
    .Line.Weight = 1#
    .Line.DashStyle = msoLineSolid
    .Line.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    End If
    
    
    
    
    '.................................Draw THE Tasks......................................................
    tempa = .Range("a1:q" & lastr).Value
    For x = 5 To lastr
    Set d1 = .Range(start_cell & x)
    
    
    
    
    
    Startx = d1.Left + d1.Width
    Starty = d1.Top + d1.Height / 2
    GAP = (tempa(x, start) - DateSerial(Year(Project_Start_Date), Month(Project_Start_Date), Day(Project_Start_Date)))
    Duration = (tempa(x, finish) - tempa(x, start))
    
    
 
    
    If DataDate <= tempa(x, start) Then
    
    R_Duration = (tempa(x, finish) - tempa(x, start))
    
            If R_Duration > 0 Then
            
              With .Shapes.AddLine(Startx + GAP * Fa, Starty, Startx + R_Duration * Fa + GAP * Fa, Starty)
               
                If tempa(x, 3) = "" Then
                .Line.DashStyle = msoLineSolid
                 .Line.Weight = 10#
                Else
                 '.Line.DashStyle = msoLineLongDash
                  .Line.Weight = 5#
                End If
                If (tempa(x, TF) <= 0 And tempa(x, TF) <> "") Then
                .Line.ForeColor.RGB = RGB(255, 0, 0)
                Else
               .Line.ForeColor.RGB = RGB(0, 176, 80)
                End If
                End With
            End If
            
             If Duration = 0 Then
            
           With .Shapes.AddShape(msoShape4pointStar, Startx + GAP * Fa - 5 * Fa, d1.Top, 10 * Fa, 10)
           End With
            End If
            
    
    End If
    
    If DataDate > tempa(x, start) And DataDate < tempa(x, finish) And tempa(x, restart) <> "" Then
    
            
            A_Duration = (DataDate - tempa(x, start))
            
            R_Duration = (tempa(x, finish) - tempa(x, restart))
            
            If Duration > 0 Then
            
                With .Shapes.AddLine(Startx + GAP * Fa, Starty, Startx + A_Duration * Fa + GAP * Fa, Starty)
                
                 If tempa(x, 3) = "" Then
                .Line.DashStyle = msoLineSolid
                .Line.Weight = 10#
                Else
                 '.Line.DashStyle = msoLineLongDash
                 .Line.Weight = 5#
                End If
               
                
                
                If tempa(x, TF) <= 0 And tempa(x, TF) <> "" Then
                .Line.ForeColor.RGB = RGB(255, 0, 0)
                Else
                .Line.ForeColor.RGB = RGB(0, 176, 240)
                End If
                End With
            
             With .Shapes.AddLine(Startx + GAP * Fa + A_Duration * Fa + (tempa(x, restart) - DataDate) * Fa, Starty, Startx + GAP * Fa + A_Duration * Fa + (tempa(x, restart) - DataDate) * Fa + R_Duration * Fa, Starty)
                 
                  If tempa(x, 3) = "" Then
                .Line.DashStyle = msoLineSolid
                .Line.Weight = 10#
                Else
                 '.Line.DashStyle = msoLineLongDash
                 .Line.Weight = 5#
                End If
                
                If tempa(x, TF) <= 0 And tempa(x, TF) <> "" Then
                .Line.ForeColor.RGB = RGB(255, 0, 0)
                Else
                .Line.ForeColor.RGB = RGB(0, 176, 80)
                End If
                End With
            End If
            
            If Duration = 0 Then
            
            With .Shapes.AddShape(msoShape4pointStar, Startx + GAP * Fa - 5 * Fa, d1.Top, 10 * Fa, 10)
            End With
                 
            End If
            
    End If
    
    If tempa(x, restart) = "" Then
    
    A_Duration = (tempa(x, finish) - tempa(x, start))
    
            If Duration > 0 Then
            
                With .Shapes.AddLine(Startx + GAP * Fa, Starty, Startx + A_Duration * Fa + GAP * Fa, Starty)
               
                 If tempa(x, 3) = "" Then
                .Line.DashStyle = msoLineSolid
                 .Line.Weight = 10#
                Else
                 '.Line.DashStyle = msoLineLongDash
                  .Line.Weight = 5#
                End If
                
                If tempa(x, TF) <= 0 And tempa(x, TF) <> "" Then
                .Line.ForeColor.RGB = RGB(255, 0, 0)
                Else
                 .Line.ForeColor.RGB = RGB(0, 176, 240)
                End If
                End With
            End If
            
             If Duration = 0 Then
            
            With .Shapes.AddShape(msoShape4pointStar, Startx + GAP * Fa - 5 * Fa, d1.Top, 10 * Fa, 10)
            
            
            End With
           
            
            End If
            End If
        
        
        
        
     


     Application.StatusBar = "Draw Tasks: " & x & " of " & lastr & ", " & Format(x / lastr, "0%")
       DoEvents

        Next x
         Application.StatusBar = ""
         DoEvents
         
         
         
.Range("F1").Select
End With



End Sub


