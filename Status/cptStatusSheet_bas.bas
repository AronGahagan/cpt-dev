Attribute VB_Name = "cptStatusSheet_bas"
'<cpt_version>v0.1</cpt_version>
Option Explicit

Declare Function GetTickCount Lib "kernel32" () As Long
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private Const adVarChar As Long = 200

Sub ShowCptStatusSheet_frm()
'populate all outline codes, text, and number fields
'populate UID,[user selections],Task Name,Duration,Forecast Start,Forecast Finish,Total Slack,[EVT],EV%,New EV%,BLW,Remaining Work,Revised ETC,BLS,BLF,Reason/Impact/Action
'add pick list for EV% or default to Physical % Complete
'objects
Dim arrFields As Object, arrEVT As Object, arrEVP As Object
'longs
Dim lgField As Long, lgItem As Long
'integers
Dim intField As Integer
'strings
Dim strFieldName As String, strFileName As String
'dates
Dim dtStatus As Date
'variants
Dim st As Variant, strFieldType As Variant, lgFieldType As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'requires ms excel
  If Not CheckReference("Excel") Then GoTo exit_here
  'requires scripting (RegEx)
  If Not CheckReference("Scripting") Then GoTo exit_here
  
  cptStatusSheet_frm.lboFields.Clear
  cptStatusSheet_frm.lboExport.Clear
  cptStatusSheet_frm.cboEVT.Clear
  cptStatusSheet_frm.cboEVP.Clear
  cptStatusSheet_frm.cboEVP.AddItem "Physical % Complete"
  cptStatusSheet_frm.cboCostTool.Clear
  cptStatusSheet_frm.cboCostTool.AddItem "COBRA"
  cptStatusSheet_frm.cboCostTool.AddItem "MPM"
  cptStatusSheet_frm.cboCostTool.AddItem "<none>"
  
  Set arrFields = CreateObject("System.Collections.SortedList")
  Set arrEVT = CreateObject("System.Collections.SortedList")
  Set arrEVP = CreateObject("System.Collections.SortedList")
  
  For Each strFieldType In Array("Text", "Outline Code", "Number")
    On Error GoTo err_here
    For intField = 1 To 30
      lgField = FieldNameToFieldConstant(strFieldType & intField, lgFieldType)
      strFieldName = CustomFieldGetName(lgField)
      If Len(strFieldName) > 0 Then
        arrFields.Add strFieldName, lgField
        If strFieldType = "Text" Then
          arrEVT.Add strFieldName, lgField
          'todo: what if this is an enterprise field?
        ElseIf strFieldType = "Number" Then
          arrEVP.Add strFieldName, lgField
          'todo: what if this is an enterprise field?
        End If
      End If
next_field:
    Next intField
  Next strFieldType
  
  'get enterprise custom fields
  For lgField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lgField) <> "<Unavailable>" Then
      arrFields.Add Application.FieldConstantToFieldName(lgField), lgField
    End If
  Next lgField
  
  'add custom fields
  For intField = 0 To arrFields.count - 1
    cptStatusSheet_frm.lboFields.AddItem
    cptStatusSheet_frm.lboFields.List(intField, 0) = arrFields.getByIndex(intField)
    cptStatusSheet_frm.lboFields.List(intField, 1) = arrFields.getKey(intField)
    If FieldNameToFieldConstant(arrFields.getKey(intField)) >= 188776000 Then
      cptStatusSheet_frm.lboFields.List(intField, 2) = "Enterprise"
    Else
      cptStatusSheet_frm.lboFields.List(intField, 2) = FieldConstantToFieldName(arrFields.getByIndex(intField))
    End If
    cptStatusSheet_frm.cboEach.AddItem arrFields.getKey(intField)
  Next
  'add EVT values
  For intField = 0 To arrEVT.count - 1
    cptStatusSheet_frm.cboEVT.AddItem arrEVT.getKey(intField)
  Next
  'add EVP values
  For intField = 0 To arrEVP.count - 1 'UBound(st)
    cptStatusSheet_frm.cboEVP.AddItem arrEVP.getKey(intField) 'st(intField)(1)
  Next
  
  'add saved settings if they exist
  strFileName = cptDir & "\settings\cpt-status-sheet.adtg"
  If Dir(strFileName) <> vbNullString Then
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      .MoveFirst
      On Error Resume Next
      cptStatusSheet_frm.cboEVT.Value = .Fields(0) 'cboEVT
      cptStatusSheet_frm.cboEVP.Value = .Fields(1) 'cboEVP
      cptStatusSheet_frm.optWorkbook = .Fields(2) = 1 'chkOutput
      cptStatusSheet_frm.optWorksheets = .Fields(2) = 2 'chkOutput
      cptStatusSheet_frm.optWorkbooks = .Fields(2) = 3 'chkOutput
      cptStatusSheet_frm.chkHide = .Fields(3) = 1 'chkHide
      If .Fields.count >= 5 Then
        If Not IsNull(.Fields(4)) Then cptStatusSheet_frm.cboCostTool.Value = .Fields(4) 'cboCostTool
      End If
      If .Fields.count >= 6 Then
        If Not IsNull(.Fields(5)) Then cptStatusSheet_frm.cboEach.Value = .Fields(5) 'cboEach
      End If
      .Close
    End With
  End If
  
  'add saved export fields if they exist
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      .MoveFirst
      lgItem = 0
      Do While Not .EOF
        cptStatusSheet_frm.lboExport.AddItem
        cptStatusSheet_frm.lboExport.List(lgItem, 0) = .Fields(0) 'Field Constant
        cptStatusSheet_frm.lboExport.List(lgItem, 1) = .Fields(1) 'Custom Field Name
        cptStatusSheet_frm.lboExport.List(lgItem, 2) = .Fields(2) 'Local Field Name
        lgItem = lgItem + 1
        .MoveNext
      Loop
      .Close
    End With
  End If
  
  'set the status date / hide complete
  If ActiveProject.StatusDate = "NA" Then
    cptStatusSheet_frm.txtStatusDate.Value = FormatDateTime(DateAdd("d", 6 - Weekday(Now), Now), vbShortDate)
  Else
    cptStatusSheet_frm.txtStatusDate = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  End If
  
  'delete pre-existing search file
  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  
  dtStatus = CDate(cptStatusSheet_frm.txtStatusDate.Value)
  cptStatusSheet_frm.txtHideCompleteBefore.Value = DateAdd("d", -(Day(dtStatus) - 1), dtStatus)
  cptStatusSheet_frm.Show False

exit_here:
  On Error Resume Next
  Set arrFields = Nothing
  Set arrEVT = Nothing
  Set arrEVP = Nothing
  Exit Sub

err_here:
  If err.Number = 1101 Or err.Number = 1004 Then
    err.Clear
    Resume next_field
  Else
    MsgBox err.Number & ": " & err.Description, vbExclamation + vbOKOnly, "Error"
    Resume exit_here
  End If

End Sub

Sub CreateStatusSheet()
'objects
Dim Tasks As Tasks, Task As Task, Resource As Resource, Assignment As Assignment
'early binding:
'Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Excel.Range
'Dim rSummaryTasks As Excel.Range, rMilestones As Excel.Range, rNormal As Excel.Range, rAssignments As Excel.Range, rLockedCells As Excel.Range
'Dim rDates As Excel.Range, rWork As Excel.Range, rMedium As Excel.Range, rCentered As Excel.Range, rEntry As Excel.Range
'Dim xlCells As Excel.Range, rngAll As Excel.Range
'late binding:
Dim xlApp As Object, Workbook As Object, Worksheet As Object, rng As Object
Dim rSummaryTasks As Object, rMilestones As Object, rNormal As Object, rAssignments As Object, rLockedCells As Object
Dim rDates As Object, rWork As Object, rMedium As Object, rCentered As Object, rEntry As Object
Dim xlCells As Object, rngAll As Object
Dim aSummaries As Object, aMilestones As Object, aNormal As Object, aAssignments As Object
Dim aEach As Object, aTaskRow As Object, aHeaders As Object
Dim aOddBalls As Object, aCentered As Object, aEntryHeaders As Object
'longs
Dim lgTaskCount As Long, lgTask As Long, lgHeaderRow As Long, lgLastCol As Long
Dim lgRow As Long, lgLastRow As Long, lgCol As Long, lgField As Long
Dim lgNameCol As Long, lgBaselineWorkCol As Long, lgRemainingWorkCol As Long, lgEach As Long
Dim lgNotesCol As Long, lgColumnWidth As Long
Dim lgASCol As Long, lgAFCol As Long, lgETCCol As Long, lgEVPCol As Long
Dim t As Long, tTotal As Long
'string
Dim strFieldName As String, strEVT As String, strEVP As String, strDir As String, strFileName As String
Dim strFirstCell As String
'dates
Dim dtStatus As Date
'variants
Dim vCol As Variant, aUserFields As Variant
'booleans
Dim blnFast As Boolean

  tTotal = GetTickCount
  
  'check reference
  If Not CheckReference("Excel") Then GoTo exit_here

  'ensure required module exists
  If Not ModuleExists("cptCore_bas") Then
    MsgBox "Please install the ClearPlan 'cptCore_bas' module.", vbExclamation + vbOKOnly, "Missing Module"
    GoTo exit_here
  End If
  
  blnFast = True
  
  On Error Resume Next
  Set Tasks = ActiveProject.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If Tasks Is Nothing Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "Create Status Sheet"
    GoTo exit_here
  End If
  
  cptStatusSheet_frm.lblStatus.Caption = " Analyzing project..."
  'get task count
  t = GetTickCount
  For Each Task In Tasks
    lgTaskCount = lgTaskCount + 1
  Next Task
  Debug.Print "<=====PERFORMANCE TEST " & Now() & "=====>"
  Debug.Print "get task count: " & (GetTickCount - t) / 1000
  
  cptStatusSheet_frm.lblStatus.Caption = " Setting up workbook..."
  'set up an excel workbook
  t = GetTickCount
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = xlCalculationManual
  xlApp.ScreenUpdating = False
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = "Status Sheet"
  Set xlCells = Worksheet.Cells
  Debug.Print "set up excel workbook: " & (GetTickCount - t) / 1000
  
  'set up legend
  t = GetTickCount
  xlCells(1, 1).Value = "Status Date:"
  xlCells(1, 1).Font.Bold = True
  If ActiveProject.StatusDate = "NA" Then
    dtStatus = Now()
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  xlCells(1, 2) = FormatDateTime(dtStatus, vbShortDate)
  Worksheet.Names.Add "STATUS_DATE", Worksheet.[B1]
  xlCells(1, 2).Font.Bold = True
  xlCells(1, 2).Font.Size = 14
  'current
  xlCells(3, 1).Interior.ThemeColor = xlThemeColorAccent2
  xlCells(3, 1).Interior.TintAndShade = 0.799981688894314
  xlCells(3, 2) = "Task is active or within current status window.  Update Required."
  'within two weeks
  xlCells(4, 1).Interior.ThemeColor = xlThemeColorAccent5
  xlCells(4, 1).Interior.TintAndShade = 0.799981688894314
  xlCells(4, 2) = "Task is within two week look-ahead.  Please review forecast dates."
  'complete
  xlCells(5, 1) = "AaBbCc"
  xlCells(5, 1).Font.Italic = True
  xlCells(5, 1).Font.ColorIndex = 16
  xlCells(5, 2) = "Task is complete."
  'summary
  xlCells(6, 1) = "AaBbCc"
  xlCells(6, 1).Font.Bold = True
  xlCells(6, 1).Interior.ThemeColor = xlThemeColorDark1
  xlCells(6, 1).Interior.TintAndShade = -0.149998474074526
  xlCells(6, 2) = "MS Project Summary Task (Rollup).  No update required."
  Debug.Print "set up legend: " & (GetTickCount - t) / 1000

  lgHeaderRow = 8
  
  'set up header
  t = GetTickCount
  
  'get selected fields for two non-standard fields
  strEVT = cptStatusSheet_frm.cboEVT.Value
  strEVP = cptStatusSheet_frm.cboEVP.Value
  
  'set up header
  Set aHeaders = CreateObject("System.Collections.ArrayList")
  
  'define non-standard columwidths - default is 10
  Set aOddBalls = CreateObject("System.Collections.SortedList")
  aOddBalls.Add "Name", 60
  aOddBalls.Add "Duration", 8
  aOddBalls.Add "Total Slack", 8
  aOddBalls.Add strEVT, 5
  aOddBalls.Add strEVP, 5
  aOddBalls.Add "Notes", 45
  
  'add standard local fields, required EVT and EV%
  'some of these will be renamed later
  For Each vCol In Array("Unique ID", _
                    "Name", _
                    "Duration", _
                    "Start", _
                    "Finish", _
                    "Actual Start", _
                    "Actual Finish", _
                    "Total Slack", _
                    strEVT, _
                    strEVP, _
                    "Baseline Work", _
                    "Remaining Work", _
                    "Baseline Start", _
                    "Baseline Finish", _
                    "Notes")
    If aOddBalls.contains(vCol) Then
      lgColumnWidth = aOddBalls.Item(vCol)
    Else
      lgColumnWidth = 10 'default
    End If
    aHeaders.Add Array(FieldNameToFieldConstant(vCol), vCol, lgColumnWidth)
  Next vCol
  
  'save fields to adtg file
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  aUserFields = cptStatusSheet_frm.lboExport.List()
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", adVarChar, 255
    .Fields.Append "Custom Field Name", adVarChar, 255
    .Fields.Append "Local Field Name", adVarChar, 255
    .Open
    For lgField = 0 To UBound(aUserFields)
      .AddNew Array(0, 1, 2), Array(aUserFields(lgField, 0), aUserFields(lgField, 1), aUserFields(lgField, 2))
    Next lgField
    .Update
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    .Save strFileName
    .Close
  End With
  
  '/===debug===\
  If MsgBox("man, that's a lot of tasks. are you sure?", vbYesNo + vbQuestion, Format(lgTaskCount, "#,##0")) = vbNo Then
    Workbook.Close False
    xlApp.Quit
    GoTo exit_here
  End If
  '\===debug===/
  
  'get user fields
  For lgField = UBound(aUserFields) To 0 Step -1
    If aUserFields(lgField, 1) = strEVT Then GoTo next_field
    If aUserFields(lgField, 1) = strEVP Then GoTo next_field
    If lgField = 0 Then lgColumnWidth = 16 Else lgColumnWidth = 10
    aHeaders.Insert 1, Array(aUserFields(lgField, 0), aUserFields(lgField, 1), lgColumnWidth)
next_field:
  Next lgField
    
  'write the headers and size the columns
  For lgField = 0 To aHeaders.count - 1
    xlCells(lgHeaderRow, lgField + 1).Value = aHeaders(lgField)(1)
    xlCells(lgHeaderRow, lgField + 1).EntireColumn.ColumnWidth = aHeaders(lgField)(2)
    'get columns needed later
    If aHeaders(lgField)(1) = "Name" Then lgNameCol = lgField + 1
    If aHeaders(lgField)(1) = "Actual Start" Then lgASCol = lgField + 1
    If aHeaders(lgField)(1) = "Actual Finish" Then lgAFCol = lgField + 1
    If aHeaders(lgField)(1) = "Baseline Work" Then lgBaselineWorkCol = lgField + 1
    If aHeaders(lgField)(1) = "Remaining Work" Then lgRemainingWorkCol = lgField + 1
    If aHeaders(lgField)(1) = "Notes" Then lgNotesCol = lgField + 1
  Next

  'format the header row
  With xlCells(lgHeaderRow, 1).Resize(, aHeaders.count)
    .Interior.ThemeColor = xlThemeColorLight2
    .Interior.TintAndShade = 0
    .Font.ThemeColor = xlThemeColorDark1
    .Font.TintAndShade = 0
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
  End With
  Debug.Print "set up header: " & (GetTickCount - t) / 1000
  
  'prepare to capture each
  If cptStatusSheet_frm.optWorkbook = False Then
    Set aEach = CreateObject("System.Collections.SortedList")
    lgEach = FieldNameToFieldConstant(cptStatusSheet_frm.cboEach.Value)
  End If
  
  'prepare to capture bulk ranges
  Set aTaskRow = CreateObject("System.Collections.ArrayList")
  Set aAssignments = CreateObject("System.Collections.ArrayList")
  Set aSummaries = CreateObject("System.Collections.ArrayList")
  Set aMilestones = CreateObject("System.Collections.ArrayList")
  Set aNormal = CreateObject("System.Collections.ArrayList")
  
  'capture task data
  t = GetTickCount
  lgRow = lgHeaderRow
  For Each Task In Tasks
    If Task Is Nothing Then GoTo next_task
    If Task.OutlineLevel = 0 Then GoTo next_task
    If Task.ExternalTask Then GoTo next_task
    If Not Task.Active Then GoTo next_task
    If cptStatusSheet_frm.chkHide = True Then
      If Task.ActualFinish <= CDate(cptStatusSheet_frm.txtHideCompleteBefore) Then GoTo next_task
    End If
    
    lgRow = lgRow + 1
    
    If cptStatusSheet_frm.optWorkbook = False Then
      aEach.Add Task.GetField(lgEach), Task.GetField(lgEach)
    End If
    
    'get common data
    For lgCol = 1 To lgNameCol
      aTaskRow.Add Task.GetField(aHeaders(lgCol - 1)(0))
    Next lgCol
    
    'indent the task name
    xlCells(lgRow, lgNameCol).IndentLevel = Task.OutlineLevel + 1
    
    'todo: error writing to worksheet
    If Task.Summary Then
      xlCells(lgRow, 1).Resize(, aTaskRow.count).Value = aTaskRow.ToArray()
      aTaskRow.Clear
      aSummaries.Add lgRow
    Else
      For lgCol = lgNameCol + 1 To aHeaders.count
        If aHeaders(lgCol - 1)(1) = "Baseline Work" Then
          aTaskRow.Add Task.BaselineWork / 60
        ElseIf aHeaders(lgCol - 1)(1) = "Remaining Work" Then
          aTaskRow.Add Task.RemainingWork / 60
        Else
          aTaskRow.Add Task.GetField(aHeaders(lgCol - 1)(0))
        End If
      Next lgCol
      
      'identify for formatting
      If Task.Milestone Then aMilestones.Add lgRow Else aNormal.Add lgRow
      
      xlCells(lgRow, lgLastCol + 1).Value = (GetTickCount - t) / 1000
      
      'write task data to sheet
      xlCells(lgRow, 1).Resize(, aTaskRow.count).Value = aTaskRow.ToArray()
      aTaskRow.Clear
      
      'get assignment data for incomplete tasks
      If Task.ActualFinish = "NA" Then
        'add a rollup formlua for Revised ETC?
        For Each Assignment In Task.Assignments
          lgRow = lgRow + 1
          'xlCells(lgRow, 1).Value = Assignment.UniqueID
          aTaskRow.Add Assignment.UniqueID
          If lgNameCol > 2 Then
            For lgCol = 2 To lgNameCol - 1
              aTaskRow.Add Task.GetField(aHeaders(lgCol - 1)(0))
            Next
          End If
          'identify for formatting
          aAssignments.Add lgRow

          'xlCells(lgRow, lgNameCol).Value = Assignment.ResourceName
          aTaskRow.Add Assignment.ResourceName
          xlCells(lgRow, lgNameCol).IndentLevel = Task.OutlineLevel + 2
          xlCells(lgRow, lgBaselineWorkCol).Value = Assignment.BaselineWork / 60
          xlCells(lgRow, lgRemainingWorkCol).Value = Assignment.RemainingWork / 60
          xlCells(lgRow, 1).Resize(, aTaskRow.count).Value = aTaskRow.ToArray()
          aTaskRow.Clear
          
next_assignment:
          '/===debug===\
          'xlCells(lgRow, aHeaders.count + 1).Value = (GetTickCount - t) / 1000
          '\===debug===/
        Next Assignment
      End If 'Task.ActualFinish = "NA"
      
    End If 'Task Summary
    
next_task:
    lgTask = lgTask + 1
    '/===debug===\
    'output execution time for the task
    'xlCells(lgRow, aHeaders.count + 1).Value = (GetTickCount - t) / 1000
    If Not BLN_TRAP_ERRORS And lgRow >= 100 Then Exit For
    '\===debug===/
    Application.StatusBar = "Exporting..." & Format(lgTask, "#,##0") & " / " & Format(lgTaskCount, "#,##0") & " (" & Format(lgTask / lgTaskCount, "0%") & ")"
    cptStatusSheet_frm.lblStatus.Caption = " Exporting..." & Format(lgTask, "#,##0") & " / " & Format(lgTaskCount, "#,##0") & " (" & Format(lgTask / lgTaskCount, "0%") & ")"
    cptStatusSheet_frm.lblProgress.Width = (lgTask / (lgTaskCount)) * cptStatusSheet_frm.lblStatus.Width
  Next Task
  
  Debug.Print "capture task data: " & (GetTickCount - t) / 1000 & " >> " & Format(((GetTickCount - t) / 1000) / (lgRow - lgHeaderRow), "#0.00000") & " per task"
  
  t = GetTickCount
  'add New EV% after EV% - update aHeaders
  lgEVPCol = Worksheet.Rows(lgHeaderRow).Find(strEVP).Column + 1
  Worksheet.Columns(lgEVPCol).Insert Shift:=xlToRight
  xlCells(lgHeaderRow, lgEVPCol).Value = "New EV%"
  aHeaders.Insert lgEVPCol - 1, Array(0, "New EV%", 10)
    
  'add Revised ETC after Remaining Work - update aHeaders
  lgETCCol = Worksheet.Rows(lgHeaderRow).Find("Remaining Work").Column + 1
  Worksheet.Columns(lgETCCol).Insert Shift:=xlToRight
  xlCells(lgHeaderRow, lgETCCol).Value = "Revised ETC"
  aHeaders.Insert lgETCCol - 1, Array(0, "Revised ETC", 10)
  Debug.Print "add columns: " & (GetTickCount - t) / 1000

  cptStatusSheet_frm.lblStatus = " Formatting rows..."
  t = GetTickCount
  'format rows
  'format summary tasks
  Set rSummaryTasks = xlCells(aSummaries(0), 1).Resize(, aHeaders.count)
  For vCol = 1 To aSummaries.count - 1
    Set rSummaryTasks = xlApp.Union(rSummaryTasks, xlCells(aSummaries(vCol), 1).Resize(, aHeaders.count))
  Next vCol
  If Not rSummaryTasks Is Nothing Then
    rSummaryTasks.Interior.ThemeColor = xlThemeColorDark1
    rSummaryTasks.Interior.TintAndShade = -0.149998474074526
    rSummaryTasks.Font.Bold = True
  End If
  'format milestones
  Set rMilestones = xlCells(aMilestones(0), 1).Resize(, aHeaders.count)
  For vCol = 1 To aMilestones.count - 1
    Set rMilestones = xlApp.Union(rMilestones, xlCells(aMilestones(vCol), 1).Resize(, aHeaders.count))
  Next vCol
  If Not rMilestones Is Nothing Then
    rMilestones.Font.ThemeColor = xlThemeColorAccent6
    rMilestones.Font.TintAndShade = -0.249977111117893
  End If
  'format normal
  Set rNormal = xlCells(aNormal(0), 1).Resize(, aHeaders.count)
  For vCol = 1 To aNormal.count - 1
    Set rNormal = xlApp.Union(rNormal, xlCells(aNormal(vCol), 1).Resize(, aHeaders.count))
  Next vCol
  If Not rNormal Is Nothing Then
    rNormal.Font.ThemeColor = xlThemeColorAccent1
    rNormal.Font.TintAndShade = -0.499984740745262
  End If
  'format assignments
  Set rAssignments = xlCells(aAssignments(0), 1).Resize(, aHeaders.count)
  For vCol = 1 To aAssignments.count - 1
    Set rAssignments = xlApp.Union(rAssignments, xlCells(aAssignments(vCol), 1).Resize(, aHeaders.count))
  Next vCol
  If Not rAssignments Is Nothing Then rAssignments.Font.Italic = True
  Debug.Print "format rows: " & (GetTickCount - t) / 1000
  
  t = GetTickCount
  'format common borders
  Set rng = Worksheet.Range(xlCells(lgHeaderRow, 1), xlCells(lgRow, aHeaders.count))
  rng.BorderAround xlContinuous, xlThin
  rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
  rng.Borders(xlInsideHorizontal).Weight = xlThin
  Debug.Print "format common borders: " & (GetTickCount - t) / 1000
  
  t = GetTickCount
  'rename headers
  Set rng = xlCells(lgHeaderRow, 1).Resize(, aHeaders.count)
  rng.Replace what:="Unique ID", Replacement:="UID", lookat:=xlWhole
  rng.Replace what:="Name", Replacement:="Task Name", lookat:=xlWhole
  rng.Replace what:="Start", Replacement:="Forecast Start", lookat:=xlWhole
  rng.Replace what:="Finish", Replacement:="Forecast Finish", lookat:=xlWhole
  rng.Replace what:="Actual Start", Replacement:="New Forecast/ Actual Start", lookat:=xlWhole
  rng.Replace what:="Actual Finish", Replacement:="New Forecast/ Actual Finish", lookat:=xlWhole
  rng.Replace what:=strEVP, Replacement:="EV%", lookat:=xlWhole
  rng.Replace what:="Notes", Replacement:="Reason / Action / Impact", lookat:=xlWhole
  Debug.Print "rename headers: " & (GetTickCount - t) / 1000

  t = GetTickCount
  cptStatusSheet_frm.lblStatus.Caption = "Formatting Columns..."
  'columns to center
  Set aCentered = CreateObject("System.Collections.ArrayList")
  For Each vCol In Array("UID", "Duration", "Total Slack", strEVT, strEVP, "New EV%")
    aCentered.Add vCol
  Next vCol
  
  'entry headers
  Set aEntryHeaders = CreateObject("System.Collections.ArrayList")
  For Each vCol In Array("Actual Start", "Actual Finish", "New EV%", "Revised ETC", "Notes")
    aEntryHeaders.Add vCol
  Next vCol
  Debug.Print "define aCentered and aEntryHeaders: " & (GetTickCount - t) / 1000

  t = GetTickCount
  'define bulk ranges for formatting
  For lgCol = 0 To aHeaders.count - 1
    
    'format dates
    If Len(RegEx(CStr(aHeaders(lgCol)(1)), "Start|Finish")) > 0 Then
      If rDates Is Nothing Then
        Set rDates = xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow)
      Else
        Set rDates = xlApp.Union(rDates, xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow))
      End If
    End If
    'format work
    If Len(RegEx(CStr(aHeaders(lgCol)(1)), "Baseline Work|Remaining Work")) > 0 Then
      If rWork Is Nothing Then
        Set rWork = xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow)
      Else
        Set rWork = xlApp.Union(rWork, xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow))
      End If
    End If
    'format centered
    If aCentered.contains(aHeaders(lgCol)(1)) Then
      If rCentered Is Nothing Then
        Set rCentered = xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow)
      Else
        Set rCentered = xlApp.Union(rCentered, xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow))
      End If
    End If
    'format entry headers and columns
    If aEntryHeaders.contains(aHeaders(lgCol)(1)) Then
      If rEntry Is Nothing Then
        Set rEntry = xlCells(lgHeaderRow, lgCol + 1)
        Set rMedium = xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow)
        Set rLockedCells = rMedium
      Else
        Set rEntry = xlApp.Union(rEntry, xlCells(lgHeaderRow, lgCol + 1))
        Set rMedium = xlApp.Union(rMedium, xlCells(lgHeaderRow + 1, lgCol + 1).Resize(rowsize:=lgRow - lgHeaderRow))
        Set rLockedCells = rMedium
      End If
    End If
    
  Next
  Debug.Print "define bulk ranges for formatting: " & (GetTickCount - t) / 1000
  
  t = GetTickCount
  'apply bulk formatting
  rDates.NumberFormat = "m/d/yy;@"
  rDates.HorizontalAlignment = xlCenter
  rDates.Replace "NA", ""
  rWork.Style = "Comma"
  rCentered.HorizontalAlignment = xlCenter
  rEntry.Interior.ThemeColor = xlThemeColorAccent3
  rEntry.Interior.TintAndShade = 0.399975585192419
  rEntry.Font.ColorIndex = xlAutomatic
  rEntry.BorderAround xlContinuous, xlMedium
  rLockedCells.SpecialCells(xlCellTypeBlanks).Locked = False
  rMedium.BorderAround xlContinuous, xlMedium
  lgCol = Worksheet.Rows(lgHeaderRow).Find("Actual Finish", lookat:=xlPart).Column
  xlCells(lgHeaderRow + 1, lgCol).Resize(lgRow - lgHeaderRow).Borders(xlEdgeLeft).Weight = xlThin
  Debug.Print "apply bulk formatting: " & (GetTickCount - t) / 1000
  
  'todo: apply conditional formatting
  'update required formatting ("neutral"): - update required
'  .Font.Color = -16754788
'  .Font.TintAndShade = 0
'  .Interior.PatternColorIndex = xlAutomatic
'  .Interior.Color = 10284031
'  .Interior.TintAndShade = 0

  'invalid required formatting ("bad"): - invalid
'  .Font.Color = -16383844
'  .Font.TintAndShade = 0
'  .Interior.PatternColorIndex = xlAutomatic
'  .Interior.Color = 13551615
'  .Interior.TintAndShade = 0

  'updated ("good"): - user entered valid update
'  .Font.Color = -16752384
'  .Font.TintAndShade = 0
'  .Interior.PatternColorIndex = xlAutomatic
'  .Interior.Color = 13561798
'  .Interior.TintAndShade = 0  t = GetTickCount
  
  'define range for new start
  xlCells(lgHeaderRow, 1).AutoFilter
  Set rngAll = Worksheet.Range(xlCells(lgHeaderRow, 1).End(xlToRight), xlCells(lgHeaderRow, 1).End(xlDown))
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lgNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AS]
  rngAll.AutoFilter Field:=lgASCol, Criteria1:="="
  'add conditions only to blank cells in the column
  Set rng = Worksheet.Range(xlCells(lgHeaderRow + 1, lgASCol), xlCells(lgRow, lgASCol)).SpecialCells(xlCellTypeVisible)
  strFirstCell = rng(1).Address(False, True)
  '-->condition 1: blank and start is less than status date > update required
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=Indirect(""STATUS_DATE"")),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = 0.799981688894314
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  '-->condition 2: blank and EV% > 0 > invalid
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lgEVPCol).Address(False, True) & ">0),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  'greater than actual finish > invalid
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & ">" & xlCells(rng(1).Row, lgAFCol).Address(False, True) & "," & xlCells(rng(1).Row, lgAFCol).Address(False, True) & "<>""""),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  'else: <> start > updated
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & "<>" & xlCells(rng(1).Row, lgASCol - 2).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  
  'new finish
  Worksheet.ShowAllData
  xlCells(lgHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lgNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lgAFCol, Criteria1:="="
  'add conditions only to blank cells in the column
  Set rng = Worksheet.Range(xlCells(lgHeaderRow + 1, lgAFCol), xlCells(lgRow, lgAFCol)).SpecialCells(xlCellTypeVisible)
  strFirstCell = rng(1).Address(False, True)
  'blank and finish is less than status date > update required
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lgAFCol - 2).Address(False, True) & "<Indirect(""STATUS_DATE"")),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = 0.799981688894314
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  'less than actual start -> invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & xlCells(rng(1).Row, lgASCol).Address(False, True) & "<>""""," & strFirstCell & "<" & xlCells(rng(1).Row, lgASCol).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  'blank and EV% = 100 > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lgEVPCol).Address(False, True) & "=100),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  'else: <> finish > updated
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & "<>" & xlCells(rng(1).Row, lgAFCol - 2).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  
  'ev%
  Worksheet.ShowAllData
  xlCells(lgHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lgNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lgEVPCol, Criteria1:="="
  'add conditions only to blank cells in the column
  Set rng = Worksheet.Range(xlCells(lgHeaderRow + 1, lgEVPCol), xlCells(lgRow, lgEVPCol)).SpecialCells(xlCellTypeVisible)
  strFirstCell = rng(1).Address(False, True)
  'Finish < Status Date AND EV% < 100 (complete but incomplete) > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<100," & xlCells(rng(1).Row, lgAFCol).Address(False, True) & "<=Indirect(""STATUS_DATE"")),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = -0.499984740745262
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = 0.799981688894314
  End With
  'EV% > 0 and new start = "" (bogus actuals) > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & ">0," & xlCells(rng(1).Row, lgASCol).Address(False, True) & "=""""),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  'EV% =100 and new finish = "" (update required) > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=100," & xlCells(rng(1).Row, lgAFCol).Address(False, True) & "=""""),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  '=100 and new finish > status date > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=100," & xlCells(rng(1).Row, lgAFCol).Address(False, True) & ">Indirect(""STATUS_DATE"")),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  
  '(new start <> "" AND new start <> start) OR (newn finish <> "" AND new finish <> finish) (update required) > update required
  '<skipped>
  
  'revised etc
  Worksheet.ShowAllData
  xlCells(lgHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lgETCCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lgEVPCol, Criteria1:="="
  'add conditions only to blank cells in the column
  Set rng = Worksheet.Range(xlCells(lgHeaderRow + 1, lgEVPCol), xlCells(lgRow, lgEVPCol)).SpecialCells(xlCellTypeVisible)
  strFirstCell = rng(1).Address(False, True)
  'if assignments.count > 0
  'filter for assignments
'      ActiveSheet.Range("$A$8:$U$100").AutoFilter Field:=6, Operator:= _
'        xlFilterAutomaticFontColor
'    ActiveSheet.Range("$A$8:$U$100").AutoFilter Field:=7, Operator:= _
'        xlFilterNoFill
  
  '>0 and ev%=100 (complete with etc) > invalid
  '>0 and finish < status date (complete with etc) > invalid
  '=0 and ev%<100 (incpmlete without etc) > invalid
  '=0 and finish > status date (incomplete without etc) > invalid
  '(new start <> "" AND new start <> start) OR (newn finish <> "" AND new finish <> finish) (update required) > update required
  
  'evt vs evp checks
  If cptStatusSheet_frm.cboCostTool = "COBRA" Then
    'EVT = E 50/50
    'EVT = F 0/100
  ElseIf cptStatusSheet_frm.cboCostTool.Value = "MPM" Then
    'EVT =1 AND EVP NOT 0 OR 100
    'EVT =4
      'AS AND EVP NOT 50
      'AF AND EVP NOT 100
  Else
    'skip it - too many variables
  End If
  Debug.Print "apply conditional formatting " & (GetTickCount - t) / 1000
  
  xlApp.Visible = True
  xlApp.ScreenUpdating = True
  
  Worksheet.ShowAllData
  xlCells(lgHeaderRow + 1, lgNameCol + 1).Select
  xlApp.ActiveWindow.FreezePanes = True
  'prettify the task name column
  Worksheet.Columns(lgNameCol).AutoFit
  
  t = GetTickCount
  cptStatusSheet_frm.lblStatus.Caption = "Saving Workbook" & IIf(cptStatusSheet_frm.optWorkbooks, "s", "") & "..."
  'todo:save the workbook, worksheets, or workbooks
  strDir = Environ("USERPROFILE") & "\Desktop\CP_Status_Sheets\"
  'get clean project name
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  strDir = strDir & Format(dtStatus, "yyyy-mm-dd")
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  strFileName = RemoveIllegalCharacters(ActiveProject.Name)
  strFileName = Replace(strFileName, ".mpp", "")
  'create folder on desktop for project(?)
  'create folder on desktop for status date
  If cptStatusSheet_frm.optWorkbook Then
    'protect the sheet
    Worksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    Worksheet.EnableSelection = xlNoRestrictions
    'save to desktop in folder for status date
    strFileName = strFileName & "_StatusSheet_" & Format(dtStatus, "yyyy-mm-dd") & ".xlsx"
    If Dir(strDir & "\" & strFileName) <> vbNullString Then Kill strDir & "\" & strFileName
    Workbook.SaveAs strDir & "\" & strFileName, 51
  Else
    'rename master, apply autofilter
    'cycle through each option and create sheet
    Worksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    Worksheet.EnableSelection = xlNoRestrictions
    If cptStatusSheet_frm.optWorksheets Then
      'save to desktop in folder for status date
    ElseIf cptStatusSheet_frm.optWorkbooks Then
      'cycle through each worksheet and create workbook
      Worksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
      Worksheet.EnableSelection = xlNoRestrictions
      'save to desktop in folder for status date
    End If
  End If
  Debug.Print "save workbook: " & (GetTickCount - t) / 1000
  Debug.Print "</=====PERFORMANCE TEST=====>"
  
  cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
  cptStatusSheet_frm.lblStatus.Caption = " Complete."
  Application.StatusBar = "Complete."
  
  MsgBox "Status Sheet Created", vbInformation + vbOKOnly, "ClearPlan Status Sheet"
  xlApp.Visible = True
  cptStatusSheet_frm.lblStatus.Caption = " Ready..."

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  
  xlApp.Calculation = xlCalculationAutomatic
  xlApp.ScreenUpdating = True
  Set Tasks = Nothing
  Set Task = Nothing
  Set Resource = Nothing
  Set Assignment = Nothing
  Set xlApp = Nothing
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set rng = Nothing
  Set rSummaryTasks = Nothing
  Set rLockedCells = Nothing
  Set rMilestones = Nothing
  Set rNormal = Nothing
  Set rAssignments = Nothing
  Set rDates = Nothing
  Set rWork = Nothing
  Set rMedium = Nothing
  Set rCentered = Nothing
  Set rEntry = Nothing
  Set aEach = Nothing
  Set aTaskRow = Nothing
  Set aHeaders = Nothing
  Set aOddBalls = Nothing
  Set aCentered = Nothing
  Set aEntryHeaders = Nothing
  Set xlCells = Nothing
  Exit Sub

err_here:
  Call HandleErr("cptStatusSheet_bas", "CreateStatusSheet", err)
  If Not xlApp Is Nothing Then
    If Not Workbook Is Nothing Then Workbook.Close False
    xlApp.Quit
  End If
  Resume exit_here

End Sub