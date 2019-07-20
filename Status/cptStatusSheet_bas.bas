Attribute VB_Name = "cptStatusSheet_bas"
'<cpt_version>v1.0.9</cpt_version>
Option Explicit
#If Win64 And VBA7 Then '<issue53>
  Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr '<issue53>
#Else '<issue53>
  Declare Function GetTickCount Lib "kernel32" () As Long
#End If '<issue53>
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
Dim lngField As Long, lngItem As Long
'integers
Dim intField As Integer
'strings
Dim strFieldName As String, strFileName As String
'dates
Dim dtStatus As Date
'variants
Dim vFieldType As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'requires ms excel
  If Not cptCheckReference("Excel") Then GoTo exit_here
  'requires scripting (cptRegEx)
  If Not cptCheckReference("Scripting") Then GoTo exit_here

  With cptStatusSheet_frm
    .lboFields.Clear
    .lboExport.Clear
    .cboEVT.Clear
    .cboEVP.Clear
    .cboEVP.AddItem "Physical % Complete"
    .cboCostTool.Clear
    .cboCostTool.AddItem "COBRA"
    .cboCostTool.AddItem "MPM"
    .cboCostTool.AddItem "<none>"
  End With

  Set arrFields = CreateObject("System.Collections.SortedList")
  Set arrEVT = CreateObject("System.Collections.SortedList")
  Set arrEVP = CreateObject("System.Collections.SortedList")

  For Each vFieldType In Array("Text", "Outline Code", "Number")
    On Error GoTo err_here
    For intField = 1 To 30
      lngField = FieldNameToFieldConstant(vFieldType & intField, pjTask)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        arrFields.Add strFieldName, lngField
        If vFieldType = "Text" Then
          arrEVT.Add strFieldName, lngField
          'todo: what if this is an enterprise field?
        ElseIf vFieldType = "Number" Then
          arrEVP.Add strFieldName, lngField
          'todo: what if this is an enterprise field?
        End If
      End If
next_field:
    Next intField
  Next vFieldType

  'get enterprise custom fields
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      arrFields.Add Application.FieldConstantToFieldName(lngField), lngField
    End If
  Next lngField

  'add custom fields
  'col0 = constant
  'col1 = name
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
      lngItem = 0
      Do While Not .EOF
        cptStatusSheet_frm.lboExport.AddItem
        cptStatusSheet_frm.lboExport.List(lngItem, 0) = .Fields(0) 'Field Constant
        cptStatusSheet_frm.lboExport.List(lngItem, 1) = .Fields(1) 'Custom Field Name
        cptStatusSheet_frm.lboExport.List(lngItem, 2) = .Fields(2) 'Local Field Name
        lngItem = lngItem + 1
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
  cptStatusSheet_frm.show False

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
    Call cptHandleErr("cptStatusSheet_frm", "ShowCptStatusSheet_frm", err, Erl)
    Resume exit_here
  End If

End Sub

Sub cptCreateStatusSheet()
'objects
Dim rCompleted As Object
Dim aCompleted As Object
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
Dim lngDayLabelDisplay As Long
Dim lngTaskRow As Long
Dim lngLastRow As Long
Dim lngDateFormat As Long
Dim lngTaskCount As Long, lngTask As Long, lngHeaderRow As Long
Dim lngRow As Long, lngCol As Long, lngField As Long
Dim lngNameCol As Long, lngBaselineWorkCol As Long, lngRemainingWorkCol As Long, lngEach As Long
Dim lngNotesCol As Long, lngColumnWidth As Long
Dim lngASCol As Long, lngAFCol As Long, lngETCCol As Long, lngEVPCol As Long
#If Win64 and VBA7 Then '<issue53>
	Dim t As LongPtr, tTotal As LongPtr '<issue53>
#Else '<issue53>
	Dim t As Long, tTotal As Long '<issue53>
#End If '<issue53>
'strings
Dim strStatusDate As String
Dim strMsg As String
Dim strEVT As String, strEVP As String, strDir As String, strFileName As String
Dim strFirstCell As String
'dates
Dim dtStatus As Date
'variants
Dim vCol As Variant, aUserFields As Variant
'booleans
Dim blnPerformanceTest As Boolean
Dim blnSpace As Boolean
Dim blnFast As Boolean

  tTotal = GetTickCount

  'check reference
  If Not cptCheckReference("Excel") Then GoTo exit_here

  'ensure required module exists
  If Not cptModuleExists("cptCore_bas") Then
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
  Application.StatusBar = "Analyzing project..."
  'get task count
  If blnPerformanceTest Then t = GetTickCount
  For Each Task In Tasks
    lngTaskCount = lngTaskCount + 1
  Next Task
  blnPerformanceTest = True
  If blnPerformanceTest Then Debug.Print "<=====PERFORMANCE TEST " & Now() & "=====>"
  If blnPerformanceTest Then Debug.Print "get task count: " & (GetTickCount - t) / 1000

  cptStatusSheet_frm.lblStatus.Caption = " Setting up workbook..."
  Application.StatusBar = "Setting up workbook..."
  'set up an excel workbook
  If blnPerformanceTest Then t = GetTickCount
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = xlCalculationManual
  xlApp.ScreenUpdating = False
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = "Status Sheet"
  Set xlCells = Worksheet.Cells
  If blnPerformanceTest Then Debug.Print "set up excel workbook: " & (GetTickCount - t) / 1000

  'set up legend
  If blnPerformanceTest Then t = GetTickCount
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
  xlCells(3, 1).Style = "Input" '<issue58>
  'xlCells(3, 1).Interior.ThemeColor = xlThemeColorAccent2 '<issue58>
  'xlCells(3, 1).Interior.TintAndShade = 0.799981688894314 '<issue58>
  xlCells(3, 2) = "Task is active or within current status window. Cell requires update."
  'within two weeks
  xlCells(4, 1).Style = "Neutral" '<issue58>
  'xlCells(4, 1).Interior.ThemeColor = xlThemeColorAccent5 '<issue58>
  'xlCells(4, 1).Interior.TintAndShade = 0.799981688894314 '<issue58>
  xlCells(4, 1).BorderAround xlContinuous, xlThin, , -8421505
  xlCells(4, 2) = "Task is within two week look-ahead. Please review forecast dates."
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
  If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000

  lngHeaderRow = 8

  'set up header
  If blnPerformanceTest Then t = GetTickCount

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
      lngColumnWidth = aOddBalls.Item(vCol)
    Else
      lngColumnWidth = 10 'default
    End If
    aHeaders.Add Array(FieldNameToFieldConstant(vCol), vCol, lngColumnWidth)
  Next vCol

  'save fields to adtg file
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  aUserFields = cptStatusSheet_frm.lboExport.List()
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", adVarChar, 255
    .Fields.Append "Custom Field Name", adVarChar, 255
    .Fields.Append "Local Field Name", adVarChar, 255
    .Open
    For lngField = 0 To UBound(aUserFields)
      .AddNew Array(0, 1, 2), Array(aUserFields(lngField, 0), aUserFields(lngField, 1), aUserFields(lngField, 2))
    Next lngField
    '<issue43> capture case when no custom fields are selected
    If cptStatusSheet_frm.lboExport.ListCount > 0 Then
      .Update
     If Dir(strFileName) <> vbNullString Then Kill strFileName
     .Save strFileName
    End If '</issue43>
    .Close
  End With

  'get user fields
  For lngField = UBound(aUserFields) To 0 Step -1
    If aUserFields(lngField, 1) = strEVT Then GoTo next_field
    If aUserFields(lngField, 1) = strEVP Then GoTo next_field
    If lngField = 0 Then lngColumnWidth = 16 Else lngColumnWidth = 10
    aHeaders.Insert 1, Array(aUserFields(lngField, 0), aUserFields(lngField, 1), lngColumnWidth)
next_field:
  Next lngField

  'write the headers and size the columns
  For lngField = 0 To aHeaders.count - 1
    xlCells(lngHeaderRow, lngField + 1).Value = aHeaders(lngField)(1)
    xlCells(lngHeaderRow, lngField + 1).EntireColumn.ColumnWidth = aHeaders(lngField)(2)
    'get columns needed later
    If aHeaders(lngField)(1) = "Name" Then lngNameCol = lngField + 1
    If aHeaders(lngField)(1) = "Actual Start" Then lngASCol = lngField + 1
    If aHeaders(lngField)(1) = "Actual Finish" Then lngAFCol = lngField + 1
    If aHeaders(lngField)(1) = "Baseline Work" Then lngBaselineWorkCol = lngField + 1
    If aHeaders(lngField)(1) = "Remaining Work" Then lngRemainingWorkCol = lngField + 1
    If aHeaders(lngField)(1) = "Notes" Then lngNotesCol = lngField + 1
  Next

  'format the header row
  With xlCells(lngHeaderRow, 1).Resize(, aHeaders.count)
    .Interior.ThemeColor = xlThemeColorLight2
    .Interior.TintAndShade = 0
    .Font.ThemeColor = xlThemeColorDark1
    .Font.TintAndShade = 0
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
  End With
  If blnPerformanceTest Then Debug.Print "set up header: " & (GetTickCount - t) / 1000

  'prepare to capture each
  If cptStatusSheet_frm.optWorkbook = False Then
    Set aEach = CreateObject("System.Collections.SortedList")
    lngEach = FieldNameToFieldConstant(cptStatusSheet_frm.cboEach.Value)
  End If

  'prepare to capture bulk ranges
  Set aTaskRow = CreateObject("System.Collections.ArrayList")
  Set aAssignments = CreateObject("System.Collections.ArrayList")
  Set aSummaries = CreateObject("System.Collections.ArrayList")
  Set aMilestones = CreateObject("System.Collections.ArrayList")
  Set aNormal = CreateObject("System.Collections.ArrayList")
  Set aCompleted = CreateObject("System.Collections.ArrayList") '<issue58>

  'set the date and duration formats '<issue58>
  lngDateFormat = Application.DefaultDateFormat
  Application.DefaultDateFormat = pjDate_mm_dd_yyyy
  blnSpace = ActiveProject.SpaceBeforeTimeLabels
  ActiveProject.SpaceBeforeTimeLabels = False
  lngDayLabelDisplay = ActiveProject.DayLabelDisplay
  ActiveProject.DayLabelDisplay = 0

  'capture task data
  If blnPerformanceTest Then t = GetTickCount
  lngRow = lngHeaderRow
  For Each Task In Tasks
    If Task Is Nothing Then GoTo next_task
    If Task.OutlineLevel = 0 Then GoTo next_task
    If Task.ExternalTask Then GoTo next_task
    If Not Task.Active Then GoTo next_task
    If cptStatusSheet_frm.chkHide = True Then
      If Task.ActualFinish <= CDate(cptStatusSheet_frm.txtHideCompleteBefore) Then GoTo next_task
    End If

    lngRow = lngRow + 1

    If cptStatusSheet_frm.optWorkbook = False Then
      aEach.Add Task.GetField(lngEach), Task.GetField(lngEach)
    End If

    'get common data
    For lngCol = 1 To lngNameCol
      aTaskRow.Add Task.GetField(aHeaders(lngCol - 1)(0))
    Next lngCol

    'indent the task name
    xlCells(lngRow, lngNameCol).IndentLevel = Task.OutlineLevel + 1

    'write to worksheet
    If Task.Summary Then
      xlCells(lngRow, 1).Resize(, aTaskRow.count).Value = aTaskRow.ToArray()
      aTaskRow.Clear
      aSummaries.Add lngRow
    Else
      For lngCol = lngNameCol + 1 To aHeaders.count
        If aHeaders(lngCol - 1)(1) = "Baseline Work" Then
          aTaskRow.Add Task.BaselineWork / 60
        ElseIf aHeaders(lngCol - 1)(1) = "Remaining Work" Then
          aTaskRow.Add Task.RemainingWork / 60
        Else
          aTaskRow.Add Task.GetField(aHeaders(lngCol - 1)(0))
        End If
      Next lngCol

      'identify for formatting
      If Task.Milestone Then aMilestones.Add lngRow Else aNormal.Add lngRow

      'write task data to sheet
      xlCells(lngRow, 1).Resize(, aTaskRow.count).Value = aTaskRow.ToArray()
      aTaskRow.Clear

      'get assignment data for incomplete tasks
      If Task.ActualFinish = "NA" Then
        'add remaining work formula '<issue58>
        If Task.Assignments.count > 0 Then
          xlCells(lngRow, lngRemainingWorkCol).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C:R" & lngRow + Task.Assignments.count & "C)"
        End If
        'get assignment data
        For Each Assignment In Task.Assignments
          lngRow = lngRow + 1
          aTaskRow.Add Assignment.UniqueID
          If lngNameCol > 2 Then
            For lngCol = 2 To lngNameCol - 1
              aTaskRow.Add Task.GetField(aHeaders(lngCol - 1)(0))
            Next
          End If

          'identify for formatting
          aAssignments.Add lngRow

          'xlCells(lngRow, lngNameCol).Value = Assignment.ResourceName
          aTaskRow.Add Assignment.ResourceName
          xlCells(lngRow, lngNameCol).IndentLevel = Task.OutlineLevel + 2
          xlCells(lngRow, lngBaselineWorkCol).Value = Assignment.BaselineWork / 60
          xlCells(lngRow, lngRemainingWorkCol).Value = Assignment.RemainingWork / 60
          xlCells(lngRow, 1).Resize(, aTaskRow.count).Value = aTaskRow.ToArray()
          aTaskRow.Clear

        Next Assignment
      Else 'task is complete '<issue58>
        aCompleted.Add lngRow '<issue58>

      End If 'Task.ActualFinish = "NA"

    End If 'Task Summary

next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Exporting..." & Format(lngTask, "#,##0") & " / " & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ")"
    cptStatusSheet_frm.lblStatus.Caption = " Exporting..." & Format(lngTask, "#,##0") & " / " & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ")"
    cptStatusSheet_frm.lblProgress.Width = (lngTask / (lngTaskCount)) * cptStatusSheet_frm.lblStatus.Width
  Next Task

  If blnPerformanceTest Then Debug.Print "capture task data: " & (GetTickCount - t) / 1000 & " >> " & Format(((GetTickCount - t) / 1000) / (lngRow - lngHeaderRow), "#0.00000") & " per task"

  If blnPerformanceTest Then t = GetTickCount
  'add New EV% after EV% - update aHeaders
  lngEVPCol = Worksheet.Rows(lngHeaderRow).Find(strEVP).Column + 1
  Worksheet.Columns(lngEVPCol - 1).Copy
  Worksheet.Columns(lngEVPCol).Insert Shift:=xlToRight
  Worksheet.Range(xlCells(lngHeaderRow + 1, lngEVPCol), xlCells(lngRow, lngEVPCol)).Cells.Locked = False
  xlCells(lngHeaderRow, lngEVPCol).Value = "New EV%"
  aHeaders.Insert lngEVPCol - 1, Array(0, "New EV%", 10)

  'add Revised ETC after Remaining Work - update aHeaders
  lngRemainingWorkCol = Worksheet.Rows(lngHeaderRow).Find("Remaining Work", lookat:=xlWhole).Column 'don't use lngRemainingWorkCol because we've added a new column (and might add more)
  lngETCCol = lngRemainingWorkCol + 1
  Worksheet.Columns(lngETCCol).Insert Shift:=xlToRight
  Worksheet.Range(xlCells(lngHeaderRow, lngRemainingWorkCol), xlCells(lngRow, lngRemainingWorkCol)).Copy
  Worksheet.Range(xlCells(lngHeaderRow, lngETCCol), xlCells(lngRow, lngETCCol)).PasteSpecial xlAll
  Worksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).Style = "Comma"
  Worksheet.Columns(lngETCCol).ColumnWidth = 10
  xlCells(lngHeaderRow, lngETCCol).Value = "Revised ETC"
  Worksheet.Calculate 'trigger Remaining Work formula to calculate
  Worksheet.Range(xlCells(lngHeaderRow, lngRemainingWorkCol), xlCells(lngRow, lngRemainingWorkCol)).Copy
  xlCells(lngHeaderRow, lngRemainingWorkCol).PasteSpecial xlValues
  aHeaders.Insert lngETCCol - 1, Array(0, "Revised ETC", 10)
  If blnPerformanceTest Then Debug.Print "add columns: " & (GetTickCount - t) / 1000

  cptStatusSheet_frm.lblStatus = " Formatting rows..."
  Application.StatusBar = "Formatting rows..."
  If blnPerformanceTest Then t = GetTickCount
  'format rows
  'format summary tasks
  If aSummaries.count > 0 Then '<issue16-17> added
    Set rSummaryTasks = xlCells(aSummaries(0), 1).Resize(, aHeaders.count)
    For vCol = 1 To aSummaries.count - 1
      Set rSummaryTasks = xlApp.Union(rSummaryTasks, xlCells(aSummaries(vCol), 1).Resize(, aHeaders.count))
    Next vCol
    If Not rSummaryTasks Is Nothing Then
      rSummaryTasks.Interior.ThemeColor = xlThemeColorDark1
      rSummaryTasks.Interior.TintAndShade = -0.149998474074526
      rSummaryTasks.Font.Bold = True
    End If
  End If '</issue16-17>
  'format milestones
  If aMilestones.count > 0 Then '<issue16-17> added
    Set rMilestones = xlCells(aMilestones(0), 1).Resize(, aHeaders.count)
    For vCol = 1 To aMilestones.count - 1
      Set rMilestones = xlApp.Union(rMilestones, xlCells(aMilestones(vCol), 1).Resize(, aHeaders.count))
    Next vCol
    If Not rMilestones Is Nothing Then
      rMilestones.Font.ThemeColor = xlThemeColorAccent6
      rMilestones.Font.TintAndShade = -0.249977111117893
    End If
  End If '</issue16-17>
  'format normal
  If aNormal.count > 0 Then '<issue16-17> added
    Set rNormal = xlCells(aNormal(0), 1).Resize(, aHeaders.count)
    For vCol = 1 To aNormal.count - 1
      Set rNormal = xlApp.Union(rNormal, xlCells(aNormal(vCol), 1).Resize(, aHeaders.count))
    Next vCol
    If Not rNormal Is Nothing Then
      rNormal.Font.Color = RGB(32, 55, 100) '<issue54>
    End If
  End If '</issue16-17>
  'format assignments
  If aAssignments.count > 0 Then '<issue16-17> added
    Set rAssignments = xlCells(aAssignments(0), 1).Resize(, aHeaders.count)
    For vCol = 1 To aAssignments.count - 1
      Set rAssignments = xlApp.Union(rAssignments, xlCells(aAssignments(vCol), 1).Resize(, aHeaders.count))
    Next vCol
    If Not rAssignments Is Nothing Then rAssignments.Font.Italic = True
  End If '</issue16-17>
  'format completed '<issue58>
  If aCompleted.count > 0 Then
    'format the entire row - assignments are skipped on completd tasks
    Set rCompleted = xlCells(aCompleted(0), 1).Resize(, aHeaders.count)
    For vCol = 1 To aCompleted.count - 1
      Set rCompleted = xlApp.Union(rCompleted, xlCells(aCompleted(vCol), 1).Resize(, aHeaders.count))
    Next vCol
    If Not rCompleted Is Nothing Then
      rCompleted.Font.Italic = True
      rCompleted.Font.ColorIndex = 16
    End If
    'update ev% complete
    lngCol = Worksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
    Set rCompleted = xlCells(aCompleted(0), lngCol)
    For vCol = 1 To aCompleted.count - 1 'we are borrowing vCol to iterate row numbers
      Set rCompleted = xlApp.Union(rCompleted, xlCells(aCompleted(vCol), lngCol))
    Next vCol
    rCompleted = 1
  End If '</issue58>
  If blnPerformanceTest Then Debug.Print "format rows: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'format common borders
  Set rng = Worksheet.Range(xlCells(lngHeaderRow, 1), xlCells(lngRow, aHeaders.count))
  rng.BorderAround xlContinuous, xlThin
  rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
  rng.Borders(xlInsideHorizontal).Weight = xlThin
  If blnPerformanceTest Then Debug.Print "format common borders: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'rename headers
  Set rng = xlCells(lngHeaderRow, 1).Resize(, aHeaders.count)
  rng.Replace what:="Unique ID", Replacement:="UID", lookat:=xlWhole
  rng.Replace what:="Name", Replacement:="Task Name", lookat:=xlWhole
  rng.Replace what:="Start", Replacement:="Forecast Start", lookat:=xlWhole
  rng.Replace what:="Finish", Replacement:="Forecast Finish", lookat:=xlWhole
  rng.Replace what:="Actual Start", Replacement:="New Forecast/ Actual Start", lookat:=xlWhole
  rng.Replace what:="Actual Finish", Replacement:="New Forecast/ Actual Finish", lookat:=xlWhole
  rng.Replace what:=strEVP, Replacement:="EV%", lookat:=xlWhole
  rng.Replace what:="Notes", Replacement:="Reason / Action / Impact", lookat:=xlWhole
  If blnPerformanceTest Then Debug.Print "rename headers: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  cptStatusSheet_frm.lblStatus.Caption = " Formatting columns..."
  Application.StatusBar = "Formatting Columns..."

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
  If blnPerformanceTest Then Debug.Print "define aCentered and aEntryHeaders: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'define bulk column ranges for formatting
  For lngCol = 0 To aHeaders.count - 1

    'format dates
    If Len(cptRegEx(CStr(aHeaders(lngCol)(1)), "Start|Finish")) > 0 Then
      If rDates Is Nothing Then
        Set rDates = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rDates = xlApp.Union(rDates, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'format work
    If Len(cptRegEx(CStr(aHeaders(lngCol)(1)), "Baseline Work|Remaining Work")) > 0 Then
      If rWork Is Nothing Then
        Set rWork = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rWork = xlApp.Union(rWork, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'format centered
    If aCentered.contains(aHeaders(lngCol)(1)) Then
      If rCentered Is Nothing Then
        Set rCentered = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rCentered = xlApp.Union(rCentered, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'format entry headers and columns
    If aEntryHeaders.contains(aHeaders(lngCol)(1)) Then 'if the column we're working on is included in the list of entry headers, then...
      If rEntry Is Nothing Then 'first iteration sets range
        Set rEntry = xlCells(lngHeaderRow, lngCol + 1)
        Set rMedium = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow) 'medium = border thickness
        Set rLockedCells = rMedium 'entry cells are unlocked cells
      Else 'second and following iterations extend the range
        Set rEntry = xlApp.Union(rEntry, xlCells(lngHeaderRow, lngCol + 1))
        Set rMedium = xlApp.Union(rMedium, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
        Set rLockedCells = rMedium
      End If
    End If

  Next
  If blnPerformanceTest Then Debug.Print "define bulk ranges for formatting: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
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
  lngCol = Worksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlPart).Column
  xlCells(lngHeaderRow + 1, lngCol).Resize(lngRow - lngHeaderRow).Borders(xlEdgeLeft).Weight = xlThin
  If blnPerformanceTest Then Debug.Print "apply bulk formatting: " & (GetTickCount - t) / 1000

  'apply conditional formatting
  'update required formatting ("input"): - update required
'  .Font.Color = 7749439
'  .Font.TintAndShade = 0
'  .Interior.PatternColorIndex = -4105
'  .Interior.Color = 10079487
'  .Interior.TintAndShade = 0
'  .BorderAround xlContinuous, xlThin, , Color:=RGB(127, 127, 127)

  'two week window ("neutral"): - review
'  .Font.Color = -16754788
'  .Font.TintAndShade = 0
'  .Interior.PatternColorIndex = xlAutomatic
'  .Interior.Color = 10284031
'  .Interior.TintAndShade = 0
'  .BorderAround xlContinuous, xlThin,,color:=RGB(127,127,127)

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

  cptStatusSheet_frm.lblStatus.Caption = " Applying conditional formats..."
  Application.StatusBar = "Applying conditional formats..."

  'capture status date address
  strStatusDate = Worksheet.Range("STATUS_DATE").Address(True, True)

new_start:
  'define range for new start
  xlCells(lngHeaderRow, 1).AutoFilter
  Set rngAll = Worksheet.Range(xlCells(lngHeaderRow, 1).End(xlToRight), xlCells(lngHeaderRow, 1).End(xlDown))
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AS]
  rngAll.AutoFilter Field:=lngASCol, Criteria1:="="
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52-no cells found>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngASCol), xlCells(lngRow, lngASCol)).SpecialCells(xlCellTypeVisible)
  If err.Number = 1004 Then 'no cells found
    err.Clear '<issue52>
    GoTo new_finish '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  strFirstCell = rng(1).Address(False, True)

  '-->condition 1: blank and start is less than status date > update required
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = 7749439
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = -4105
    .Color = 10079487
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  cptStatusSheet_frm.lblProgress.Width = (1 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 2: two-week-window                      '=IF($E50<=(INDIRECT("STATUS_DATE")+14),TRUE,FALSE)
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=(" & strStatusDate & "+14)),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16754788
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 10284031
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  cptStatusSheet_frm.lblProgress.Width = (2 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 3: blank and EV% > 0 > invalid
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lngEVPCol).Address(False, True) & ">0),TRUE,FALSE)"
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
  cptStatusSheet_frm.lblProgress.Width = (3 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 4: greater than actual finish > invalid
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & ">" & xlCells(rng(1).Row, lngAFCol).Address(False, True) & "," & xlCells(rng(1).Row, lngAFCol).Address(False, True) & "<>""""),TRUE,FALSE)"
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
  cptStatusSheet_frm.lblProgress.Width = (4 / 14) * cptStatusSheet_frm.lblStatus.Width

  'else: <> start > updated
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & "<>" & xlCells(rng(1).Row, lngASCol - 2).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  cptStatusSheet_frm.lblProgress.Width = (5 / 14) * cptStatusSheet_frm.lblStatus.Width

new_finish: '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  Worksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lngAFCol, Criteria1:="="
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngAFCol), xlCells(lngRow, lngAFCol)).SpecialCells(xlCellTypeVisible)
  If err.Number = 1004 Then '<issue52>
    err.Clear '<issue52>
    GoTo ev_percent '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  strFirstCell = rng(1).Address(False, True)

  '-->condition 1: blank and finish is less than status date > update required
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & rng(1).Offset(0, 4).Address(False, True) & "<1," & rng(1).Offset(0, -3).Address(False, True) & "<" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = 7749439
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = -4105
    .Color = 10079487
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  cptStatusSheet_frm.lblProgress.Width = (6 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 2: two-week-window
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=(" & strStatusDate & "+14)),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16754788
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 10284031
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  cptStatusSheet_frm.lblProgress.Width = (7 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 3: less than actual start -> invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & xlCells(rng(1).Row, lngASCol).Address(False, True) & "<>""""," & strFirstCell & "<" & xlCells(rng(1).Row, lngASCol).Address(False, True) & "),TRUE,FALSE)"
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
  cptStatusSheet_frm.lblProgress.Width = (8 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 4: blank and EV% = 100 > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lngEVPCol).Address(False, True) & "=100),TRUE,FALSE)"
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
  cptStatusSheet_frm.lblProgress.Width = (9 / 14) * cptStatusSheet_frm.lblStatus.Width

  'else: <> finish > updated
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & "<>" & xlCells(rng(1).Row, lngAFCol - 2).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  cptStatusSheet_frm.lblProgress.Width = (10 / 14) * cptStatusSheet_frm.lblStatus.Width

ev_percent:
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  Worksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52-noCellsFound>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngEVPCol), xlCells(lngRow, lngEVPCol)).SpecialCells(xlCellTypeVisible)
  If err.Number = 1004 Then '<issue52>
    err.Clear '<issue52>
    GoTo revised_etc '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  strFirstCell = rng(1).Address(False, True)

  '-->condition 1: Start < Status Date AND EV% < 100 > update required
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & rng(1).Offset(0, -1).Address(False, True) & "<1," & rng(1).Address(False, True) & "<1," & rng(1).Offset(0, -7).Address(False, True) & "<=" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = 7749439
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = -4105
    .Color = 10079487
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True
  cptStatusSheet_frm.lblProgress.Width = (11 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 2: EV% > 0 and new start = "" (bogus actuals) > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & ">0," & xlCells(rng(1).Row, lngASCol).Address(False, True) & "=""""),TRUE,FALSE)"
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
  cptStatusSheet_frm.lblProgress.Width = (12 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 3: EV% =100 and new finish = "" (update required) > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=100," & xlCells(rng(1).Row, lngAFCol).Address(False, True) & "=""""),TRUE,FALSE)"
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
  cptStatusSheet_frm.lblProgress.Width = (13 / 14) * cptStatusSheet_frm.lblStatus.Width

  '-->condition 4: =100 and new finish > status date > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=100," & xlCells(rng(1).Row, lngAFCol).Address(False, True) & ">" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  cptStatusSheet_frm.lblProgress.Width = (14 / 14) * cptStatusSheet_frm.lblStatus.Width

  '(new start <> "" AND new start <> start) OR (new finish <> "" AND new finish <> finish) (update required) > update required
  '<skipped>

revised_etc:
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  Worksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for Task
  rngAll.AutoFilter Field:=lngETCCol, Operator:=xlFilterAutomaticFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lngRemainingWorkCol, Operator:=xlFilterNoFill
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).SpecialCells(xlCellTypeVisible)
  If err.Number = 1004 Then '<issue52>
    err.Clear '<issue52>
    GoTo evt_vs_evp '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  rng.Cells.Locked = False
  strFirstCell = rng(1).Address(False, True)

  lngLastRow = lngRow

  For lngRow = lngHeaderRow + 1 To lngLastRow
    If xlCells(lngRow, lngETCCol).Font.Color = RGB(32, 55, 100) Then
      lngTaskRow = lngRow
    ElseIf xlCells(lngRow, lngETCCol).Font.Italic And xlCells(lngRow, lngETCCol).Font.Color = xlAutomaticFontColor Then
      Set rng = xlCells(lngRow, lngETCCol)
        '-->condition 1: Start < Status Date AND EV% < 100 > update required
        rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & xlCells(lngTaskRow, lngEVPCol).Address(True, True) & "<1," & xlCells(lngTaskRow, lngASCol - 2).Address(True, True) & "<=(INDIRECT(""STATUS_DATE""))),TRUE,FALSE)"
        With rng.FormatConditions(rng.FormatConditions.count).Font
          .Color = 7749439
          .TintAndShade = 0
        End With
        With rng.FormatConditions(rng.FormatConditions.count).Interior
          .PatternColorIndex = -4105
          .Color = 10079487
          .TintAndShade = 0
        End With
        rng.FormatConditions(rng.FormatConditions.count).StopIfTrue = True

    End If
  Next lngRow

  '>0 and ev%=100 (complete with etc) > invalid
  '>0 and finish < status date (complete with etc) > invalid
  '=0 and ev%<100 (incpmlete without etc) > invalid
  '=0 and finish > status date (incomplete without etc) > invalid
  '(new start <> "" AND new start <> start) OR (new finish <> "" AND new finish <> finish) (update required) > update required

evt_vs_evp:
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
  If blnPerformanceTest Then Debug.Print "apply conditional formatting " & (GetTickCount - t) / 1000

  xlApp.Visible = True
  xlApp.ScreenUpdating = True

  Worksheet.ShowAllData
  xlApp.ActiveWindow.ScrollRow = 1 '<issue54>
  xlCells(lngHeaderRow + 1, lngNameCol + 1).Select
  xlApp.ActiveWindow.FreezePanes = True
  'prettify the task name column
  Worksheet.Columns(lngNameCol).AutoFit

  If blnPerformanceTest Then t = GetTickCount
  cptStatusSheet_frm.lblStatus.Caption = "Saving Workbook" & IIf(cptStatusSheet_frm.optWorkbooks, "s", "") & "..."
  Application.StatusBar = "Saving Workbook" & IIf(cptStatusSheet_frm.optWorkbooks, "s", "") & "..."
  'todo:save the workbook, worksheets, or workbooks
  strDir = Environ("USERPROFILE") & "\Desktop\CP_Status_Sheets\"
  'get clean project name
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  strDir = strDir & Format(dtStatus, "yyyy-mm-dd")
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  strFileName = cptRemoveIllegalCharacters(ActiveProject.Name)
  strFileName = Replace(strFileName, ".mpp", "")
  'create folder on desktop for project(?)
  'create folder on desktop for status date
  xlApp.Calculation = xlAutomatic
  If cptStatusSheet_frm.optWorkbook Then
    'protect the sheet
    Worksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    Worksheet.EnableSelection = xlNoRestrictions
    'save to desktop in folder for status date
    strFileName = strFileName & "_StatusSheet_" & Format(dtStatus, "yyyy-mm-dd") & ".xlsx"
    On Error Resume Next
    If Dir(strDir & "\" & strFileName) <> vbNullString Then Kill strDir & "\" & strFileName
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    'account for if the file exists and is open in the background
    If Dir(strDir & "\" & strFileName) <> vbNullString Then 'delete failed, rename with timestamp
      strMsg = "'" & strFileName & "' already exists, and is likely open." & vbCrLf
      strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "-hh-nn-ss") & ".xlsx")
      strMsg = strMsg & "The file you are now creating will be named '" & strFileName & "'"
      MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
      Workbook.SaveAs strDir & "\" & strFileName, 51
    Else
      Workbook.SaveAs strDir & "\" & strFileName, 51
    End If
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
  If blnPerformanceTest Then Debug.Print "save workbook: " & (GetTickCount - t) / 1000
  If blnPerformanceTest Then Debug.Print "</=====PERFORMANCE TEST=====>"

  cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
  cptStatusSheet_frm.lblStatus.Caption = " Complete."
  Application.StatusBar = "Complete."
  xlApp.Visible = True
  'MsgBox "Status Sheet Created", vbInformation + vbOKOnly, "ClearPlan Status Sheet"
  cptStatusSheet_frm.lblStatus.Caption = " Ready..."

exit_here:
  On Error Resume Next
  Application.DefaultDateFormat = lngDateFormat
  ActiveProject.SpaceBeforeTimeLabels = blnSpace
  ActiveProject.DayLabelDisplay = lngDayLabelDisplay
  Set rCompleted = Nothing
  Set aCompleted = Nothing
  Application.StatusBar = ""
  cptSpeed False
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
  Call cptHandleErr("cptStatusSheet_bas", "cptCreateStatusSheet", err, Erl)
  If Not xlApp Is Nothing Then
    If Not Workbook Is Nothing Then Workbook.Close False
    xlApp.Quit
  End If
  Resume exit_here

End Sub
