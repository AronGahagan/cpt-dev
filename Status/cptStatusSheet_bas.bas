Attribute VB_Name = "cptStatusSheet_bas"
'<cpt_version>v1.2.11</cpt_version>
Option Explicit
#If Win64 And VBA7 Then '<issue53>
  Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr '<issue53>
#Else '<issue53>
  Declare Function GetTickCount Lib "kernel32" () As Long
#End If '<issue53>
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_heref Else On Error GoTo 0
Private Const adVarChar As Long = 200

Sub cptShowStatusSheet_frm()
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
Dim strFieldNamesChanged As String
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
  
  'reset options
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
    .cboCreate.AddItem
    For lngItem = 0 To 2
      .cboCreate.AddItem
      .cboCreate.List(lngItem, 0) = lngItem
      .cboCreate.List(lngItem, 1) = Choose(lngItem + 1, "A Single Workbook", "A Worksheet for each", "A Workbook for each")
    Next lngItem
    .chkSendEmails.Enabled = cptCheckReference("Outlook")
  End With

  'set up arrays to capture values
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
  
  'add Physical % Complete
  arrEVP.Add "Physical % Complete", FieldNameToFieldConstant("Physical % Complete")
  
  'add Contact field
  arrFields.Add "Contact", FieldNameToFieldConstant("Contact")
  
  'get enterprise custom fields
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      arrFields.Add Application.FieldConstantToFieldName(lngField), lngField
    End If
  Next lngField

  'add custom fields
  'col0 = constant
  'col1 = name
  For intField = 0 To arrFields.Count - 1
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
  For intField = 0 To arrEVT.Count - 1
    cptStatusSheet_frm.cboEVT.AddItem arrEVT.getKey(intField)
  Next
  'add EVP values
  For intField = 0 To arrEVP.Count - 1 'UBound(st)
    cptStatusSheet_frm.cboEVP.AddItem arrEVP.getKey(intField) 'st(intField)(1)
  Next

  'add saved settings if they exist
  strFileName = cptDir & "\settings\cpt-status-sheet.adtg"
  If Dir(strFileName) <> vbNullString Then
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      .MoveFirst
      
      On Error Resume Next
      lngField = FieldNameToFieldConstant(.Fields(0))
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      'auto-select if saved setting exists and if saved field exists in the comboBox
      If lngField > 0 And arrEVT.Contains(CStr(.Fields(0))) Then cptStatusSheet_frm.cboEVT.Value = .Fields(0) 'cboEVT
      lngField = 0
      
      On Error Resume Next
      lngField = FieldNameToFieldConstant(.Fields(1))
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      'auto-select if saved setting exists and if saved field exists in the comboBox
      If lngField > 0 And arrEVP.Contains(CStr(.Fields(1))) Then cptStatusSheet_frm.cboEVP.Value = .Fields(1) 'cboEVP
      lngField = 0
      
      cptStatusSheet_frm.cboCreate = .Fields(2) - 1 'cboCreate
      
      cptStatusSheet_frm.chkHide = .Fields(3) = 1 'chkHide
      
      If .Fields.Count >= 5 Then
        If Not IsNull(.Fields(4)) Then cptStatusSheet_frm.cboCostTool.Value = .Fields(4) 'cboCostTool
      End If
      If .Fields.Count >= 6 Then
        If Not IsNull(.Fields(5)) Then
          On Error Resume Next
          lngField = FieldNameToFieldConstant(.Fields(5))
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          If lngField > 0 Then cptStatusSheet_frm.cboEach.Value = .Fields(5) 'cboEach
          lngField = 0
        End If
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
        If CustomFieldGetName(.Fields(0)) <> CStr(.Fields(1)) Then
          strFieldNamesChanged = strFieldNamesChanged & .Fields(2) & " '" & .Fields(1) & "' is now "
          If Len(CustomFieldGetName(.Fields(0))) > 0 Then
            strFieldNamesChanged = strFieldNamesChanged & "'" & CustomFieldGetName(.Fields(0)) & "'" & vbCrLf
          Else
            strFieldNamesChanged = strFieldNamesChanged & "<unnamed>" & vbCrLf
          End If
        End If
        lngItem = lngItem + 1
        .MoveNext
      Loop
      .Close
    End With
  End If

  'reset the view
  If InStr(ActiveWindow.TopPane.View, "Gantt") = 0 Then
    If MsgBox("Current view must be changed for successful export.", vbInformation + vbOKCancel, "Incompatible View") = vbOK Then
      ViewApply "Gantt Chart"
    Else
      GoTo exit_here
    End If
  End If
  Call cptRefreshStatusTable
  FilterClear
  OptionsViewEx displaysummarytasks:=True
  OutlineShowAllTasks

  'set the status date / hide complete
  If ActiveProject.StatusDate = "NA" Then
    cptStatusSheet_frm.txtStatusDate.Value = FormatDateTime(DateAdd("d", 6 - Weekday(Now), Now), vbShortDate)
  Else
    cptStatusSheet_frm.txtStatusDate = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  End If
  dtStatus = CDate(cptStatusSheet_frm.txtStatusDate.Value)
  'default to one week prior to status date
  cptStatusSheet_frm.txtHideCompleteBefore.Value = DateAdd("d", -7, dtStatus)

  'delete pre-existing search file
  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName

  'set up the view/table/filter
  cptStatusSheet_frm.Show False
  cptRefreshStatusTable
  FilterClear
  OptionsViewEx displaysummarytasks:=True
  OutlineShowAllTasks

  If Len(strFieldNamesChanged) > 0 Then
    strFieldNamesChanged = "The following saved export field names have changed:" & vbCrLf & vbCrLf & strFieldNamesChanged
    strFieldNamesChanged = strFieldNamesChanged & vbCrLf & vbCrLf & "You may wish to remove them from the export list."
    MsgBox strFieldNamesChanged, vbInformation + vbOKOnly, "Saved Settings - Mismatches"
  End If

exit_here:
  On Error Resume Next
  Set arrFields = Nothing
  Set arrEVT = Nothing
  Set arrEVP = Nothing
  Exit Sub

err_here:
  If Err.Number = 1101 Or Err.Number = 1004 Then
    Err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptStatusSheet_frm", "cptShowStatusSheet_frm", Err, Erl)
    Resume exit_here
  End If

End Sub

Sub cptCreateStatusSheet()
'objects
Dim rCompleted As Object
Dim aCompleted As Object
Dim aGroups As Object
Dim rngKeep As Object
Dim Tasks As Tasks, Task As Task, Resource As Resource, Assignment As Assignment
'early binding:
'Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Excel.Range
'Dim rSummaryTasks As Excel.Range, rMilestones As Excel.Range, rNormal As Excel.Range, rAssignments As Excel.Range, rLockedCells As Excel.Range
'Dim rDates As Excel.Range, rWork As Excel.Range, rMedium As Excel.Range, rCentered As Excel.Range, rEntry As Excel.Range
'Dim xlCells As Excel.Range, rngAll As Excel.Range
'Dim olApp As Outlook.Application, MailItem As MailItem, objWord As Word.Application, objSel As Word.Selection, objETemp As Word.Template
'late binding:
Dim xlApp As Object, Workbook As Object, Worksheet As Object, rng As Object
Dim rSummaryTasks As Object, rMilestones As Object, rNormal As Object, rAssignments As Object, rLockedCells As Object
Dim rDates As Object, rWork As Object, rMedium As Object, rCentered As Object, rEntry As Object
Dim xlCells As Object, rngAll As Object
Dim olApp As Object, MailItem As Object, objDoc As Object, objWord As Object, objSel As Object, objETemp As Object
Dim aSummaries As Object, aMilestones As Object, aNormal As Object, aAssignments As Object
Dim aEach As Object, aTaskRow As Object, aHeaders As Object
Dim aOddBalls As Object, aCentered As Object, aEntryHeaders As Object
'longs
Dim lngFormatCondition As Long
Dim lngConditionalFormats As Long
Dim lngDayLabelDisplay As Long
Dim lngTaskRow As Long
Dim lngLastRow As Long
Dim lngDateFormat As Long
Dim lngGroups As Long
Dim lngLastCol As Long
Dim lngTaskCount As Long, lngTask As Long, lngHeaderRow As Long
Dim lngRow As Long, lngCol As Long, lngField As Long
Dim lngNameCol As Long, lngBaselineWorkCol As Long, lngRemainingWorkCol As Long, lngEach As Long
Dim lngNotesCol As Long, lngColumnWidth As Long
Dim lngASCol As Long, lngAFCol As Long, lngETCCol As Long, lngEVPCol As Long
#If Win64 And VBA7 Then '<issue53>
        Dim t As LongPtr, tTotal As LongPtr '<issue53>
#Else '<issue53>
        Dim t As Long, tTotal As Long '<issue53>
#End If '<issue53>
Dim lngItem As Long
'strings
Dim strStatusDate As String
Dim strCriteria As String
Dim strFieldName As String
Dim strMsg As String
Dim strEVT As String, strEVP As String, strDir As String, strFileName As String
Dim strFirstCell As String
Dim strItem As String
'dates
Dim dtStatus As Date
'variants
Dim vCol As Variant, aUserFields As Variant
'booleans
Dim blnAddConditionalFormats As Boolean
Dim blnPerformanceTest As Boolean
Dim blnSpace As Boolean
Dim blnEmail As Boolean

  'check reference
  If Not cptCheckReference("Excel") Then GoTo exit_here
  'If Not cptCheckReference("Outlook") Then GoTo exit_here '<issue50>

  'ensure required module exists
  If Not cptModuleExists("cptCore_bas") Then
    MsgBox "Please install the ClearPlan 'cptCore_bas' module.", vbExclamation + vbOKOnly, "Missing Module"
    GoTo exit_here
  End If

  'this boolean spits out a speed test to the immediate window
  blnPerformanceTest = False
  If blnPerformanceTest Then tTotal = GetTickCount

  On Error Resume Next
  Set Tasks = ActiveProject.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ensure project has tasks
  If Tasks Is Nothing Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "Create Status Sheet"
    GoTo exit_here
  End If

  cptStatusSheet_frm.lblStatus.Caption = " Analyzing project..."
  Application.StatusBar = "Analyzing project..."
  'get task count
  If blnPerformanceTest Then t = GetTickCount
  
  SelectAll
  Set Tasks = ActiveSelection.Tasks
  For Each Task In Tasks
    lngTaskCount = lngTaskCount + 1
  Next Task
  If blnPerformanceTest Then Debug.Print "<=====PERFORMANCE TEST " & Now() & "=====>"
  If blnPerformanceTest Then Debug.Print "get task count: " & (GetTickCount - t) / 1000

  cptStatusSheet_frm.lblStatus.Caption = " Setting up workbook..."
  Application.StatusBar = "Setting up workbook..."
  'set up an excel workbook
  If blnPerformanceTest Then t = GetTickCount
  Set xlApp = CreateObject("Excel.Application")
  '/=== debug ==\
  'xlApp.Visible = True
  '\=== debug ===/
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = xlCalculationManual
  'xlApp.ScreenUpdating = False
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
  xlCells(3, 2) = "Task is active or within current status window. Cell requires update."
  'within two weeks
  xlCells(4, 1).Style = "Neutral" '<issue58>
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
  With cptStatusSheet_frm
    If .cboCostTool.Value <> "<none>" Then strEVT = .cboEVT.Value Else strEVT = "SKIP" '<issue64>
    strEVP = .cboEVP.Value
  End With

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
    If aOddBalls.Contains(vCol) Then
      lngColumnWidth = aOddBalls.Item(vCol)
    Else
      lngColumnWidth = 10 'default
    End If
    If CStr(vCol) <> "SKIP" Then aHeaders.Add Array(FieldNameToFieldConstant(vCol), vCol, lngColumnWidth) '<issue64>
  Next vCol

  'save fields to adtg file
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  aUserFields = cptStatusSheet_frm.lboExport.List()
  '<issue43> capture case when no custom fields are selected
  If cptStatusSheet_frm.lboExport.ListCount > 0 Then
    With CreateObject("ADODB.Recordset")
      .Fields.Append "Field Constant", adVarChar, 255
      .Fields.Append "Custom Field Name", adVarChar, 255
      .Fields.Append "Local Field Name", adVarChar, 255
      .Open
      For lngField = 0 To UBound(aUserFields)
        .AddNew Array(0, 1, 2), Array(aUserFields(lngField, 0), aUserFields(lngField, 1), aUserFields(lngField, 2))
      Next lngField
        .Update
       If Dir(strFileName) <> vbNullString Then Kill strFileName
       .Save strFileName
      .Close
    End With
  Else
    If Dir(strFileName) <> vbNullString Then Kill strFileName
  End If '</issue43>

  'get user fields
  For lngField = UBound(aUserFields) To 0 Step -1
    If aUserFields(lngField, 1) = strEVT Then GoTo next_field
    If aUserFields(lngField, 1) = strEVP Then GoTo next_field
    If lngField = 0 Then lngColumnWidth = 16 Else lngColumnWidth = 10
    aHeaders.Insert 1, Array(aUserFields(lngField, 0), aUserFields(lngField, 1), lngColumnWidth)
next_field:
  Next lngField

  'write the headers and size the columns
  For lngField = 0 To aHeaders.Count - 1
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
  With xlCells(lngHeaderRow, 1).Resize(, aHeaders.Count)
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
  If cptStatusSheet_frm.cboCreate.Value <> "0" Then
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
  SelectAll
  Set Tasks = ActiveSelection.Tasks

  SelectBeginning
  Set aGroups = CreateObject("System.Collections.SortedList")
  Do
    On Error Resume Next
    Set Task = ActiveCell.Task
    If Err.Number > 0 Then
      Err.Number = 0
      Err.Clear
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      Exit Do
    End If
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

    If Task Is Nothing Then GoTo next_task
    If Task.OutlineLevel = 0 Then GoTo next_task
    If Task.ExternalTask Then GoTo next_task
    If Not Task.Active Then GoTo next_task
    If cptStatusSheet_frm.chkHide = True And Not Task.GroupBySummary And Not Task.Summary Then
      If Task.ActualFinish <= CDate(cptStatusSheet_frm.txtHideCompleteBefore) Then GoTo next_task
    End If

    lngRow = lngRow + 1
    
    'retain-grouping
    If Task.GroupBySummary Then
      aSummaries.Add lngRow 'capture for formatting
      aGroups.Add Task.Name, Task.UniqueID
      xlCells(lngRow, lngNameCol).IndentLevel = Task.OutlineLevel - 1
      xlCells(lngRow, 1).Value = Task.UniqueID
      xlCells(lngRow, lngNameCol).Value = Task.Name
      lngTaskCount = lngTaskCount + 1
      GoTo next_task
    End If
    
    'add unique list of values to aEach->lboEach in the selected field
    If cptStatusSheet_frm.cboCreate.Value <> "0" Then
      If Not aEach.Contains(Task.GetField(lngEach)) Then
        If Len(Task.GetField(lngEach)) > 0 And Not Task.Summary Then
          aEach.Add Task.GetField(lngEach), Task.GetField(lngEach)
        End If
      End If
    End If

    'get common data
    For lngCol = 1 To lngNameCol
      aTaskRow.Add Task.GetField(aHeaders(lngCol - 1)(0))
    Next lngCol

    'indent the task name
    xlCells(lngRow, lngNameCol).IndentLevel = Task.OutlineLevel + 1

    'write to worksheet
    If Task.Summary Then
      xlCells(lngRow, 1).Resize(, aTaskRow.Count).Value = aTaskRow.ToArray()
      aTaskRow.Clear
      aSummaries.Add lngRow
    Else
      For lngCol = lngNameCol + 1 To aHeaders.Count
        'this gets overwritten by a formula; account for resource type at assignment level
        If aHeaders(lngCol - 1)(1) = "Baseline Work" Then
          aTaskRow.Add Task.BaselineWork / 60
        ElseIf aHeaders(lngCol - 1)(1) = "Remaining Work" Then
          aTaskRow.Add Task.RemainingWork / 60
        'elseif = new evp then get physical %
        'elseif = revised etc then get remaining work and divide
        Else
          aTaskRow.Add Task.GetField(aHeaders(lngCol - 1)(0))
        End If
      Next lngCol

      'identify for formatting
      If Task.Milestone Then aMilestones.Add lngRow Else aNormal.Add lngRow

      'write task data to sheet
      xlCells(lngRow, 1).Resize(, aTaskRow.Count).Value = aTaskRow.ToArray()
      aTaskRow.Clear

      'get assignment data for incomplete tasks
      If Task.ActualFinish = "NA" Then
        'add remaining work formula '<issue58>
        If Task.Assignments.Count > 0 Then
          xlCells(lngRow, lngRemainingWorkCol).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C:R" & lngRow + Task.Assignments.Count & "C)"
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

          aTaskRow.Add Assignment.ResourceName
          xlCells(lngRow, lngNameCol).IndentLevel = Task.OutlineLevel + 2
          If Assignment.ResourceType = pjResourceTypeWork Then
            xlCells(lngRow, lngBaselineWorkCol).Value = Val(Assignment.BaselineWork) / 60     'Val() function prevents error when
            xlCells(lngRow, lngRemainingWorkCol).Value = Val(Assignment.RemainingWork) / 60   'values are null or "" which is read as text
          Else
            xlCells(lngRow, lngBaselineWorkCol).Value = Val(Assignment.BaselineWork)         'Val() function prevents error when
            xlCells(lngRow, lngRemainingWorkCol).Value = Val(Assignment.RemainingWork)       'values are null or "" which is read as text
          End If
          'revised etc = same as above
          xlCells(lngRow, 1).Resize(, aTaskRow.Count).Value = aTaskRow.ToArray()
          aTaskRow.Clear

        Next Assignment
      Else 'task is complete '<issue58>
        aCompleted.Add lngRow '<issue58>

      End If 'Task.ActualFinish = "NA"

    End If 'Task Summary

next_task:
    lngTask = lngTask + 1
    If lngTask > lngTaskCount Then lngTaskCount = lngTask
    Application.StatusBar = "Exporting..." & Format(lngTask, "#,##0") & " / " & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ")"
    cptStatusSheet_frm.lblStatus.Caption = " Exporting..." & Format(lngTask, "#,##0") & " / " & Format(lngTaskCount, "#,##0") & " (" & Format(lngTask / lngTaskCount, "0%") & ")"
    cptStatusSheet_frm.lblProgress.Width = (lngTask / (lngTaskCount)) * cptStatusSheet_frm.lblStatus.Width
    SelectCellDown
  Loop

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
  'lock task rows which will be a formula; unlock Assignment Rows
  'Worksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).Cells.Locked = False
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
  If aSummaries.Count > 0 Then '<issue16-17> added
    Set rSummaryTasks = xlCells(aSummaries(0), 1).Resize(, aHeaders.Count)
    For vCol = 1 To aSummaries.Count - 1
      Set rSummaryTasks = xlApp.Union(rSummaryTasks, xlCells(aSummaries(vCol), 1).Resize(, aHeaders.Count))
    Next vCol
    If Not rSummaryTasks Is Nothing Then
      rSummaryTasks.Interior.ThemeColor = xlThemeColorDark1
      rSummaryTasks.Interior.TintAndShade = -0.149998474074526
      rSummaryTasks.Font.Bold = True
    End If
  End If '</issue16-17>
  'format milestones
  If aMilestones.Count > 0 Then '<issue16-17> added
    Set rMilestones = xlCells(aMilestones(0), 1).Resize(, aHeaders.Count)
    For vCol = 1 To aMilestones.Count - 1
      Set rMilestones = xlApp.Union(rMilestones, xlCells(aMilestones(vCol), 1).Resize(, aHeaders.Count))
    Next vCol
    If Not rMilestones Is Nothing Then
      rMilestones.Font.ThemeColor = xlThemeColorAccent6
      rMilestones.Font.TintAndShade = -0.249977111117893
    End If
  End If '</issue16-17>
  'format normal
  If aNormal.Count > 0 Then '<issue16-17> added
    Set rNormal = xlCells(aNormal(0), 1).Resize(, aHeaders.Count)
    For vCol = 1 To aNormal.Count - 1
      Set rNormal = xlApp.Union(rNormal, xlCells(aNormal(vCol), 1).Resize(, aHeaders.Count))
    Next vCol
    If Not rNormal Is Nothing Then
      rNormal.Font.Color = RGB(32, 55, 100) '<issue54>
    End If
  End If '</issue16-17>
  'format assignments
  If aAssignments.Count > 0 Then '<issue16-17> added
    Set rAssignments = xlCells(aAssignments(0), 1).Resize(, aHeaders.Count)
    For vCol = 1 To aAssignments.Count - 1
      Set rAssignments = xlApp.Union(rAssignments, xlCells(aAssignments(vCol), 1).Resize(, aHeaders.Count))
    Next vCol
    If Not rAssignments Is Nothing Then rAssignments.Font.Italic = True
  End If '</issue16-17>
  'format completed '<issue58>
  If aCompleted.Count > 0 Then
    'format the entire row - assignments are skipped on completd tasks
    Set rCompleted = xlCells(aCompleted(0), 1).Resize(, aHeaders.Count)
    For vCol = 1 To aCompleted.Count - 1
      Set rCompleted = xlApp.Union(rCompleted, xlCells(aCompleted(vCol), 1).Resize(, aHeaders.Count))
    Next vCol
    If Not rCompleted Is Nothing Then
      rCompleted.Font.Italic = True
      rCompleted.Font.ColorIndex = 16
    End If
    'update ev% complete
    lngCol = Worksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
    Set rCompleted = xlCells(aCompleted(0), lngCol)
    For vCol = 1 To aCompleted.Count - 1 'we are borrowing vCol to iterate row numbers
      Set rCompleted = xlApp.Union(rCompleted, xlCells(aCompleted(vCol), lngCol))
    Next vCol
    rCompleted = 1
  End If '</issue58>
  If blnPerformanceTest Then Debug.Print "format rows: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'format common borders
  Set rng = Worksheet.Range(xlCells(lngHeaderRow, 1), xlCells(lngRow, aHeaders.Count))
  rng.BorderAround xlContinuous, xlThin
  rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
  rng.Borders(xlInsideHorizontal).Weight = xlThin
  If blnPerformanceTest Then Debug.Print "format common borders: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'rename headers
  Set rng = xlCells(lngHeaderRow, 1).Resize(, aHeaders.Count)
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
  For lngCol = 0 To aHeaders.Count - 1

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
    If aCentered.Contains(aHeaders(lngCol)(1)) Then
      If rCentered Is Nothing Then
        Set rCentered = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rCentered = xlApp.Union(rCentered, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'format entry headers and columns
    If aEntryHeaders.Contains(aHeaders(lngCol)(1)) Then 'if the column we're working on is included in the list of entry headers, then...
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

  Next 'lngCol = aHeader item
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
'  .Interior.TintAndShade = 0

  If blnPerformanceTest Then t = GetTickCount
  
  blnAddConditionalFormats = cptStatusSheet_frm.chkAddConditionalFormats = True
  If Not blnAddConditionalFormats Then GoTo conditional_formatting_skipped
  
  cptStatusSheet_frm.lblStatus.Caption = " Applying conditional formats..."
  Application.StatusBar = "Applying conditional formats..."
  cptStatusSheet_frm.lblProgress.Width = (1 / 100) * cptStatusSheet_frm.lblStatus.Width
  lngConditionalFormats = 18 'bash: "grep -c 'lngFormatCondition = lngFormatCondition + 1' Status/cptStatusSheet_bas.bas"
  'capture status date address
  strStatusDate = Worksheet.Range("STATUS_DATE").Address(True, True)

  'attempt to speed up
  xlApp.EnableEvents = False
  Worksheet.DisplayPageBreaks = False

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
  If Err.Number = 1004 Then 'no cells found
    Err.Clear '<issue52>
    GoTo new_finish '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  strFirstCell = rng(1).Address(False, True)

  '-->condition 1: blank and start is less than status date > update required
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = 7749439
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = -4105
    .Color = 10079487
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Start (1/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Start (1/5)"

  '-->condition 2: two-week-window                      '=IF($E50<=(INDIRECT("STATUS_DATE")+14),TRUE,FALSE)
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=(" & strStatusDate & "+14)),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16754788
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 10284031
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Start (2/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Start (2/5)"
  
  '-->condition 3: blank and EV% > 0 > invalid
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lngEVPCol).Address(False, True) & ">0),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Start (3/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Start (3/5)"

  '-->condition 4: greater than actual finish > invalid
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & ">" & xlCells(rng(1).Row, lngAFCol).Address(False, True) & "," & xlCells(rng(1).Row, lngAFCol).Address(False, True) & "<>""""),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Start (4/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Start (4/5)"

  'else: <> start > updated
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & "<>" & xlCells(rng(1).Row, lngASCol - 2).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Start (5/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Start (5/5)"

new_finish: '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  If Worksheet.AutoFilterMode Then Worksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lngAFCol, Criteria1:="="
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngAFCol), xlCells(lngRow, lngAFCol)).SpecialCells(xlCellTypeVisible)
  If Err.Number = 1004 Then '<issue52>
    Err.Clear '<issue52>
    GoTo ev_percent '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  strFirstCell = rng(1).Address(False, True)

  '-->condition 1: blank and finish is less than status date > update required
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, 4).Address(False, True) & "<1," & rng(1).Offset(0, -3).Address(False, True) & "<" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = 7749439
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = -4105
    .Color = 10079487
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Finish (1/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Finish (1/5)"

  '-->condition 2: two-week-window
  rng.FormatConditions.Add Type:=xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & rng(1).Offset(0, -2).Address(False, True) & "<=(" & strStatusDate & "+14)),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16754788
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 10284031
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Finish (2/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Finish (2/5)"

  '-->condition 3: less than actual start -> invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & xlCells(rng(1).Row, lngASCol).Address(False, True) & "<>""""," & strFirstCell & "<" & xlCells(rng(1).Row, lngASCol).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Finish (3/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Finish (3/5)"

  '-->condition 4: blank and EV% = 100 > invalid
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=""""," & xlCells(rng(1).Row, lngEVPCol).Address(False, True) & "=1),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Finish (4/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Finish (4/5)"

  'else: <> finish > updated
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "<>""""," & strFirstCell & "<>" & xlCells(rng(1).Row, lngAFCol - 2).Address(False, True) & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New Finish (5/5)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New Finish (5/5)"

ev_percent:
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Worksheet.AutoFilterMode Then Worksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52-noCellsFound>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngEVPCol), xlCells(lngRow, lngEVPCol)).SpecialCells(xlCellTypeVisible)
  If Err.Number = 1004 Then '<issue52>
    Err.Clear '<issue52>
    GoTo revised_etc '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  strFirstCell = rng(1).Address(False, True)

  '-->condition 0: =IF(AND($H48>$B$1,$L48>$K48,$L48<1),TRUE,FALSE) 'green
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & rng(1).Offset(0, -4).Address(False, True) & ">" & strStatusDate & "," & strFirstCell & ">" & rng(1).Offset(0, -1).Address(False, True) & "," & strFirstCell & "<1),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (1/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (1/7)"

  '-->condition 1: =IF(AND($H48>0,$H48<=$B$1,$L48=1),TRUE,FALSE) 'green
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & rng(1).Offset(0, -4).Address(False, True) & ">0," & rng(1).Offset(0, -4).Address(False, True) & "<=" & strStatusDate & "," & strFirstCell & "=1),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16752384
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13561798
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (2/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (2/7)"

  '-->condition 2: New Finish < Status Date (complete) and EV%<100 > invalid '=IF(AND($G48>0,$H48<=$B$1,$L48<1),TRUE,FALSE)
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & rng(1).Offset(0, -4).Address(False, True) & ">0," & rng(1).Offset(0, -3).Address(False, True) & ">1," & rng(1).Offset(0, -3).Address(False, True) & "<=" & strStatusDate & "," & strFirstCell & "<1),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (3/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (3/7)"

  '-->condition 3: Start < Status Date AND EV% < 100 > update required '  =IF(AND($L48<1,$E48<=$B$1),TRUE,FALSE) 'orange
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & rng(1).Address(False, True) & "<1," & rng(1).Offset(0, -7).Address(False, True) & "<=" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = 7749439
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = -4105
    .Color = 10079487
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (4/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (4/7)"

  '-->condition 4: EV% > 0 and new start = "" (bogus actuals) > invalid '  =IF(AND($L48>0,$G48=0),TRUE,FALSE) 'red
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & ">0," & xlCells(rng(1).Row, lngASCol).Address(False, True) & "=0),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (5/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (5/7)"

  '-->condition 5: EV% =100 and new finish = "" (update required) > invalid '  =IF(AND($L48=1,$H48=0),TRUE,FALSE) 'red
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=1," & xlCells(rng(1).Row, lngAFCol).Address(False, True) & "=0),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (6/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (6/7)"

  '-->condition 6: =100 and new finish > status date > invalid '  =IF(AND($L48=1,$H48>$B$1),TRUE,FALSE) 'red
  rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & strFirstCell & "=1," & xlCells(rng(1).Row, lngAFCol).Address(False, True) & ">" & strStatusDate & "),TRUE,FALSE)"
  With rng.FormatConditions(rng.FormatConditions.Count).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(rng.FormatConditions.Count).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New EV% (7/7)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New EV% (7/7)"

revised_etc:
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  If Worksheet.AutoFilterMode Then Worksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for Task
  rngAll.AutoFilter Field:=lngETCCol, Operator:=xlFilterAutomaticFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lngRemainingWorkCol, Operator:=xlFilterNoFill
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52>
  Set rng = Worksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).SpecialCells(xlCellTypeVisible)
  If Err.Number = 1004 Then '<issue52>
    Err.Clear '<issue52>
    GoTo evt_vs_evp '<issue52>
  End If '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  rng.Cells.Locked = False
  strFirstCell = rng(1).Address(False, True)

  lngLastRow = lngRow

  For lngRow = lngHeaderRow + 1 To lngLastRow
    If xlCells(lngRow, lngETCCol).Font.Color = RGB(32, 55, 100) Then
      lngTaskRow = lngRow
    ElseIf xlCells(lngRow, lngETCCol).Font.Italic And xlCells(lngRow, lngETCCol).Font.Color = 0 Then '0 = xlAutomaticFontColor
      Set rng = xlCells(lngRow, lngETCCol)

      '-->condition 1: =IF(AND($H$48>0,$H$48<=$B$1,$O$49=0),TRUE,FALSE) 'green
      rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & xlCells(lngTaskRow, lngAFCol).Address(True, True) & ">0," & xlCells(lngTaskRow, lngAFCol).Address(True, True) & "<=" & strStatusDate & "," & xlCells(lngRow, lngETCCol).Address(True, True) & "=0),TRUE,FALSE)"
      With rng.FormatConditions(rng.FormatConditions.Count).Font
        .Color = -16752384
        .TintAndShade = 0
      End With
      With rng.FormatConditions(rng.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
      End With
      rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True

      'todo: if in-progress and new etc <> prev etc

      '-->condition 2: new actual start and ETC > 0 =IF(AND($H48>0,$H48<=$B$1,$O$49>0),TRUE,FALSE) 'red
      rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & xlCells(lngTaskRow, lngAFCol).Address(True, True) & ">0," & xlCells(lngTaskRow, lngAFCol).Address(True, True) & "<=" & strStatusDate & "," & xlCells(lngRow, lngETCCol).Address(True, True) & ">0),TRUE,FALSE)"
      With rng.FormatConditions(rng.FormatConditions.Count).Font
        .Color = -16383844
        .TintAndShade = 0
      End With
      With rng.FormatConditions(rng.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
      End With
      rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True

      '-->condition 3: ev=100 and etc> 0 '=IF(AND($L48=1,$O49>0),TRUE,FALSE) 'red
      rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & xlCells(lngTaskRow, lngEVPCol).Address(True, True) & "=1," & rng.Address(True, True) & ">0),TRUE,FALSE)"
      With rng.FormatConditions(rng.FormatConditions.Count).Font
        .Color = -16383844
        .TintAndShade = 0
      End With
      With rng.FormatConditions(rng.FormatConditions.Count).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
      End With
      rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True

      '-->condition 4: Start < Status Date AND EV% < 100 > update required =IF(AND($L$48<1,$E$48<=(INDIRECT("STATUS_DATE"))),TRUE,FALSE) 'orange
      rng.FormatConditions.Add xlExpression, Formula1:="=IF(AND(" & xlCells(lngTaskRow, lngEVPCol).Address(True, True) & "<1," & xlCells(lngTaskRow, lngASCol - 2).Address(True, True) & "<=" & strStatusDate & "),TRUE,FALSE)"
      With rng.FormatConditions(rng.FormatConditions.Count).Font
        .Color = 7749439
        .TintAndShade = 0
      End With
      With rng.FormatConditions(rng.FormatConditions.Count).Interior
        .PatternColorIndex = -4105
        .Color = 10079487
        .TintAndShade = 0
      End With
      rng.FormatConditions(rng.FormatConditions.Count).StopIfTrue = True

    End If
    cptStatusSheet_frm.lblStatus.Caption = "Adding ETC conditional formats (" & Format(lngRow / lngLastRow, "0%") & ")"
    Application.StatusBar = "Adding ETC conditional formats (" & Format(lngRow / lngLastRow, "0%") & ")"
  Next lngRow
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New ETC (4/4)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New ETC (4/4)"

evt_vs_evp:
  If cptStatusSheet_frm.cboCostTool <> "<none>" Then '<issue64>
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
  End If '</issue64>
  If blnPerformanceTest Then Debug.Print "apply conditional formatting " & (GetTickCount - t) / 1000

  Debug.Print lngFormatCondition & " format conditions applied."

conditional_formatting_skipped:

  'unlock Revised ETC at assignment level
  lngCol = Worksheet.Rows(lngHeaderRow).Find("Revised ETC", lookat:=xlWhole).Column
  For lngItem = 0 To aAssignments.Count - 1
    xlCells(aAssignments(lngItem), lngCol).Locked = False
  Next lngItem

  'optionallly set reference to Outlook and prepare to email
  blnEmail = cptStatusSheet_frm.chkSendEmails = True
  If blnEmail Then
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    cptStatusSheet_frm.lblStatus.Caption = "Getting Outlook..."
    Application.StatusBar = "Getting Outlook..."
  End If

  If blnPerformanceTest Then t = GetTickCount
  cptStatusSheet_frm.lblStatus.Caption = "Saving Workbook" & IIf(cptStatusSheet_frm.cboCreate.Value = "1", "s", "") & "..."
  Application.StatusBar = "Saving Workbook" & IIf(cptStatusSheet_frm.cboCreate.Value = "1", "s", "") & "..."
  strDir = Environ("USERPROFILE") & "\CP_Status_Sheets\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'get clean project name
  strFileName = cptRemoveIllegalCharacters(ActiveProject.Name)
  strFileName = Replace(strFileName, ".mpp", "")
  'create project status folder
  strDir = strDir & Replace(strFileName, " ", "") & "\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'create project status date folder
  strDir = strDir & Format(dtStatus, "yyyy-mm-dd") & "\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  strFileName = "SS_" & strFileName & "_" & Format(dtStatus, "yyyy-mm-dd") & ".xlsx"
  strFileName = Replace(strFileName, " ", "")
  
  'create single workbook
  If cptStatusSheet_frm.cboCreate.Value = "0" Then
  
    xlApp.Visible = True '<issue81> - move this below if option = (0|other)
    xlApp.WindowState = xlMaximized
    xlApp.ScreenUpdating = True
    
    If Worksheet.AutoFilterMode Then Worksheet.ShowAllData
    
    xlApp.ActiveWindow.ScrollColumn = 1
    xlApp.ActiveWindow.ScrollRow = 1 '<issue54>
    xlCells(lngHeaderRow + 1, lngNameCol + 1).Select
    xlApp.ActiveWindow.FreezePanes = True
    'prettify the task name column
    Worksheet.Columns(lngNameCol).AutoFit
  
    'protect the sheet
    Worksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
    Worksheet.EnableSelection = xlNoRestrictions
    'save to desktop in folder for status date
    On Error Resume Next
    If Dir(strDir & strFileName) <> vbNullString Then Kill strDir & strFileName
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    'account for if the file exists and is open in the background
    If Dir(strDir & strFileName) <> vbNullString Then  'delete failed, rename with timestamp
      strMsg = "'" & strFileName & "' already exists, and is likely open." & vbCrLf
      strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "hh-nn-ss") & ".xlsx")
      strMsg = strMsg & "The file you are now creating will be named '" & strFileName & "'"
      MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
      Workbook.SaveAs strDir & strFileName, 51
    Else
      Workbook.SaveAs strDir & strFileName, 51
    End If
    If blnEmail Then
      Set MailItem = olApp.CreateItem(0) '0 = olMailItem
      MailItem.Attachments.Add strDir & strFileName
      MailItem.Subject = "Status Request - " & Format(dtStatus, "yyyy-mm-dd")
      MailItem.Display False
    End If
  Else
  
    xlCells(lngHeaderRow + 1, lngNameCol + 1).Select
    xlApp.ActiveWindow.FreezePanes = True
    'prettify the task name column
    Worksheet.Columns(lngNameCol).AutoFit
  
    'cycle through each option and create sheet
    For lngItem = 0 To aEach.Count - 1 'To 0 Step -1 don't reverse it
      cptStatusSheet_frm.lblStatus.Caption = "Creating " & aEach.getKey(lngItem) & "..."
      Application.StatusBar = "Creating " & aEach.getKey(lngItem) & "..."
      lngLastRow = Workbook.Worksheets(1).[A8].End(xlDown).Row
      lngRow = Workbook.Worksheets(1).[A8].End(xlDown).Row
      Workbook.Sheets(1).Copy After:=Workbook.Sheets(Workbook.Sheets.Count) 'Set = Copy( doesn't work
      Set Worksheet = Workbook.Sheets(Workbook.Sheets.Count)
      Worksheet.Name = aEach.getKey(lngItem)
      SetAutoFilter FieldName:=cptStatusSheet_frm.cboEach, FilterType:=pjAutoFilterIn, Criteria1:=aEach.getKey(lngItem)
      'get array of task and assignment unique ids
      Worksheet.Cells(lngRow + 2, 1).Value = "KEEP"
      'get group by summaries
      SelectBeginning
      Do
        On Error Resume Next
        Set Task = ActiveCell.Task
        If Err.Number > 0 Then
          Err.Number = 0
          Err.Clear
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          Exit Do
        End If
        If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
        If Task.GroupBySummary Then
          Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Offset(1, 0) = aGroups.getValueList()(aGroups.indexOfKey(Task.Name))
        End If
        SelectCellDown
      Loop
      SelectAll
      For Each Task In ActiveSelection.Tasks
        Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Offset(1, 0) = Task.UniqueID
        For Each Assignment In Task.Assignments
          Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Offset(1, 0) = Assignment.UniqueID
        Next Assignment
      Next Task
      'name the range of uids to keep
      Set rngKeep = Worksheet.Cells(lngRow + 2, 1)
      Set rngKeep = Worksheet.Range(rngKeep, rngKeep.End(xlDown))
      Workbook.Names.Add Name:="KEEP", RefersToR1C1:="='" & aEach.getKey(lngItem) & "'!" & rngKeep.Address(True, True, xlR1C1)
      'add a formula to find which rows to keep
      lngLastCol = Worksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(1, 1).Column
      Set rngKeep = Worksheet.Range(Worksheet.Cells(lngHeaderRow + 1, lngLastCol), Worksheet.Cells(lngRow, lngLastCol))
      rngKeep(1).Offset(-1, 0).Value = "KEEP"
      rngKeep.Formula = "=IFERROR(VLOOKUP(A" & lngHeaderRow + 1 & ",KEEP,1,FALSE),""DELETE"")"
      xlApp.Calculate
      'Worksheet.Cells(lngHeaderRow, 1).AutoFilter
      lngLastRow = lngRow
      For lngRow = lngLastRow To lngHeaderRow + 1 Step -1
        If Worksheet.Cells(lngRow, rngKeep.Column) = "DELETE" Then Worksheet.Rows(lngRow).Delete Shift:=xlUp
        cptStatusSheet_frm.lblStatus.Caption = "Creating " & aEach.getKey(lngItem) & "...(" & Format(((lngLastRow - lngRow) / (lngLastRow - lngHeaderRow)), "0%") & ")"
        cptStatusSheet_frm.lblProgress.Width = ((lngLastRow - lngRow) / (lngLastRow - lngHeaderRow)) * cptStatusSheet_frm.lblStatus.Width
        Application.StatusBar = "Creating " & aEach.getKey(lngItem) & "...(" & Format(((lngLastRow - lngRow) / (lngLastRow - lngHeaderRow)), "0%") & ")"
      Next lngRow
      cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
      Worksheet.Cells(lngHeaderRow, 1).Select
      xlApp.Selection.AutoFilter
      Worksheet.Columns(lngLastCol).Delete
      Worksheet.Range("KEEP").Clear
      Workbook.Names("KEEP").Delete
      Worksheet.[B1].Select
      Worksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
      Worksheet.EnableSelection = xlNoRestrictions
      DoEvents
    Next lngItem

    xlApp.ScreenUpdating = True
    xlApp.Calculation = True

    'handle for each
    If cptStatusSheet_frm.cboCreate = "1" Then 'worksheet for each
      'Workbook.SaveAs strDir & strFileName, 51 '</issue80>
      'save in folder for status date '<issue80>
      On Error Resume Next
      If Dir(strDir & strFileName) <> vbNullString Then Kill strDir & strFileName
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      'account for if the file exists and is open in the background
      If Dir(strDir & strFileName) <> vbNullString Then  'delete failed, rename with timestamp
        strMsg = "'" & strFileName & "' already exists, and is likely open." & vbCrLf
        strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "hh-nn-ss") & ".xlsx")
        strMsg = strMsg & "The file you are now creating will be named '" & strFileName & "'"
        MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
        Workbook.SaveAs strDir & strFileName, 51
      Else
        Workbook.SaveAs strDir & strFileName, 51
      End If
      If blnEmail Then
        Set MailItem = olApp.CreateItem(0) '0 = olMailItem
        MailItem.Attachments.Add strDir & strFileName
        MailItem.Subject = "Status Request - " & Format(dtStatus, "yyyy-mm-dd")
        MailItem.Display False
      End If
    ElseIf cptStatusSheet_frm.cboCreate.Value = "2" Then 'workbook for each
      For lngItem = aEach.Count - 1 To 0 Step -1
        cptStatusSheet_frm.lblStatus.Caption = "Saving " & aEach.getKey(lngItem) & "..."
        Application.StatusBar = "Saving " & aEach.getKey(lngItem) & "..."
        Workbook.Sheets(aEach.getKey(lngItem)).Copy
        xlApp.ActiveWorkbook.Worksheets(1).Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, UserInterfaceOnly:=True
        xlApp.ActiveWorkbook.Worksheets(1).EnableSelection = xlNoRestrictions
        On Error Resume Next
        If Dir(strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx")) <> vbNullString Then Kill strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx")
        'xlApp.ActiveWorkbook.SaveAs strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx"), 51 '</issue80>
        If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
        'account for if the file exists and is open in the background
        If Dir(strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx")) <> vbNullString Then  'delete failed, rename with timestamp
          strMsg = "'" & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx") & "' already exists, and is likely open." & vbCrLf
          strFileName = Replace(Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx"), ".xlsx", "_" & Format(Now, "hh-nn-ss") & ".xlsx")
          strMsg = strMsg & "The file you are now creating will be named '" & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx") & "'"
          MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
          xlApp.ActiveWorkbook.SaveAs strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx"), 51
        Else
          xlApp.ActiveWorkbook.SaveAs strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx"), 51
        End If
        If blnEmail Then
          Set MailItem = olApp.CreateItem(0) '0 = olMailItem
          MailItem.Attachments.Add strDir & Replace(strFileName, ".xlsx", "_" & aEach.getKey(lngItem) & ".xlsx")
          MailItem.Subject = "Status Request [" & aEach.getKey(lngItem) & "] " & Format(dtStatus, "yyyy-mm-dd")
          MailItem.Display False
        End If
      Next lngItem
      Workbook.Close False
    End If

    cptStatusSheet_frm.lblStatus.Caption = "Wrapping up..."
    Application.StatusBar = "Wrapping up..."
    
    'reset autofilter
    strFieldName = cptStatusSheet_frm.cboEach.Value
    For lngItem = 0 To aEach.Count - 1
      strCriteria = strCriteria & aEach.getKey(lngItem) & Chr$(9)
    Next
    strCriteria = Left(strCriteria, Len(strCriteria) - 1)
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
    SelectBeginning

  End If
  If blnPerformanceTest Then Debug.Print "save workbook: " & (GetTickCount - t) / 1000
  If blnPerformanceTest Then Debug.Print "</=====PERFORMANCE TEST=====>"

  cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
  cptStatusSheet_frm.lblStatus.Caption = " Complete."
  Application.StatusBar = "Complete."
  MsgBox "Status Sheet(s) Created", vbInformation + vbOKOnly, "ClearPlan Status Sheet"
  xlApp.Visible = True
  cptStatusSheet_frm.lblStatus.Caption = " Ready..."
  Application.StatusBar = ""

exit_here:
  On Error Resume Next
  Application.DefaultDateFormat = lngDateFormat
  ActiveProject.SpaceBeforeTimeLabels = blnSpace
  ActiveProject.DayLabelDisplay = lngDayLabelDisplay
  If xlApp.Workbooks.Count > 0 Then xlApp.Calculation = xlAutomatic
  xlApp.ScreenUpdating = True
  xlApp.EnableEvents = True
  Set rCompleted = Nothing
  Set aCompleted = Nothing
  Set aGroups = Nothing
  Set rngKeep = Nothing
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
  Set olApp = Nothing
  Set MailItem = Nothing
  Set objDoc = Nothing
  Set objWord = Nothing
  Set objSel = Nothing
  Set objETemp = Nothing
  Exit Sub

err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCreateStatusSheet", Err, Erl)
  If Not xlApp Is Nothing Then
    If Not Workbook Is Nothing Then Workbook.Close False
    xlApp.Quit
  End If
  Resume exit_here

End Sub

Sub cptRefreshStatusTable()
'objects
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not cptStatusSheet_frm.Visible Then GoTo exit_here

  cptSpeed True

  'reset the table
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  lngItem = 0
  If cptStatusSheet_frm.lboExport.ListCount > 0 Then
    For lngItem = 0 To cptStatusSheet_frm.lboExport.ListCount - 1
      TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(cptStatusSheet_frm.lboExport.List(lngItem, 0)), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    Next lngItem
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Name", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Duration", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Actual Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Actual Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If cptStatusSheet_frm.cboEVT <> 0 Then
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=cptStatusSheet_frm.cboEVT.Value, Title:="", Width:=5, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  If cptStatusSheet_frm.cboEVP <> 0 Then
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=cptStatusSheet_frm.cboEVP.Value, Title:="EV%", Width:=5, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Work", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False, ShowAddNewColumn:=False
  TableApply Name:="cptStatusSheet Table"

  'reset the filter
  FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Actual Finish", test:="equals", Value:="NA", ShowInMenu:=False, ShowSummaryTasks:=True
  If cptStatusSheet_frm.chkHide And IsDate(cptStatusSheet_frm.txtHideCompleteBefore) Then
    FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Actual Finish", test:="is greater than or equal to", Value:=cptStatusSheet_frm.txtHideCompleteBefore, Operation:="Or", ShowSummaryTasks:=True
  End If
  FilterApply "cptStatusSheet Filter"

exit_here:
  On Error Resume Next
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptRefreshStatusView", Err, Erl)
  Err.Clear
  Resume exit_here
End Sub
