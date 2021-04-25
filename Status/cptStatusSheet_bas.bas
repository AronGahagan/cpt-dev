Attribute VB_Name = "cptStatusSheet_bas"
'<cpt_version>v1.3.0</cpt_version>
Option Explicit
#If Win64 And VBA7 Then '<issue53>
  Declare PtrSafe Function GetTickCount Lib "Kernel32" () As LongPtr '<issue53>
#Else '<issue53>
  Declare Function GetTickCount Lib "kernel32" () As Long
#End If '<issue53>
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_heref Else On Error GoTo 0
Private Const adVarChar As Long = 200
Private strStartingViewTopPane As String
Private strStartingViewBottomPane As String
Private strStartingTable As String
Private strStartingFilter As String
Private strStartingGroup As String
Private oAssignmentRange As Excel.Range
Private oNumberValidationRange As Excel.Range
Private oInputRange As Excel.Range
Private oUnlockedRange As Excel.Range
Private oEntryHeaderRange As Excel.Range

Sub cptShowStatusSheet_frm()
'populate all outline codes, text, and number fields
'populate UID,[user selections],Task Name,Duration,Forecast Start,Forecast Finish,Total Slack,[EVT],EV%,New EV%,BLW,Remaining Work,Revised ETC,BLS,BLF,Reason/Impact/Action
'add pick list for EV% or default to Physical % Complete
'objects
Dim rstFields As ADODB.Recordset 'Object
Dim rstEVT As ADODB.Recordset 'Object
Dim rstEVP As ADODB.Recordset 'Object
'longs
Dim lngField As Long, lngItem As Long
'integers
Dim intField As Integer
'strings
Dim strLocked As String
Dim strDataValidation As String
Dim strConditionalFormats As String
Dim strEmail As String
Dim strEach As String
Dim strCostTool As String
Dim strHide As String
Dim strCreate As String
Dim strEVP As String
Dim strEVT As String
Dim strFieldNamesChanged As String
Dim strFieldName As String, strFileName As String
'dates
Dim dtStatus As Date
'variants
Dim vFieldType As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'requires ms excel
  Application.StatusBar = "Validating OLE references..."
  DoEvents
  If Not cptCheckReference("Excel") Then
    #If Win64 Then
      MsgBox "A reference to Microsoft Excel (64-bit) could not be set.", vbExclamation + vbOKOnly, "Excel Required"
      GoTo exit_here
    #Else
      MsgBox "A reference to Microsoft Excel (32-bit) could not be set.", vbExclamation + vbOKOnly, "Excel Required"
    #End If
  End If
  'requires scripting (cptRegEx)
  If Not cptCheckReference("Scripting") Then GoTo exit_here
  
  'reset options
  Application.StatusBar = "Loading default settings..."
  DoEvents
  With cptStatusSheet_frm
    .lboFields.Clear
    .lboExport.Clear
    .cboEVT.Clear
    .cboEVP.Clear
    '.cboEVP.AddItem "Physical % Complete"
    .cboCostTool.Clear
    .cboCostTool.AddItem "COBRA"
    .cboCostTool.AddItem "MPM"
    .cboCostTool.AddItem "<none>"
    For lngItem = 0 To 2
      .cboCreate.AddItem
      .cboCreate.List(lngItem, 0) = lngItem
      .cboCreate.List(lngItem, 1) = Choose(lngItem + 1, "A Single Workbook", "A Worksheet for each", "A Workbook for each")
    Next lngItem
    .chkSendEmails.Enabled = cptCheckReference("Outlook")
    .chkHide = True
    .chkAddConditionalFormats = False
    .chkValidation = True
    .chkLocked = True
  End With

  'set up arrays to capture values
  Application.StatusBar = "Getting local custom fields..."
  DoEvents
  Set rstFields = CreateObject("ADODB.Recordset")
  rstFields.Fields.Append "CONSTANT", adBigInt
  rstFields.Fields.Append "NAME", adVarChar, 200
  rstFields.Fields.Append "TYPE", adVarChar, 50
  rstFields.Open
  
  For Each vFieldType In Array("Text", "Outline Code", "Number")
    On Error GoTo err_here
    For intField = 1 To 30
      lngField = FieldNameToFieldConstant(vFieldType & intField, pjTask)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        If vFieldType = "Number" Then
          rstFields.AddNew Array(0, 1, 2), Array(lngField, strFieldName, "Number")
        Else
          rstFields.AddNew Array(0, 1, 2), Array(lngField, strFieldName, "Text")
        End If
      End If
next_field:
    Next intField
  Next vFieldType
  
  'add Physical % Complete
  rstFields.AddNew Array(0, 1, 2), Array(FieldNameToFieldConstant("Physical % Complete"), "Physical % Complete", "Number")
  
  'add Contact field
  rstFields.AddNew Array(0, 1, 2), Array(FieldNameToFieldConstant("Contact"), "Contact", "Text")
  
  'get enterprise custom fields
  Application.StatusBar = "Getting Enterprise custom fields..."
  DoEvents
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      rstFields.AddNew Array(0, 1, 2), Array(lngField, Application.FieldConstantToFieldName(lngField), "Enterprise")
    End If
  Next lngField

  'add custom fields
  Application.StatusBar = "Populating Export Field list box..."
  DoEvents
  rstFields.Sort = "NAME"
  If rstFields.RecordCount > 0 Then
    rstFields.MoveFirst
    With cptStatusSheet_frm
      Do While Not rstFields.EOF
        If rstFields(1) = "Physical % Complete" Then GoTo skip_fields
        .lboFields.AddItem
        .lboFields.List(.lboFields.ListCount - 1, 0) = rstFields(0)
        .lboFields.List(.lboFields.ListCount - 1, 1) = rstFields(1)
        If FieldNameToFieldConstant(rstFields(1)) >= 188776000 Then
          .lboFields.List(.lboFields.ListCount - 1, 2) = "Enterprise"
        Else
          .lboFields.List(.lboFields.ListCount - 1, 2) = FieldConstantToFieldName(rstFields(0))
        End If
skip_fields:
        'add to Each
        If rstFields(1) <> "Physical % Complete" Then .cboEach.AddItem rstFields(1)
        If rstFields(2) = "Text" Then 'add to EVT
          .cboEVT.AddItem rstFields(1)
        ElseIf rstFields(2) = "Number" Or rstFields(1) = "Physical % Complete" Then 'add to EVP
          .cboEVP.AddItem rstFields(1)
        Else 'todo: add to both?
          .cboEVT.AddItem rstFields(1)
          .cboEVP.AddItem rstFields(1)
        End If
        rstFields.MoveNext
      Loop
    End With
  Else
    MsgBox "No Custom Fields have been set up in this file.", vbInformation + vbOKOnly, "No Fields Found"
    GoTo exit_here
  End If
  
  'convert saved settings if they exist
  strFileName = cptDir & "\settings\cpt-status-sheet.adtg"
  If Dir(strFileName) <> vbNullString Then
    Application.StatusBar = "Converting saved settings..."
    DoEvents
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      If Not .EOF Then
        .MoveFirst
        cptSaveSetting "StatusSheet", "cboEVT", .Fields(0)
        cptSaveSetting "StatusSheet", "cboEVP", .Fields(1)
        cptSaveSetting "StatusSheet", "cboCreate", .Fields(2) - 1
        cptSaveSetting "StatusSheet", "chkHide", .Fields(3)
        If .Fields.Count >= 5 Then
          cptSaveSetting "StatusSheet", "cboCostTool", .Fields(4)
        End If
        If .Fields.Count >= 6 Then
          cptSaveSetting "StatusSheet", "cboEach", .Fields(5)
        End If
      End If
      .Close
      Kill strFileName
    End With
  End If
  
  'import saved settings
  With cptStatusSheet_frm
    Application.StatusBar = "Getting saved settings..."
    DoEvents
    strEVT = cptGetSetting("StatusSheet", "cboEVT")
    If strEVT <> "" Then
      If rstFields.RecordCount > 0 Then
        rstFields.MoveFirst
        rstFields.Find "NAME='" & strEVT & "'"
        If Not rstFields.EOF Then
          On Error Resume Next
          .cboEVT.Value = strEVT
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          If Err.Number > 0 Then
            MsgBox "Unable to set EVT Field to '" & rstFields(1) & "' - contact cpt@ClearPlanConsulting.com if you need assistance.", vbExclamation + vbOKOnly, "Cannot assign EVT"
            Err.Clear
          End If
        End If
      End If
    End If
    strEVP = cptGetSetting("StatusSheet", "cboEVP")
    If strEVP <> "" Then
      If rstFields.RecordCount > 0 Then
        rstFields.MoveFirst
        rstFields.Find "NAME='" & strEVP & "'"
        If Not rstFields.EOF Then
          On Error Resume Next
          .cboEVP.Value = strEVP
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          If Err.Number > 0 Then
            MsgBox "Unable to set EV% Field to '" & rstFields(1) & "' - contact cpt@ClearPlanConsulting.com if you need assistance.", vbExclamation + vbOKOnly, "Cannot assign EVP"
            Err.Clear
          End If
        End If
      End If
    End If
    strCreate = cptGetSetting("StatusSheet", "cboCreate")
    If strCreate <> "" Then .cboCreate.Value = CLng(strCreate)
    strHide = cptGetSetting("StatusSheet", "chkHide")
    If strHide <> "" Then .chkHide = CBool(strHide)
    strCostTool = cptGetSetting("StatusSheet", "cboCostTool")
    If strCostTool <> "" Then .cboCostTool.Value = strCostTool
    If .cboCreate <> 0 Then
      strEach = cptGetSetting("StatusSheet", "cboEach")
      If strEach <> "" Then
        If rstFields.RecordCount > 0 Then
          rstFields.MoveFirst
          rstFields.Find "NAME='" & strEach & "'"
          If Not rstFields.EOF Then
            On Error Resume Next
            .cboEach.Value = strEach '<none> would not be found
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            If Err.Number > 0 Then
              MsgBox "Unable to set 'For Each' Field to '" & rstFields(1) & "' - contact cpt@ClearPlanConsulting.com if you need assistance.", vbExclamation + vbOKOnly, "Cannot assign For Each"
              Err.Clear
            End If
          End If
        End If
      End If
    End If
    strEmail = cptGetSetting("StatusSheet", "chkEmail")
    If strEmail <> "" Then .chkSendEmails = CBool(strEmail)
    strConditionalFormats = cptGetSetting("StatusSheet", "chkConditionalFormatting")
    If strConditionalFormats <> "" Then .chkAddConditionalFormats = CBool(strConditionalFormats)
    strDataValidation = cptGetSetting("StatusSheet", "chkDataValidation")
    If strDataValidation <> "" Then .chkValidation = CBool(strDataValidation)
    strLocked = cptGetSetting("StatusSheet", "chkLocked")
    If strLocked <> "" Then .chkLocked = CBool(strLocked)
  End With

  'add saved export fields if they exist
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      If .RecordCount > 0 Then
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
      End If
      .Close
    End With
  End If

  'reset the view - must be of type pjTaskItem
  'todo: restore starting table, filter, and group (but keep existing view)
  ActiveWindow.TopPane.Activate
  strStartingViewTopPane = ActiveWindow.TopPane.View.Name
  If Not ActiveWindow.BottomPane Is Nothing Then
    strStartingViewBottomPane = ActiveWindow.BottomPane.View.Name
  Else
    strStartingViewBottomPane = "None"
  End If
  strStartingTable = ActiveProject.CurrentTable
  strStartingFilter = ActiveProject.CurrentFilter
  strStartingGroup = ActiveProject.CurrentGroup
  
  'cptSpeed True
  If strStartingGroup <> "No Group" Then GroupApply "No Group"
  If ActiveWindow.TopPane.View.Name <> "Gantt Chart" Then
    If MsgBox("Current view must be changed for successful export.", vbInformation + vbOKCancel, "Incompatible View") = vbOK Then
      ActiveWindow.TopPane.Activate
      ViewApply "Gantt Chart"
    Else
      GoTo exit_here
    End If
  End If

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
  Application.StatusBar = "Preparing View/Table/Filter..."
  DoEvents
  FilterClear
  OptionsViewEx displaysummarytasks:=True, displaynameindent:=True
  If strStartingGroup = "No Group" Then
    Application.StatusBar = "Ensuring Sort by ID keeping Outline Structure..."
    DoEvents
    Sort "ID", , , , , , False, True 'OutlineShowAllTasks won't work without this
  Else
    GroupApply strStartingGroup
  End If
  Application.StatusBar = "Showing all tasks..."
  DoEvents
  OutlineShowAllTasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  cptSpeed False
  cptStatusSheet_frm.Show False
  cptRefreshStatusTable 'this only runs when form is visible
  Application.StatusBar = "Ready..."
  DoEvents

  If Len(strFieldNamesChanged) > 0 Then
    strFieldNamesChanged = "The following saved export field names have changed:" & vbCrLf & vbCrLf & strFieldNamesChanged
    strFieldNamesChanged = strFieldNamesChanged & vbCrLf & vbCrLf & "You may wish to remove them from the export list."
    MsgBox strFieldNamesChanged, vbInformation + vbOKOnly, "Saved Settings - Mismatches"
  End If

exit_here:
  On Error Resume Next
  If rstFields.State Then rstFields.Close
  Set rstFields = Nothing
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
  Dim oTasks As Tasks, oTask As Task, oAssignment As Assignment
  'early binding:
  Dim oExcel As Excel.Application, oWorkbook As Workbook, oWorksheet As Worksheet, rng As Excel.Range
  Dim rSummaryTasks As Excel.Range, rMilestones As Excel.Range, rNormal As Excel.Range, rAssignments As Excel.Range, rLockedCells As Excel.Range
  Dim rDates As Excel.Range, rWork As Excel.Range, rMedium As Excel.Range, rCentered As Excel.Range, rEntry As Excel.Range
  Dim xlCells As Excel.Range, rngAll As Excel.Range
  Dim oOutlook As Outlook.Application, oMailItem As MailItem, oDoc As Word.Document, oWord As Word.Application, oSel As Word.Selection, oETemp As Word.Template
  Dim aSummaries As Object, aMilestones As Object, aNormal As Object, aAssignments As Object
  Dim rstEach As ADODB.Recordset, aTaskRow As Object, rstColumns As ADODB.Recordset
  'late binding:
'  Dim oExcel As Object, oWorkbook As Object, oWorksheet As Object, rng As Object
'  Dim rSummaryTasks As Object, rMilestones As Object, rNormal As Object, rAssignments As Object, rLockedCells As Object
'  Dim rDates As Object, rWork As Object, rMedium As Object, rCentered As Object, rEntry As Object
'  Dim xlCells As Object, rngAll As Object
'  Dim oOutlook As Object, oMailItem As Object, objDoc As Object, oWord As Object, oSel As Object, oETemp As Object
'  Dim aSummaries As Object, aMilestones As Object, aNormal As Object, aAssignments As Object
'  Dim rstEach As Object, aTaskRow As Object, rstColumns As Object
'  Dim oOddBalls As Object, aCentered As Object, aEntryHeaders As Object
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
  Dim vHeader As Variant
  Dim vCol As Variant, vUserFields As Variant
  'booleans
  Dim blnLocked As Boolean
  Dim blnValidation As Boolean
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
  Set oTasks = ActiveProject.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ensure project has tasks
  If oTasks Is Nothing Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "Create Status Sheet"
    GoTo exit_here
  End If
    
  cptStatusSheet_frm.lblStatus.Caption = " Analyzing project..."
  Application.StatusBar = "Analyzing project..."
  blnValidation = cptStatusSheet_frm.chkValidation = True
  blnLocked = cptStatusSheet_frm.chkLocked = True
  
  'get task count
  If blnPerformanceTest Then t = GetTickCount
  SelectAll
  Set oTasks = ActiveSelection.Tasks
  lngTaskCount = oTasks.Count
  If blnPerformanceTest Then Debug.Print "<=====PERFORMANCE TEST " & Now() & "=====>"

  cptStatusSheet_frm.lblStatus.Caption = " Setting up Workbook..."
  Application.StatusBar = "Setting up Workbook..."
  DoEvents
  'set up an excel Workbook
  If blnPerformanceTest Then t = GetTickCount
  Set oExcel = CreateObject("Excel.Application")
  'oExcel.Visible = False
  oExcel.WindowState = xlMinimized
  '/=== debug ==\
  If Not BLN_TRAP_ERRORS Then oExcel.Visible = True
  '\=== debug ===/
  
  If blnPerformanceTest Then Debug.Print "set up excel Workbook: " & (GetTickCount - t) / 1000

  'get status date
  If ActiveProject.StatusDate = "NA" Then
    dtStatus = Now()
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'cptCopyData task data and applies formatting, validation, protection
  'cptCopyData also extracts existing assignment data and applies formatting, validation, protection
  'obective is to only loop through tasks once
  
  'copy/paste the data
  lngHeaderRow = 8
  With cptStatusSheet_frm
    If .cboCreate.Value = "0" Then 'single workbook
      Set oWorkbook = oExcel.Workbooks.Add
      oExcel.Calculation = xlCalculationManual
      oExcel.ScreenUpdating = False
      Set oWorksheet = oWorkbook.Sheets(1)
      oWorksheet.Name = "Status Sheet"
      'copy data
      If blnPerformanceTest Then t = GetTickCount
      .lblStatus = "Creating Workbook..."
      cptCopyData oWorksheet, lngHeaderRow
      If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
      'add legend
      If blnPerformanceTest Then t = GetTickCount
      cptAddLegend oWorksheet, dtStatus
      If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
      
      'final formatting
      cptFinalFormats oWorksheet
      
      oWorksheet.Calculate
      
      If blnLocked Then 'protect the sheet
        oWorksheet.Protect Password:="NoTouching!", AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
        oWorksheet.EnableSelection = xlNoRestrictions
      End If
      
      .lblStatus = "Creating Workbook...done"
      
      'todo: save workbook
      
    ElseIf .cboCreate.Value = "1" Then  'worksheet for each
      Set oWorkbook = oExcel.Workbooks.Add
      oExcel.Calculation = xlCalculationManual
      oExcel.ScreenUpdating = False
      For lngItem = 0 To .lboItems.ListCount - 1
        If .lboItems.Selected(lngItem) Then
          strItem = .lboItems.List(lngItem, 0)
          Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
          oWorksheet.Name = strItem
          SetAutoFilter .cboEach.Value, pjAutoFilterCustom, "equals", strItem
          'copy data
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus = "Creating Worksheet for " & strItem & "..."
          cptCopyData oWorksheet, lngHeaderRow
          If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
          'add legend
          If blnPerformanceTest Then t = GetTickCount
          cptAddLegend oWorksheet, dtStatus
          If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
          
          'final formatting
          cptFinalFormats oWorksheet
          
          oWorksheet.Calculate
          
          If blnLocked Then 'protect the sheet
            oWorksheet.Protect Password:="NoTouching!", AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
          
          .lblStatus = "Creating Worksheet for " & strItem & "...done"

        End If
      Next lngItem
      
      'todo: save workbook
      
    ElseIf .cboCreate.Value = "2" Then  'workbook for each
      For lngItem = 0 To .lboItems.ListCount - 1
        If .lboItems.Selected(lngItem) Then
          strItem = .lboItems.List(lngItem, 0)
          Set oWorkbook = oExcel.Workbooks.Add
          oExcel.Calculation = xlCalculationManual
          oExcel.ScreenUpdating = False
          Set oWorksheet = oWorkbook.Sheets(1)
          oWorksheet.Name = "Status Request"
          SetAutoFilter .cboEach.Value, pjAutoFilterCustom, "equals", .lboItems.List(lngItem, 0)
          'copy data
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus = "Creating Workbook for " & strItem & "..."
          cptCopyData oWorksheet, lngHeaderRow
          If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
          'add legend
          If blnPerformanceTest Then t = GetTickCount
          cptAddLegend oWorksheet, dtStatus
          If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
          
          'final formatting
          cptFinalFormats oWorksheet
          
          oWorksheet.Calculate
          
          If blnLocked Then 'protect the sheet
            oWorksheet.Protect Password:="NoTouching!", AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
                    
          .lblStatus = "Creating Workbook for " & strItem & "...done"
          
          'todo: save workbook for each
          
        End If
      Next lngItem
    End If
    .lblStatus.Caption = Choose(.cboCreate + 1, "Workbook", "Workbook", "Workbooks") & " Complete"
  End With

  oExcel.Calculation = xlCalculationAutomatic
  oExcel.ScreenUpdating = True
  oExcel.Visible = True
  
  
  
  GoTo exit_here

  'set up header
  If blnPerformanceTest Then t = GetTickCount

  'get selected fields for two non-standard fields
  With cptStatusSheet_frm
    If .cboCostTool.Value <> "<none>" Then strEVT = .cboEVT.Value Else strEVT = "SKIP" '<issue64>
    strEVP = .cboEVP.Value
  End With

  'set up header
  Set rstColumns = CreateObject("ADODB.Recordset")
  rstColumns.Fields.Append "FieldConstant", adBigInt
  rstColumns.Fields.Append "FieldName", adVarChar, 200
  rstColumns.Fields.Append "ColumnWidth", adInteger
  rstColumns.Fields.Append "NumberFormat", adVarChar, 200
  rstColumns.Fields.Append "EntryHeader", adBoolean
  rstColumns.Fields.Append "Alignment", adVarChar, 1
  rstColumns.Open
  
  'add first column
  lngField = FieldNameToFieldConstant("Unique ID")
  rstColumns.AddNew Array(0, 1, 2, 4, 5), Array(lngField, "Unique ID", 16, 0, "C")
  rstColumns.Update
  
  'add user fields (and save to adtg)
  vUserFields = cptStatusSheet_frm.lboExport.List()
  'save fields to adtg file
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  If UBound(vUserFields) > 0 Then
    With CreateObject("ADODB.Recordset")
      .Fields.Append "Field Constant", adVarChar, 255
      .Fields.Append "Custom Field Name", adVarChar, 255
      .Fields.Append "Local Field Name", adVarChar, 255
      .Open
      For lngItem = 0 To UBound(vUserFields)
        .AddNew Array(0, 1, 2), Array(vUserFields(lngItem, 0), vUserFields(lngItem, 1), vUserFields(lngItem, 2))
        If vUserFields(lngItem, 1) = strEVT Then GoTo next_field 'already included
        If vUserFields(lngItem, 1) = strEVP Then GoTo next_field 'already included
        rstColumns.AddNew Array(0, 1, 2, 4, 5), Array(vUserFields(lngItem, 0), vUserFields(lngItem, 1), 10, 0, "L")
        rstColumns.Update
next_field:
      Next lngItem
        .Update
       If Dir(strFileName) <> vbNullString Then Kill strFileName
       .Save strFileName
      .Close
    End With
  Else
    If Dir(strFileName) <> vbNullString Then Kill strFileName
  End If
  
  'add standard local fields, required EVT and EV%
  'some of these will be renamed later
  For Each vCol In Array("Name", _
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
    lngItem = lngItem + 1
    If CStr(vCol) <> "SKIP" Then
      rstColumns.AddNew Array(0, 1, 2, 5), Array(FieldNameToFieldConstant(vCol), vCol, 10, "C")
      If vCol = "Name" Then rstColumns.Update Array(2, 5), Array(60, "L")
      If vCol = "Duration" Then rstColumns.Update Array(2), Array(8)
      If vCol = "Total Slack" Then rstColumns.Update Array(2), Array(8)
      If vCol = strEVT Then rstColumns.Update Array(2), Array(5)
      If vCol = strEVP Then rstColumns.Update Array(2), Array(5)
      If vCol = "Notes" Then rstColumns.Update Array(2, 5), Array(45, "L")
    End If
  Next vCol

  'write the headers and size the columns
  vHeader = oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1), oWorksheet.Cells(lngHeaderRow, rstColumns.RecordCount)).Value
  rstColumns.MoveFirst
  lngItem = 1
  Do While Not rstColumns.EOF
    vHeader(1, lngItem) = rstColumns(1)
    oWorksheet.Columns(lngItem).EntireColumn.ColumnWidth = rstColumns(2)
    rstColumns.MoveNext
    lngItem = lngItem + 1
  Loop
  oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1), oWorksheet.Cells(lngHeaderRow, lngItem)).Value = vHeader
  'get columns needed later
  lngNameCol = oWorksheet.Rows(lngHeaderRow).Find("Name", lookat:=xlWhole).Column
  lngASCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Start", lookat:=xlWhole).Column
  lngAFCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlWhole).Column
  lngBaselineWorkCol = oWorksheet.Rows(lngHeaderRow).Find("Baseline Work", lookat:=xlWhole).Column
  lngRemainingWorkCol = oWorksheet.Rows(lngHeaderRow).Find("Remaining Work", lookat:=xlWhole).Column
  lngNotesCol = oWorksheet.Rows(lngHeaderRow).Find("Notes", lookat:=xlWhole).Column
  
  'format the header row
  With oWorksheet.Cells(lngHeaderRow, 1).Resize(, rstColumns.RecordCount)
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
    Set rstEach = CreateObject("ADODB.Recordset")
    rstEach.Fields.Append "EachItem", adVarChar, 200
    rstEach.Open
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
  Set oTasks = ActiveSelection.Tasks

  SelectBeginning
  Set aGroups = CreateObject("System.Collections.SortedList")
  Do
    On Error Resume Next
    Set oTask = ActiveCell.Task
    If Err.Number > 0 Then
      Err.Number = 0
      Err.Clear
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      Exit Do
    End If
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

    If oTask Is Nothing Then GoTo next_task
    If oTask.OutlineLevel = 0 Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If cptStatusSheet_frm.chkHide = True And Not oTask.GroupBySummary And Not oTask.Summary Then
      If oTask.ActualFinish <= CDate(cptStatusSheet_frm.txtHideCompleteBefore) Then GoTo next_task
    End If

    lngRow = lngRow + 1
    
    'retain-grouping
    If oTask.GroupBySummary Then
      aSummaries.Add lngRow 'capture for formatting
      aGroups.Add oTask.UniqueID, oTask.Name 'this will fail if view is task usage
      xlCells(lngRow, lngNameCol).IndentLevel = oTask.OutlineLevel - 1
      xlCells(lngRow, 1).Value = oTask.UniqueID
      xlCells(lngRow, lngNameCol).Value = oTask.Name
      lngTaskCount = lngTaskCount + 1
      GoTo next_task
    End If
    
    'add unique list of values to rstEach->lboEach in the selected field
    If cptStatusSheet_frm.cboCreate.Value <> "0" Then
      strItem = oTask.GetField(lngEach)
      If Len(strItem) > 0 And Not oTask.Summary Then
        If rstEach.RecordCount > 0 Then rstEach.MoveFirst
        rstEach.Find "EachItem='" & oTask.GetField(lngEach) & "'"
        If rstEach.EOF Then
          rstEach.AddNew Array(0), Array(strItem)
          rstEach.Update
        End If
      End If
    End If

    'get common data
    For lngCol = 1 To lngNameCol
      aTaskRow.Add oTask.GetField(FieldNameToFieldConstant(oWorksheet.Cells(lngHeaderRow, lngCol)))
    Next lngCol

    'indent the task name
    xlCells(lngRow, lngNameCol).IndentLevel = oTask.OutlineLevel + 1

    'write to Worksheet
    If oTask.Summary Then
      xlCells(lngRow, 1).Resize(, aTaskRow.Count).Value = aTaskRow.ToArray()
      aTaskRow.Clear
      aSummaries.Add lngRow
    Else
      For lngCol = lngNameCol + 1 To rstColumns.RecordCount
        'this gets overwritten by a formula; account for resource type at assignment level
        If oWorksheet.Cells(lngHeaderRow, lngCol) = "Baseline Work" Then
          aTaskRow.Add oTask.BaselineWork / 60
        ElseIf oWorksheet.Cells(lngHeaderRow, lngCol) = "Remaining Work" Then
          aTaskRow.Add oTask.RemainingWork / 60
        'elseif = new evp then get physical %
        'elseif = revised etc then get remaining work and divide
        Else
          aTaskRow.Add oTask.GetField(FieldNameToFieldConstant(oWorksheet.Cells(lngHeaderRow, lngCol)))
        End If
      Next lngCol

      'identify for formatting
      If oTask.Milestone Then aMilestones.Add lngRow Else aNormal.Add lngRow

      'write task data to sheet
      xlCells(lngRow, 1).Resize(, aTaskRow.Count).Value = aTaskRow.ToArray()
      aTaskRow.Clear

      'get assignment data for incomplete tasks
      If oTask.ActualFinish = "NA" Then
        'add remaining work formula '<issue58>
        If oTask.Assignments.Count > 0 Then
          xlCells(lngRow, lngRemainingWorkCol).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C:R" & lngRow + oTask.Assignments.Count & "C)"
        End If
        'get assignment data
        For Each oAssignment In oTask.Assignments
          lngRow = lngRow + 1
          aTaskRow.Add oAssignment.UniqueID
          If lngNameCol > 2 Then
            For lngCol = 2 To lngNameCol - 1
              aTaskRow.Add oTask.GetField(FieldNameToFieldConstant(oWorksheet.Cells(lngHeaderRow, lngCol)))
            Next
          End If

          'identify for formatting
          aAssignments.Add lngRow

          aTaskRow.Add oAssignment.ResourceName
          xlCells(lngRow, lngNameCol).IndentLevel = oTask.OutlineLevel + 2
          If oAssignment.ResourceType = pjResourceTypeWork Then
            xlCells(lngRow, lngBaselineWorkCol).Value = Val(oAssignment.BaselineWork) / 60     'Val() function prevents error when
            xlCells(lngRow, lngRemainingWorkCol).Value = Val(oAssignment.RemainingWork) / 60   'values are null or "" which is read as text
          Else
            xlCells(lngRow, lngBaselineWorkCol).Value = Val(oAssignment.BaselineWork)         'Val() function prevents error when
            xlCells(lngRow, lngRemainingWorkCol).Value = Val(oAssignment.RemainingWork)       'values are null or "" which is read as text
          End If
          xlCells(lngRow, 1).Resize(, aTaskRow.Count).Value = aTaskRow.ToArray()
          aTaskRow.Clear

        Next oAssignment
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
    DoEvents
    SelectCellDown
  Loop

  If blnPerformanceTest Then Debug.Print "capture task data: " & (GetTickCount - t) / 1000 & " >> " & Format(((GetTickCount - t) / 1000) / (lngRow - lngHeaderRow), "#0.00000") & " per task"

  If blnPerformanceTest Then t = GetTickCount
  'add New EV% after EV% - update rstColumns
  lngEVPCol = oWorksheet.Rows(lngHeaderRow).Find(strEVP).Column + 1
  oWorksheet.Columns(lngEVPCol - 1).Copy
  oWorksheet.Columns(lngEVPCol).Insert Shift:=xlToRight
  oWorksheet.Range(xlCells(lngHeaderRow + 1, lngEVPCol), xlCells(lngRow, lngEVPCol)).Cells.Locked = False
  xlCells(lngHeaderRow, lngEVPCol).Value = "New EV%"
  'todo: rstColumns.Insert lngEVPCol - 1, Array(0, "New EV%", 10)

  'add Revised ETC after Remaining Work - update rstColumns
  lngRemainingWorkCol = oWorksheet.Rows(lngHeaderRow).Find("Remaining Work", lookat:=xlWhole).Column 'don't use lngRemainingWorkCol because we've added a new column (and might add more)
  lngETCCol = lngRemainingWorkCol + 1
  oWorksheet.Columns(lngETCCol).Insert Shift:=xlToRight
  'lock task rows which will be a formula; unlock Assignment Rows
  'oWorksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).Cells.Locked = False
  oWorksheet.Range(xlCells(lngHeaderRow, lngRemainingWorkCol), xlCells(lngRow, lngRemainingWorkCol)).Copy
  oWorksheet.Range(xlCells(lngHeaderRow, lngETCCol), xlCells(lngRow, lngETCCol)).PasteSpecial xlAll
  oWorksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).Style = "Comma"
  oWorksheet.Columns(lngETCCol).ColumnWidth = 10
  xlCells(lngHeaderRow, lngETCCol).Value = "Revised ETC"
  oWorksheet.Calculate 'trigger Remaining Work formula to calculate
  oWorksheet.Range(xlCells(lngHeaderRow, lngRemainingWorkCol), xlCells(lngRow, lngRemainingWorkCol)).Copy
  xlCells(lngHeaderRow, lngRemainingWorkCol).PasteSpecial xlValues
  'todo: rstColumns.Insert lngETCCol - 1, Array(0, "Revised ETC", 10)
  If blnPerformanceTest Then Debug.Print "add columns: " & (GetTickCount - t) / 1000

  cptStatusSheet_frm.lblStatus = " Formatting rows..."
  Application.StatusBar = "Formatting rows..."
  If blnPerformanceTest Then t = GetTickCount
  'format rows
  'format summary tasks
  If aSummaries.Count > 0 Then '<issue16-17> added
    Set rSummaryTasks = xlCells(aSummaries(0), 1).Resize(, rstColumns.RecordCount)
    For vCol = 1 To aSummaries.Count - 1
      Set rSummaryTasks = oExcel.Union(rSummaryTasks, xlCells(aSummaries(vCol), 1).Resize(, rstColumns.RecordCount))
    Next vCol
    If Not rSummaryTasks Is Nothing Then
      rSummaryTasks.Interior.ThemeColor = xlThemeColorDark1
      rSummaryTasks.Interior.TintAndShade = -0.149998474074526
      rSummaryTasks.Font.Bold = True
    End If
  End If '</issue16-17>
  'format milestones
  If aMilestones.Count > 0 Then '<issue16-17> added
    Set rMilestones = xlCells(aMilestones(0), 1).Resize(, rstColumns.RecordCount)
    For vCol = 1 To aMilestones.Count - 1
      Set rMilestones = oExcel.Union(rMilestones, xlCells(aMilestones(vCol), 1).Resize(, rstColumns.RecordCount))
    Next vCol
    If Not rMilestones Is Nothing Then
      rMilestones.Font.ThemeColor = xlThemeColorAccent6
      rMilestones.Font.TintAndShade = -0.249977111117893
    End If
  End If '</issue16-17>
  'format normal tasks
  If aNormal.Count > 0 Then '<issue16-17> added
    Set rNormal = xlCells(aNormal(0), 1).Resize(, rstColumns.RecordCount) 'resize to entire used row
    For vCol = 1 To aNormal.Count - 1
      Set rNormal = oExcel.Union(rNormal, xlCells(aNormal(vCol), 1).Resize(, rstColumns.RecordCount))
    Next vCol
    If Not rNormal Is Nothing Then
      rNormal.Font.Color = RGB(32, 55, 100) '<issue54>
    End If
    If blnValidation Then
      'add data validation to new start
      lngCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Start", lookat:=xlWhole).Column
      For vCol = 0 To aNormal.Count - 1 'we are borrowing vCol to iterate row numbers
        With xlCells(aNormal(vCol), lngCol).Validation
          .Delete
          .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:=FormatDateTime(ActiveProject.ProjectStart, vbShortDate)
          .IgnoreBlank = True
          .InCellDropdown = True
          .InputTitle = "Date Only"
          .ErrorTitle = "Date Only"
          .InputMessage = "Please enter a date in format mm/dd/yyyy."
          .ErrorMessage = "Please enter a date in format mm/dd/yyyy."
          .ShowInput = True
          .ShowError = True
        End With
      Next vCol
      'add data validation to new finish
      lngCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlWhole).Column
      For vCol = 0 To aNormal.Count - 1 'we are borrowing vCol to iterate row numbers
        With xlCells(aNormal(vCol), lngCol).Validation
          .Delete
          .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:=FormatDateTime(ActiveProject.ProjectStart, vbShortDate)
          .IgnoreBlank = True
          .InCellDropdown = True
          .InputTitle = "Date Only"
          .ErrorTitle = "Date Only"
          .InputMessage = "Please enter a date in format mm/dd/yyyy."
          .ErrorMessage = "Please enter a date in format mm/dd/yyyy."
          .ShowInput = True
          .ShowError = True
        End With
      Next vCol
      'add data validation to new ev%
      lngCol = oWorksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
      For vCol = 0 To aNormal.Count - 1 'we are borrowing vCol to iterate row numbers
        With xlCells(aNormal(vCol), lngCol).Validation
          .Delete
          .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
          .IgnoreBlank = True
          .InCellDropdown = True
          .InputTitle = "Number Only"
          .ErrorTitle = "Number Only"
          .InputMessage = "Please enter a number greater than or equal to zero. Decimals permitted."
          .ErrorMessage = "Please enter a number greater than or equal to zero. Decimals permitted."
          .ShowInput = True
          .ShowError = True
        End With
      Next vCol
    End If 'blnValidation
  End If '</issue16-17>
  'format assignments
  If aAssignments.Count > 0 Then '<issue16-17> added
    Set rAssignments = xlCells(aAssignments(0), 1).Resize(, rstColumns.RecordCount)
    For vCol = 1 To aAssignments.Count - 1 'we are borrowing vCol to iterate row numbers
      Set rAssignments = oExcel.Union(rAssignments, xlCells(aAssignments(vCol), 1).Resize(, rstColumns.RecordCount))
    Next vCol
    If Not rAssignments Is Nothing Then rAssignments.Font.Italic = True
    If blnValidation Then
      'add data validation to new ev%
      lngCol = oWorksheet.Rows(lngHeaderRow).Find("Revised ETC", lookat:=xlWhole).Column
      For vCol = 0 To aAssignments.Count - 1
        With xlCells(aAssignments(vCol), lngCol).Validation
          .Delete
          .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
          .IgnoreBlank = True
          .InCellDropdown = True
          .InputTitle = "Number Only"
          .ErrorTitle = "Number Only"
          .InputMessage = "Please enter a number greater than or equal to zero. Decimals permitted."
          .ErrorMessage = "Please enter a number greater than or equal to zero. Decimals permitted."
          .ShowInput = True
          .ShowError = True
        End With
      Next vCol
    End If 'blnValidation
  End If '</issue16-17>
  'format completed '<issue58>
  If aCompleted.Count > 0 Then
    'format the entire row - assignments are skipped on completd tasks
    Set rCompleted = xlCells(aCompleted(0), 1).Resize(, rstColumns.RecordCount)
    For vCol = 1 To aCompleted.Count - 1 'we are borrowing vCol to iterate row numbers
      Set rCompleted = oExcel.Union(rCompleted, xlCells(aCompleted(vCol), 1).Resize(, rstColumns.RecordCount))
    Next vCol
    If Not rCompleted Is Nothing Then
      rCompleted.Font.Italic = True
      rCompleted.Font.ColorIndex = 16
    End If
    'update ev% complete
    lngCol = oWorksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
    Set rCompleted = xlCells(aCompleted(0), lngCol)
    For vCol = 1 To aCompleted.Count - 1 'we are borrowing vCol to iterate row numbers
      Set rCompleted = oExcel.Union(rCompleted, xlCells(aCompleted(vCol), lngCol))
    Next vCol
    rCompleted = 1
  End If '</issue58>
  If blnPerformanceTest Then Debug.Print "format rows: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'format common borders
  Set rng = oWorksheet.Range(xlCells(lngHeaderRow, 1), xlCells(lngRow, rstColumns.RecordCount))
  rng.BorderAround xlContinuous, xlThin
  rng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
  rng.Borders(xlInsideHorizontal).Weight = xlThin
  If blnPerformanceTest Then Debug.Print "format common borders: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'rename headers
  Set rng = xlCells(lngHeaderRow, 1).Resize(, rstColumns.RecordCount)
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
  DoEvents

  If blnPerformanceTest Then t = GetTickCount
  'define bulk column ranges for formatting
  rstColumns.MoveFirst
  Do While Not rstColumns.EOF

    'get range of dates
    If Len(cptRegEx(CStr(rstColumns("FieldName")), "Start|Finish")) > 0 Then
      If rDates Is Nothing Then
        Set rDates = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rDates = oExcel.Union(rDates, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'get range of work
    If Len(cptRegEx(CStr(rstColumns("FieldName")), "Baseline Work|Remaining Work")) > 0 Then
      If rWork Is Nothing Then
        Set rWork = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rWork = oExcel.Union(rWork, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'get range of centered
    If rstColumns("Alignment") = "C" Then
      If rCentered Is Nothing Then
        Set rCentered = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow)
      Else
        Set rCentered = oExcel.Union(rCentered, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
      End If
    End If
    'format entry headers and columns
    If rstColumns("EntryHeader") = 1 Then 'if the column we're working on is included in the list of entry headers, then...
      If rEntry Is Nothing Then 'first iteration sets range
        Set rEntry = xlCells(lngHeaderRow, lngCol + 1)
        Set rMedium = xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow) 'medium = border thickness
        Set rLockedCells = rMedium 'entry cells are unlocked cells
      Else 'second and following iterations extend the range
        Set rEntry = oExcel.Union(rEntry, xlCells(lngHeaderRow, lngCol + 1))
        Set rMedium = oExcel.Union(rMedium, xlCells(lngHeaderRow + 1, lngCol + 1).Resize(rowsize:=lngRow - lngHeaderRow))
        Set rLockedCells = rMedium
      End If
    End If
    rstColumns.MoveNext
  Loop 'rstColumns
  If blnPerformanceTest Then Debug.Print "define bulk ranges for formatting: " & (GetTickCount - t) / 1000

  If blnPerformanceTest Then t = GetTickCount
  'apply bulk formatting
  rDates.NumberFormat = "m/d/yy;@"
  rDates.HorizontalAlignment = xlCenter
  rDates.Replace "NA", ""
  'format work columns
  rWork.Style = "Comma"
  rCentered.HorizontalAlignment = xlCenter
  If Not rEntry Is Nothing Then 'todo: rEntry doesn't exist yet
    rEntry.Interior.ThemeColor = xlThemeColorAccent3
    rEntry.Interior.TintAndShade = 0.399975585192419
    rEntry.Font.ColorIndex = xlAutomatic
    rEntry.BorderAround xlContinuous, xlMedium
  End If
  If Not rLockedCells Is Nothing Then 'todo: rLockedCells doesn't exist yet
    rLockedCells.SpecialCells(xlCellTypeBlanks).Locked = False
  End If

  If Not rMedium Is Nothing Then 'todo: rMedium doesn't exist yet
    rMedium.BorderAround xlContinuous, xlMedium
  End If
  lngCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlPart).Column
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
  DoEvents
  cptStatusSheet_frm.lblProgress.Width = (1 / 100) * cptStatusSheet_frm.lblStatus.Width
  lngConditionalFormats = 18 'bash: "grep -c 'lngFormatCondition = lngFormatCondition + 1' Status/cptStatusSheet_bas.bas"
  'capture status date address
  strStatusDate = oWorksheet.Range("STATUS_DATE").Address(True, True)

  'attempt to speed up
  oExcel.EnableEvents = False
  oWorksheet.DisplayPageBreaks = False

new_start:
  'define range for new start
  xlCells(lngHeaderRow, 1).AutoFilter
  Set rngAll = oWorksheet.Range(xlCells(lngHeaderRow, 1).End(xlToRight), xlCells(lngHeaderRow, 1).End(xlDown))
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AS]
  rngAll.AutoFilter Field:=lngASCol, Criteria1:="="
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52-no cells found>
  Set rng = oWorksheet.Range(xlCells(lngHeaderRow + 1, lngASCol), xlCells(lngRow, lngASCol)).SpecialCells(xlCellTypeVisible)
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
  DoEvents

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
  DoEvents
  
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
  DoEvents

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
  DoEvents

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
  DoEvents

new_finish: '<issue52>
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  If oWorksheet.AutoFilterMode Then oWorksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lngAFCol, Criteria1:="="
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52>
  Set rng = oWorksheet.Range(xlCells(lngHeaderRow + 1, lngAFCol), xlCells(lngRow, lngAFCol)).SpecialCells(xlCellTypeVisible)
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
  DoEvents

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
  DoEvents

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
  DoEvents

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
  DoEvents

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
  DoEvents

ev_percent:
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oWorksheet.AutoFilterMode Then oWorksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for task rows [blue font]
  rngAll.AutoFilter Field:=lngNameCol, Criteria1:=RGB(32, 55, 100), Operator:=xlFilterFontColor
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52-noCellsFound>
  Set rng = oWorksheet.Range(xlCells(lngHeaderRow + 1, lngEVPCol), xlCells(lngRow, lngEVPCol)).SpecialCells(xlCellTypeVisible)
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
  DoEvents

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
  DoEvents

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
  DoEvents

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
  DoEvents

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
  DoEvents

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
  DoEvents

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
  DoEvents

revised_etc:
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 '<issue52>
  If oWorksheet.AutoFilterMode Then oWorksheet.ShowAllData
  xlCells(lngHeaderRow, 1).AutoFilter
  'filter for Task
  rngAll.AutoFilter Field:=lngETCCol, Operator:=xlFilterAutomaticFontColor
  '...with blank Actual Start dates [blank AF]
  rngAll.AutoFilter Field:=lngRemainingWorkCol, Operator:=xlFilterNoFill
  'add conditions only to blank cells in the column
  On Error Resume Next '<issue52>
  Set rng = oWorksheet.Range(xlCells(lngHeaderRow + 1, lngETCCol), xlCells(lngRow, lngETCCol)).SpecialCells(xlCellTypeVisible)
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
    DoEvents
  Next lngRow
  lngFormatCondition = lngFormatCondition + 1
  cptStatusSheet_frm.lblStatus.Caption = "Adding conditionanl formats - New ETC (4/4)"
  cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngConditionalFormats) * cptStatusSheet_frm.lblStatus.Width
  Application.StatusBar = "Adding conditionanl formats - New ETC (4/4)"
  DoEvents

evt_vs_evp:
  If cptStatusSheet_frm.cboCostTool <> "<none>" Then '<issue64>
    'evt vs evp checks
    If cptStatusSheet_frm.cboCostTool = "COBRA" Then
      'EVT = E 50/50
      'todo: just make EVP a formula in this case
      'EVT = F 0/100
      'todo: just make EVP a formula in this case
    ElseIf cptStatusSheet_frm.cboCostTool.Value = "MPM" Then
      'EVT =1 0/100
      'todo: just make EVP a formula in this case
      'EVT =4 50/50
      'todo: just make EVP a formula in this case
    'todo: what about forProject EVMS
    End If
  End If '</issue64>
  If blnPerformanceTest Then Debug.Print "apply conditional formatting " & (GetTickCount - t) / 1000

  Debug.Print lngFormatCondition & " format conditions applied."

conditional_formatting_skipped:

  'unlock Revised ETC at assignment level
  lngCol = oWorksheet.Rows(lngHeaderRow).Find("Revised ETC", lookat:=xlWhole).Column
  For lngItem = 0 To aAssignments.Count - 1
    xlCells(aAssignments(lngItem), lngCol).Locked = False
  Next lngItem

  'optionallly set reference to Outlook and prepare to email
  blnEmail = cptStatusSheet_frm.chkSendEmails = True
  If blnEmail Then
    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If oOutlook Is Nothing Then Set oOutlook = CreateObject("Outlook.Application")
    cptStatusSheet_frm.lblStatus.Caption = "Getting Outlook..."
    Application.StatusBar = "Getting Outlook..."
    DoEvents
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
  
  'create single Workbook
  If cptStatusSheet_frm.cboCreate.Value = "0" Then
  
    oExcel.Visible = True '<issue81> - move this below if option = (0|other)
    oExcel.WindowState = xlMaximized
    oExcel.ScreenUpdating = True
    
    If oWorksheet.AutoFilterMode Then oWorksheet.ShowAllData
    
    oExcel.ActiveWindow.ScrollColumn = 1
    oExcel.ActiveWindow.ScrollRow = 1 '<issue54>
    xlCells(lngHeaderRow + 1, lngNameCol + 1).Select
    oExcel.ActiveWindow.FreezePanes = True
    'prettify the task name column
    oWorksheet.Columns(lngNameCol).AutoFit
  
    If blnLocked Then 'protect the sheet
      oWorksheet.Protect Password:="NoTouching!", AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
      oWorksheet.EnableSelection = xlNoRestrictions
    End If
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
      oWorkbook.SaveAs strDir & strFileName, 51
    Else
      oWorkbook.SaveAs strDir & strFileName, 51
    End If
    If blnEmail Then
      Set oMailItem = oOutlook.CreateItem(0) '0 = oloMailItem
      oMailItem.Attachments.Add strDir & strFileName
      oMailItem.Subject = "Status Request - " & Format(dtStatus, "yyyy-mm-dd")
      oMailItem.Display False
    End If
  Else
  
    xlCells(lngHeaderRow + 1, lngNameCol + 1).Select
    oExcel.ActiveWindow.FreezePanes = True
    'prettify the task name column
    oWorksheet.Columns(lngNameCol).AutoFit
  
    'cycle through each option and create sheet
    lngItem = 0
    rstEach.MoveFirst
    Do While Not rstEach.EOF
      cptStatusSheet_frm.lblStatus.Caption = "Extracting data for " & rstEach(0) & "..."
      Application.StatusBar = "Extracting data for " & rstEach(0) & "..."
      DoEvents
      lngLastRow = oWorkbook.Worksheets(1).[A8].End(xlDown).Row
      lngRow = oWorkbook.Worksheets(1).[A8].End(xlDown).Row
      oWorkbook.Sheets(1).Copy After:=oWorkbook.Sheets(oWorkbook.Sheets.Count) 'Set = Copy( doesn't work
      Set oWorksheet = oWorkbook.Sheets(oWorkbook.Sheets.Count)
      oWorksheet.Name = rstEach(0)
      SetAutoFilter FieldName:=cptStatusSheet_frm.cboEach, FilterType:=pjAutoFilterIn, Criteria1:=rstEach(0)
      'get array of task and assignment unique ids
      oWorksheet.Cells(lngRow + 2, 1).Value = "KEEP"
      'get group by summaries
      SelectBeginning
      Do
        On Error Resume Next
        Set oTask = ActiveCell.Task
        If Not oTask Is Nothing Then
          If oTask.GroupBySummary Or oTask.Summary Then
            oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Offset(1, 0) = oTask.UniqueID 'aGroups.getValueList()(aGroups.indexOfKey(oTask.Name))
          End If
        'todo: else skip it and move to next
        End If
        If Err.Number > 0 Then
          Err.Number = 0
          Err.Clear
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          Exit Do
        End If
        If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
        SelectCellDown
      Loop
      'now get assignment uids to keep
      SelectAll
      For Each oTask In ActiveSelection.Tasks
        oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = oTask.UniqueID
        For Each oAssignment In oTask.Assignments
          oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Offset(1, 0) = oAssignment.UniqueID
        Next oAssignment
      Next oTask
      'name the range of uids to keep
      Set rngKeep = oWorksheet.Cells(lngRow + 2, 1)
      Set rngKeep = oWorksheet.Range(rngKeep, rngKeep.End(xlDown))
      oWorkbook.Names.Add Name:="KEEP", RefersToR1C1:="='" & rstEach(0) & "'!" & rngKeep.Address(True, True, xlR1C1)
      'add a formula to find which rows to keep
      lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(1, 1).Column
      Set rngKeep = oWorksheet.Range(oWorksheet.Cells(lngHeaderRow + 1, lngLastCol), oWorksheet.Cells(lngRow, lngLastCol))
      rngKeep(1).Offset(-1, 0).Value = "KEEP"
      rngKeep.Formula = "=IFERROR(VLOOKUP(A" & lngHeaderRow + 1 & ",KEEP,1,FALSE),""DELETE"")"
      oExcel.Calculate
      'oWorksheet.Cells(lngHeaderRow, 1).AutoFilter
      lngLastRow = lngRow
      For lngRow = lngLastRow To lngHeaderRow + 1 Step -1
        If oWorksheet.Cells(lngRow, rngKeep.Column) = "DELETE" Then oWorksheet.Rows(lngRow).Delete Shift:=xlUp
        cptStatusSheet_frm.lblStatus.Caption = "Creating " & rstEach(0) & "...(" & Format(((lngLastRow - lngRow) / (lngLastRow - lngHeaderRow)), "0%") & ")"
        cptStatusSheet_frm.lblProgress.Width = ((lngLastRow - lngRow) / (lngLastRow - lngHeaderRow)) * cptStatusSheet_frm.lblStatus.Width
        Application.StatusBar = "Creating " & rstEach(0) & "...(" & Format(((lngLastRow - lngRow) / (lngLastRow - lngHeaderRow)), "0%") & ")"
        DoEvents
      Next lngRow
      cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
      oWorksheet.Cells(lngHeaderRow, 1).Select
      oExcel.Selection.AutoFilter
      oWorksheet.Columns(lngLastCol).Delete
      oWorksheet.Range("KEEP").Clear
      oWorkbook.Names("KEEP").Delete
      oWorksheet.[B1].Select
      If blnLocked Then
        oWorksheet.Protect Password:="NoTouching!", AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
        oWorksheet.EnableSelection = xlNoRestrictions
      End If
      DoEvents
      rstEach.MoveNext
    Loop 'rstEach

    oExcel.ScreenUpdating = True
    oExcel.Calculation = True

    'handle for each
    If cptStatusSheet_frm.cboCreate.Value = "1" Then  'Worksheet for each
      'oWorkbook.SaveAs strDir & strFileName, 51 '</issue80>
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
        oWorkbook.SaveAs strDir & strFileName, 51
      Else
        oWorkbook.SaveAs strDir & strFileName, 51
      End If
      If blnEmail Then
        Set oMailItem = oOutlook.CreateItem(0) '0 = oloMailItem
        oMailItem.Attachments.Add strDir & strFileName
        oMailItem.Subject = "Status Request - " & Format(dtStatus, "yyyy-mm-dd")
        oMailItem.Display False
      End If
    ElseIf cptStatusSheet_frm.cboCreate.Value = "2" Then 'Workbook for each
      rstEach.MoveLast 'start from the bottom
      Do While Not rstEach.BOF
        cptStatusSheet_frm.lblStatus.Caption = "Saving " & rstEach(0) & "..."
        Application.StatusBar = "Saving " & rstEach(0) & "..."
        oWorkbook.Sheets(CStr(rstEach(0))).Copy
        On Error Resume Next
        If Dir(strDir & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx")) <> vbNullString Then Kill strDir & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx")
        If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
        'account for if the file exists and is open in the background
        If Dir(strDir & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx")) <> vbNullString Then  'delete failed, rename with timestamp
          strMsg = "'" & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx") & "' already exists, and is likely open." & vbCrLf
          strFileName = Replace(Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx"), ".xlsx", "_" & Format(Now, "hh-nn-ss") & ".xlsx")
          strMsg = strMsg & "The file you are now creating will be named '" & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx") & "'"
          MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
          oExcel.ActiveWorkbook.SaveAs strDir & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx"), 51
        Else
          oExcel.ActiveWorkbook.SaveAs strDir & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx"), 51
        End If
        If blnLocked Then 'protect the worksheet
          oExcel.ActiveWorkbook.Worksheets(1).Protect Password:="NoTouching!", AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
          oExcel.ActiveWorkbook.Worksheets(1).EnableSelection = xlNoRestrictions
        End If
        If blnEmail Then
          Set oMailItem = oOutlook.CreateItem(0) '0 = olMailItem
          oMailItem.Attachments.Add strDir & Replace(strFileName, ".xlsx", "_" & rstEach(0) & ".xlsx")
          oMailItem.Subject = "Status Request [" & rstEach(0) & "] " & Format(dtStatus, "yyyy-mm-dd")
          oMailItem.Display False
        End If
        rstEach.MovePrevious
      Loop
      oWorkbook.Close False
    End If

    cptStatusSheet_frm.lblStatus.Caption = "Wrapping up..."
    Application.StatusBar = "Wrapping up..."
    DoEvents
    
    'reset autofilter
    strFieldName = cptStatusSheet_frm.cboEach.Value
    strCriteria = ""
    rstEach.MoveFirst
    Do While Not rstEach.EOF
      strCriteria = strCriteria & rstEach(0) & Chr$(9)
      rstEach.MoveNext
    Loop
    strCriteria = Left(strCriteria, Len(strCriteria) - 1)
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
    SelectBeginning

  End If
  If blnPerformanceTest Then Debug.Print "save Workbook: " & (GetTickCount - t) / 1000
  If blnPerformanceTest Then Debug.Print "</=====PERFORMANCE TEST=====>"

  cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
  cptStatusSheet_frm.lblStatus.Caption = " Complete."
  Application.StatusBar = "Complete."
  MsgBox "Status Sheet(s) Created", vbInformation + vbOKOnly, "ClearPlan Status Sheet"
  oExcel.Visible = True
  cptStatusSheet_frm.lblStatus.Caption = " Ready..."
  Application.StatusBar = "Ready..."
  DoEvents

exit_here:
  On Error Resume Next
  Application.DefaultDateFormat = lngDateFormat
  ActiveProject.SpaceBeforeTimeLabels = blnSpace
  ActiveProject.DayLabelDisplay = lngDayLabelDisplay
  If oExcel.Workbooks.Count > 0 Then oExcel.Calculation = xlAutomatic
  oExcel.ScreenUpdating = True
  oExcel.EnableEvents = True
  Set rCompleted = Nothing
  Set aCompleted = Nothing
  Set aGroups = Nothing
  Set rngKeep = Nothing
  Application.StatusBar = ""
  cptSpeed False
  Set oTasks = Nothing
  Set oTask = Nothing
  Set oAssignment = Nothing
  Set oExcel = Nothing
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing
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
  If rstEach.State Then rstEach.Close
  Set rstEach = Nothing
  Set aTaskRow = Nothing
  If rstColumns.State Then rstColumns.Close
  Set rstColumns = Nothing
  Set xlCells = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDoc = Nothing
  Set oWord = Nothing
  Set oSel = Nothing
  Set oETemp = Nothing
  Exit Sub

err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCreateStatusSheet", Err, Erl)
  If Not oExcel Is Nothing Then
    If Not oWorkbook Is Nothing Then oWorkbook.Close False
    oExcel.Quit
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
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  lngItem = 0
  If cptStatusSheet_frm.lboExport.ListCount > 0 Then
    For lngItem = 0 To cptStatusSheet_frm.lboExport.ListCount - 1
      TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(cptStatusSheet_frm.lboExport.List(lngItem, 0)), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    Next lngItem
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Name", Title:="Task Name / Scope", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Remaining Duration", Title:="", Width:=12, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Baseline Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Baseline Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False, ShowAddNewColumn:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Start", Title:="Forecast Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Finish", Title:="Forecast Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Actual Start", Title:="New Forecast/ Actual Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Actual Finish", Title:="New Forecast/ Actual Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If cptStatusSheet_frm.cboEVT <> 0 Then
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=cptStatusSheet_frm.cboEVT.Value, Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  If cptStatusSheet_frm.cboEVP <> 0 Then
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=cptStatusSheet_frm.cboEVP.Value, Title:="EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=cptStatusSheet_frm.cboEVP.Value, Title:="New EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Baseline Work", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Remaining Work", Title:="Previous ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Remaining Work", Title:="Revised ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableApply Name:="cptStatusSheet Table"

  'reset the filter
  FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Actual Finish", test:="equals", Value:="NA", ShowInMenu:=False, showsummarytasks:=True
  If cptStatusSheet_frm.chkHide And IsDate(cptStatusSheet_frm.txtHideCompleteBefore) Then
    FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", newfieldname:="Actual Finish", test:="is greater than or equal to", Value:=cptStatusSheet_frm.txtHideCompleteBefore, Operation:="Or", showsummarytasks:=True
  End If
  FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", newfieldname:="Active", test:="equals", Value:="Yes", ShowInMenu:=False, showsummarytasks:=True, parenthesis:=True
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

Private Sub cptAddLegend(ByRef oWorksheet As Worksheet, dtStatus As Date)
  'objects
  'strings
  'longs
  'todo: delete this
  Dim lngCol As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    
  oWorksheet.Cells(1, 1).Value = "Status Date:"
  oWorksheet.Cells(1, 1).Font.Bold = True
  oWorksheet.Cells(1, 2) = FormatDateTime(dtStatus, vbShortDate)
  oWorksheet.Names.Add "STATUS_DATE", oWorksheet.[B1]
  oWorksheet.Cells(1, 2).Font.Bold = True
  oWorksheet.Cells(1, 2).Font.Size = 14
  'current
  oWorksheet.Cells(3, 1).Style = "Input" '<issue58>
  oWorksheet.Cells(3, 2) = "Task is active or within current status window. Cell requires update."
  'within two weeks
  oWorksheet.Cells(4, 1).Style = "Neutral" '<issue58>
  oWorksheet.Cells(4, 1).BorderAround xlContinuous, xlThin, , -8421505
  oWorksheet.Cells(4, 2) = "Task is within two week look-ahead. Please review forecast dates."
  'complete
  oWorksheet.Cells(5, 1) = "AaBbCc"
  oWorksheet.Cells(5, 1).Font.Italic = True
  oWorksheet.Cells(5, 1).Font.ColorIndex = 16
  oWorksheet.Cells(5, 2) = "Task is complete."
  'summary
  oWorksheet.Cells(6, 1) = "AaBbCc"
  oWorksheet.Cells(6, 1).Font.Bold = True
  oWorksheet.Cells(6, 1).Interior.ThemeColor = xlThemeColorDark1
  oWorksheet.Cells(6, 1).Interior.TintAndShade = -0.149998474074526
  oWorksheet.Cells(6, 2) = "MS Project Summary Task (Rollup).  No update required."

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptAddLegend", Err, Erl)
  Resume exit_here

End Sub

Private Sub cptCopyData(ByRef oWorksheet As Worksheet, lngHeaderRow As Long)
  'objects
  Dim oComment As Excel.Comment
  Dim oEVTRange As Excel.Range
  Dim oCompleted As Excel.Range
  Dim oMilestoneRange As Excel.Range
  Dim oClearRange As Excel.Range
  Dim oSummaryRange As Excel.Range
  Dim oDateValidationRange As Excel.Range
  Dim oTwoWeekWindowRange As Excel.Range
  Dim oTask As Task
  'strings
  Dim strEVTList As String
  'longs
  Dim lngEVTCol As Long
  Dim lngLastCol As Long
  Dim lngETCCol As Long
  Dim lngTask As Long
  Dim lngRow As Long
  Dim lngNameCol As Long
  Dim lngTasks As Long
  Dim lngCol As Long
  Dim lngBLSCol As Long
  Dim lngBLFCol As Long
  Dim lngASCol As Long
  Dim lngAFCol As Long
  Dim lngEVPCol As Long
  'integers
  'doubles
  'booleans
  Dim blnValidation As Boolean
  Dim blnConditionalFormats As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate
  blnValidation = cptStatusSheet_frm.chkValidation
  blnConditionalFormats = cptStatusSheet_frm.chkAddConditionalFormats
  ActiveWindow.TopPane.Activate
try_again:
  SelectAll
  EditCopy
  oWorksheet.Application.Wait 5000
  On Error Resume Next
  oWorksheet.Paste oWorksheet.Cells(lngHeaderRow, 1), False
  If Err.Number = 1004 Then 'try again
    EditCopy
    oWorksheet.Application.Wait 5000
    oWorksheet.Paste oWorksheet.Cells(lngHeaderRow, 1), False
    oWorksheet.Application.Wait 5000
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  oWorksheet.Application.Wait 5000
  oWorksheet.Cells.WrapText = False
  oWorksheet.Application.ActiveWindow.Zoom = 85
  oWorksheet.Cells.Font.Name = "Calibri"
  oWorksheet.Cells.Font.Size = 11
  oWorksheet.Rows(lngHeaderRow).Font.Bold = True
  oWorksheet.Columns.AutoFit
  'format the columns
  For lngCol = 1 To ActiveSelection.FieldIDList.Count
    oWorksheet.Columns(lngCol).ColumnWidth = ActiveProject.TaskTables("cptStatusSheet Table").TableFields(lngCol + 1).Width + 2
    oWorksheet.Cells(lngHeaderRow, lngCol).WrapText = True
    If InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Start") > 0 Then
      oWorksheet.Columns(lngCol).Replace "NA", ""
      oWorksheet.Columns(lngCol).NumberFormat = "mm/dd/yyyy"
    ElseIf InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Finish") > 0 Then
      oWorksheet.Columns(lngCol).Replace "NA", ""
      oWorksheet.Columns(lngCol).NumberFormat = "mm/dd/yyyy"
    ElseIf InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Work") > 0 Or InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "ETC") > 0 Then
      oWorksheet.Columns(lngCol).Style = "Comma"
    End If
    If Len(cptRegEx(oWorksheet.Cells(lngHeaderRow, lngCol), "New|Revised")) > 0 Then
      If oEntryHeaderRange Is Nothing Then
        Set oEntryHeaderRange = oWorksheet.Cells(lngHeaderRow, lngCol)
      Else
        Set oEntryHeaderRange = oWorksheet.Application.Union(oEntryHeaderRange, oWorksheet.Cells(lngHeaderRow, lngCol))
      End If
    End If
    'todo: replace "h"
  Next lngCol
  
  'format the header
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  If lngLastCol > ActiveProject.TaskTables("cptStatusSheet Table").TableFields.Count + 10 Then GoTo try_again
  oWorksheet.Columns(lngLastCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
  oWorksheet.Columns(lngLastCol + 1).ColumnWidth = 40
  oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(0, 1).Value = "Comment / Action / Impact"
  With oWorksheet.Cells(lngHeaderRow, 1).Resize(, ActiveProject.TaskTables(ActiveProject.CurrentTable).TableFields.Count)
    .Interior.ThemeColor = xlThemeColorLight2
    .Interior.TintAndShade = 0
    .Font.ThemeColor = xlThemeColorDark1
    .Font.TintAndShade = 0
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
  End With
  
  'format the data rows
  lngNameCol = oWorksheet.Rows(lngHeaderRow).Find("Task Name / Scope", lookat:=xlWhole).Column
  lngASCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Start", lookat:=xlPart).Column
  lngAFCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlPart).Column
  lngEVPCol = oWorksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
  lngEVTCol = oWorksheet.Rows(lngHeaderRow).Find("EVT", lookat:=xlWhole).Column
  lngETCCol = oWorksheet.Rows(lngHeaderRow).Find("Revised ETC", lookat:=xlWhole).Column
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  lngTasks = ActiveSelection.Tasks.Count
  lngTask = 0
  For Each oTask In ActiveSelection.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    'todo: use Task Usage view if a group is applied there are no duplicate UIDs
    lngRow = oWorksheet.Columns(1).Find(oTask.UniqueID, lookat:=xlWhole).Row
    If oTask.Summary Then
      If oSummaryRange Is Nothing Then
        Set oSummaryRange = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oSummaryRange = oWorksheet.Application.Union(oSummaryRange, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      If oClearRange Is Nothing Then
        Set oClearRange = oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oClearRange = oWorksheet.Application.Union(oClearRange, oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      GoTo next_task
    End If
    If oTask.Milestone Then
      If oMilestoneRange Is Nothing Then
        Set oMilestoneRange = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oMilestoneRange = oWorksheet.Application.Union(oMilestoneRange, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
'      If oClearRange Is Nothing Then
'        Set oClearRange = oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol))
'      Else
'        Set oClearRange = oWorksheet.Application.Union(oClearRange, oWorksheet.Range(oWorksheet.Cells(lngRow, lngNameCol + 1), oWorksheet.Cells(lngRow, lngLastCol)))
'      End If
      GoTo next_task
    End If
    'format completed
    If IsDate(oTask.ActualFinish) Then
      If oCompleted Is Nothing Then
        Set oCompleted = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oCompleted = oWorksheet.Application.Union(oCompleted, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      GoTo next_task
    End If
    'capture status formating:
    'tasks requiring status:
    If oTask.Start < dtStatus And Not IsDate(oTask.ActualStart) Then
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
    End If
    If oTask.Finish <= dtStatus And Not IsDate(oTask.ActualFinish) Then
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
      Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
    End If
    If IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish) Then
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngEVPCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
      End If
    End If
    'two week window
    If oTask.Start > dtStatus And oTask.Start <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
    End If
    If oTask.Finish > dtStatus And oTask.Finish <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
    End If
    
    'capture data validation
    If blnValidation Then
      If Not IsDate(oTask.ActualStart) Then
        If oDateValidationRange Is Nothing Then
          Set oDateValidationRange = oWorksheet.Cells(lngRow, lngASCol)
        Else
          Set oDateValidationRange = oWorksheet.Application.Union(oDateValidationRange, oWorksheet.Cells(lngRow, lngASCol))
        End If
      End If
      If Not IsDate(oTask.ActualFinish) Then
        If oDateValidationRange Is Nothing Then
          Set oDateValidationRange = oWorksheet.Cells(lngRow, lngAFCol)
        Else
          Set oDateValidationRange = oWorksheet.Application.Union(oDateValidationRange, oWorksheet.Cells(lngRow, lngAFCol))
        End If
      End If
      If oNumberValidationRange Is Nothing Then
        Set oNumberValidationRange = oWorksheet.Cells(lngRow, lngEVPCol)
      Else
        Set oNumberValidationRange = oWorksheet.Application.Union(oNumberValidationRange, oWorksheet.Cells(lngRow, lngEVPCol))
      End If
    End If 'blnValidation
    
    If oEVTRange Is Nothing Then
      Set oEVTRange = oWorksheet.Cells(lngRow, lngEVTCol)
    Else
      Set oEVTRange = oWorksheet.Application.Union(oEVTRange, oWorksheet.Cells(lngRow, lngEVTCol))
    End If
    
    If oTask.Assignments.Count > 0 Then
      cptGetAssignmentData oTask, oWorksheet, lngRow, lngHeaderRow, lngNameCol, lngETCCol - 1
    End If
    
    'todo: apply conditional formatting
    
    oWorksheet.Columns(1).AutoFit

next_task:
    lngTask = lngTask + 1
    cptStatusSheet_frm.lblProgress.Width = (lngTask / lngTasks) * cptStatusSheet_frm.lblStatus.Width
  Next oTask
  
  If Not oClearRange Is Nothing Then oClearRange.ClearContents
  If Not oSummaryRange Is Nothing Then
    oSummaryRange.Interior.ThemeColor = xlThemeColorDark1
    oSummaryRange.Interior.TintAndShade = -0.149998474074526
    oSummaryRange.Font.Bold = True
  End If
  If Not oMilestoneRange Is Nothing Then
    oMilestoneRange.Font.ThemeColor = xlThemeColorAccent6
    oMilestoneRange.Font.TintAndShade = -0.249977111117893
  End If
  If Not oCompleted Is Nothing Then
    oCompleted.Font.Italic = True
    oCompleted.Font.ColorIndex = 16
  End If
  If blnValidation Then
    'date validation range
    With oDateValidationRange.Validation
      .Delete
      oWorksheet.Application.WindowState = xlNormal
      .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:=FormatDateTime(ActiveProject.ProjectStart, vbShortDate)
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Date Only"
      .ErrorTitle = "Date Only"
      .InputMessage = "Please enter a date in format mm/dd/yyyy."
      .ErrorMessage = "Please enter a date in format mm/dd/yyyy."
      .ShowInput = True
      .ShowError = True
    End With
    'number validation range
    With oNumberValidationRange.Validation
      .Delete
      .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Number Only"
      .ErrorTitle = "Number Only"
      .InputMessage = "Please enter a number greater than or equal to zero. Decimals permitted."
      .ErrorMessage = "Please enter a number greater than or equal to zero. Decimals permitted."
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If Not oAssignmentRange Is Nothing Then
    With oAssignmentRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -4.99893185216834E-02
      .PatternTintAndShade = 0
    End With
  End If
  If Not oInputRange Is Nothing Then oInputRange.Style = "Input"
  If Not oTwoWeekWindowRange Is Nothing Then oTwoWeekWindowRange.Style = "Neutral"
  
  'add EVT gloassary - test comment
  'todo: drop full list of EVTs for COBRA and MPM into ListObject
  'todo: add comment for individual EVT
  If Not oEVTRange Is Nothing Then
    If cptStatusSheet_frm.cboCostTool = "COBRA" Then
      strEVTList = "A - Level of Effort,"
      strEVTList = strEVTList & "B - Milestones,"
      strEVTList = strEVTList & "C - % Complete,"
      strEVTList = strEVTList & "D - Units Complete,"
      strEVTList = strEVTList & "E - 50-50,"
      strEVTList = strEVTList & "F - 0-100,"
      strEVTList = strEVTList & "G - 100-0,"
      strEVTList = strEVTList & "H - User Defined,"
      strEVTList = strEVTList & "J - Apportioned,"
      strEVTList = strEVTList & "K - Planning Package,"
      strEVTList = strEVTList & "L - Assignment % Complete,"
      strEVTList = strEVTList & "M - Calculated Apportionment,"
      strEVTList = strEVTList & "N - Steps,"
      strEVTList = strEVTList & "O - Earned As Spent,"
      strEVTList = strEVTList & "P - % Complete Manual Entry,"
    ElseIf cptStatusSheet_frm.cboCostTool = "MPM" Then
      
    End If
    oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).Value = "Earned Value Techniques (EVT)"
    oWorksheet.Cells(lngHeaderRow, lngLastCol).Copy
    oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).PasteSpecial xlPasteFormats
    oWorksheet.Range(oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2), oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2).Offset(UBound(Split(strEVTList, ",")), 0)).Value = oWorksheet.Application.Transpose(Split(strEVTList, ","))

  End If
  
  If blnConditionalFormats Then
    'todo: conditional formats
  End If
  
exit_here:
  On Error Resume Next
  Set oComment = Nothing
  Set oEVTRange = Nothing
  Set oCompleted = Nothing
  Set oMilestoneRange = Nothing
  Set oClearRange = Nothing
  Set oSummaryRange = Nothing
  Set oNumberValidationRange = Nothing
  Set oDateValidationRange = Nothing
  Set oTwoWeekWindowRange = Nothing
  Set oInputRange = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCopyData", Err, Erl)
  Resume exit_here
End Sub

Private Sub cptGetAssignmentData(ByRef oTask As Task, ByRef oWorksheet As Worksheet, lngRow As Long, lngHeaderRow As Long, lngNameCol As Long, lngRemainingWorkCol As Long)
  'objects
  Dim oAssignment As Assignment
  'strings
  Dim strProtect As String
  Dim strDataValidation As String
  'longs
  Dim lngIndent As Long
  Dim lngItem As Long
  Dim lngLastCol As Long
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vAssignment As Variant
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngIndent = Len(cptRegEx(oWorksheet.Cells(lngRow, lngNameCol).Value, "^\s*"))
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  lngLastRow = oWorksheet.Cells(1048576, 1).End(xlUp).Row
  lngItem = 0
  For Each oAssignment In oTask.Assignments
    lngItem = lngItem + 1
    oWorksheet.Rows(lngRow + lngItem).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Font.Italic = True 'todo: limit to columns
    vAssignment = oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Value
    vAssignment(1, 1) = oTask.UniqueID
    vAssignment(1, lngNameCol) = String(lngIndent + 3, " ") & oAssignment.ResourceName
    If oAssignment.ResourceType = pjWork Then
      vAssignment(1, lngRemainingWorkCol) = oAssignment.RemainingWork / 60
      vAssignment(1, lngRemainingWorkCol + 1) = oAssignment.RemainingWork / 60
    Else
      vAssignment(1, lngRemainingWorkCol) = oAssignment.RemainingWork
      vAssignment(1, lngRemainingWorkCol + 1) = oAssignment.RemainingWork
    End If
    'add validation
    If oNumberValidationRange Is Nothing Then
      Set oNumberValidationRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
    Else
      Set oNumberValidationRange = oWorksheet.Application.Union(oNumberValidationRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
    End If
    If (Not IsDate(oTask.ActualStart) And oTask.Start <= ActiveProject.StatusDate) Or (IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish)) Then
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
      End If
    End If
    'add protection
    If oUnlockedRange Is Nothing Then
      Set oUnlockedRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
    Else
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
    End If
    oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Value = vAssignment
    If oAssignmentRange Is Nothing Then
      Set oAssignmentRange = oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol))
    Else
      Set oAssignmentRange = oWorksheet.Application.Union(oAssignmentRange, oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)))
    End If
  Next oAssignment
  'add formulae
  If oTask.Assignments.Count > 0 Then
    oWorksheet.Cells(lngRow, lngRemainingWorkCol + 1).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C" & lngRemainingWorkCol + 1 & ":R" & lngRow + lngItem & "C" & lngRemainingWorkCol + 1 & ")"
  End If

exit_here:
  On Error Resume Next
  Set oAssignment = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptGetAssignmentData", Err, Erl)
  Resume exit_here
End Sub

Sub cptAddStatusFormats(ByRef oWorksheet As Worksheet)
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptStatusSheet_bas", "cptAddStatusFormats", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptAddConditionalFormatting(oWorksheet As Worksheet)
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptStatusSheet_bas", "cptAddConditionalFormatting", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptFinalFormats(ByRef oWorksheet As Worksheet)
Dim lngHeaderRow As Long
Dim vBorder As Variant
  lngHeaderRow = 8
  oWorksheet.Columns(1).AutoFit
  With oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight), oWorksheet.Cells(lngHeaderRow, 1).End(xlDown))
    .Borders(xlDiagonalDown).LineStyle = xlNone
    .Borders(xlDiagonalUp).LineStyle = xlNone
    For Each vBorder In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
      With .Borders(vBorder)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
      End With
    Next vBorder
  End With
'  'format entry headers
'  With oEntryHeaderRange
'    .Interior.ThemeColor = xlThemeColorAccent3
'    .Interior.TintAndShade = 0.399975585192419
'    .Font.ColorIndex = xlAutomatic
'    .BorderAround xlContinuous, xlMedium
'  End With
  oWorksheet.Application.WindowState = xlNormal
  oWorksheet.Application.Calculation = xlCalculationAutomatic
  oWorksheet.Application.ScreenUpdating = True
  oWorksheet.Application.ActiveWindow.DisplayGridLines = False
  oWorksheet.Application.ActiveWindow.SplitRow = 8
  oWorksheet.Application.ActiveWindow.SplitColumn = 0
  oWorksheet.Application.ActiveWindow.FreezePanes = True
  oWorksheet.Application.WindowState = xlMinimized
  Set oEntryHeaderRange = Nothing
End Sub
