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
Public oEVTs As Scripting.Dictionary

Sub cptShowStatusSheet_frm()
'populate all outline codes, text, and number fields
'populate UID,[user selections],Task Name,Duration,Forecast Start,Forecast Finish,Total Slack,[EVT],EV%,New EV%,BLW,Remaining Work,Revised ETC,BLS,BLF,Reason/Impact/Action
'add pick list for EV% or default to Physical % Complete
'objects
Dim rstFields As ADODB.Recordset 'Object
Dim rstEVT As ADODB.Recordset 'Object
Dim rstEVP As ADODB.Recordset 'Object
'longs
Dim lngField As Long, lngItem As Long, lngSelectedItems As Long
'integers
Dim intField As Integer
'strings
Dim strFileNamingConvention As String
Dim strDir As String
Dim strAllItems As String
Dim strAppendStatusDate As String
Dim strQuickPart As String
Dim strCC As String
Dim strSubject As String
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
    .chkAllItems = False
    .txtDir = ActiveProject.Path & "\Status Requests\" & IIf(.chkAppendStatusDate, "[yyyy-mm-dd]\", "")
    .txtFileName = "StatusRequest_[yyyy-mm-dd]"
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
            .txtFileName = "StatusRequest_[item]_[yyyy-mm-dd]"
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            If Err.Number > 0 Then
              MsgBox "Unable to set 'For Each' Field to '" & rstFields(1) & "' - contact cpt@ClearPlanConsulting.com if you need assistance.", vbExclamation + vbOKOnly, "Cannot assign For Each"
              Err.Clear
            End If
          End If
        End If
      End If
    End If
    strDir = cptGetSetting("StatusSheet", "txtDir")
    If strDir <> "" Then .txtDir = strDir
    strFileNamingConvention = cptGetSetting("StatusSheet", "txtFileName")
    If strFileNamingConvention <> "" Then .txtFileName = strFileNamingConvention
    
    strEmail = cptGetSetting("StatusSheet", "chkEmail")
    If strEmail <> "" Then .chkSendEmails = CBool(strEmail)
    If .chkSendEmails Then
      strSubject = cptGetSetting("StatusSheet", "txtSubject")
      If strSubject <> "" Then
        .txtSubject.Value = strSubject
      Else
        .txtSubject = "Status Request WE [yyyy-mm-dd]"
      End If
      strCC = cptGetSetting("StatusSheet", "txtCC")
      If strCC <> "" Then .txtCC.Value = strCC
      strQuickPart = cptGetSetting("StatusSheet", "cboQuickPart")
      If strQuickPart <> "" Then
        For lngItem = 0 To .cboQuickParts.ListCount - 1
          If .cboQuickParts.List(lngItem, 0) = strQuickPart Then
            .cboQuickParts.Value = strQuickPart
            Exit For
          End If
        Next lngItem
        If IsNull(.cboQuickParts.Value) Then
          MsgBox "A QuickPart named '" & strQuickPart & "' was not found.", vbExclamation + vbOKOnly, "Missing Quick Part?"
        End If
      End If
    End If
    strConditionalFormats = cptGetSetting("StatusSheet", "chkConditionalFormatting")
    If strConditionalFormats <> "" Then .chkAddConditionalFormats = CBool(strConditionalFormats)
    strDataValidation = cptGetSetting("StatusSheet", "chkDataValidation")
    If strDataValidation <> "" Then .chkValidation = CBool(strDataValidation)
    strLocked = cptGetSetting("StatusSheet", "chkLocked")
    If strLocked <> "" Then .chkLocked = CBool(strLocked)
    strAllItems = cptGetSetting("StatusSheet", "chkAllItems")
    If strAllItems <> "" Then .chkAllItems = CBool(strAllItems)
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
  If ActiveWindow.TopPane.View.Name <> "Gantt Chart" And ActiveWindow.TopPane.View.Name <> "Task Usage" Then
    If MsgBox("Current view must be changed for successful export.", vbInformation + vbOKCancel, "Incompatible View") = vbOK Then
      ActiveWindow.TopPane.Activate
      If strStartingGroup <> "No Group" Then
        ViewApply "Task Usage" 'required to prevent inaccurate UIDs getting into the sheet
      Else
        ViewApply "Gantt Chart"
      End If
    Else
      GoTo exit_here
    End If
  End If

  'set the status date / hide complete
  If ActiveProject.StatusDate = "NA" Then
    cptStatusSheet_frm.txtStatusDate.Value = FormatDateTime(DateAdd("d", 6 - Weekday(Now), Now), vbShortDate)
  Else
    cptStatusSheet_frm.txtStatusDate.Value = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  End If
  dtStatus = CDate(cptStatusSheet_frm.txtStatusDate.Value)
  'default to one week prior to status date
  cptStatusSheet_frm.txtHideCompleteBefore.Value = DateAdd("d", -7, dtStatus)

  strAppendStatusDate = cptGetSetting("StatusSheet", "chkAppendStatusDate")
  If strAppendStatusDate <> "" Then cptStatusSheet_frm.chkAppendStatusDate = CBool(strAppendStatusDate)

  'delete pre-existing search file
  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName

  'set up the view/table/filter
  Application.StatusBar = "Preparing View/Table/Filter..."
  DoEvents
  lngSelectedItems = 0
  For lngItem = 0 To cptStatusSheet_frm.lboItems.ListCount - 1
    If cptStatusSheet_frm.lboItems.Selected(lngItem) Then lngSelectedItems = lngSelectedItems + 1
  Next lngItem
  If lngSelectedItems = 0 Then FilterClear
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
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  cptSpeed False
  cptStatusSheet_frm.Show False
  DoEvents
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
  If Not cptCheckReference("Excel") Then
    MsgBox "Reference to Microsoft Excel not found.", vbCritical + vbOKOnly, "Is Excel installed?"
    GoTo exit_here
  End If

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
  blnEmail = cptStatusSheet_frm.chkSendEmails = True
  If blnEmail Then
    If Not cptCheckReference("Outlook") Then
      MsgBox "Reference to Microsoft Outlook not found.", vbCritical + vbOKOnly, "Is Outlook installed?"
      blnEmail = False
    Else
      On Error Resume Next
      Set oOutlook = GetObject(, "Outlook.Application")
      If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
      End If
    End If
  End If
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
        oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
        oWorksheet.EnableSelection = xlNoRestrictions
      End If
      
      .lblStatus = "Creating Workbook...done"
      
      'save the workbook
      strFileName = cptSaveStatusSheet(oWorkbook)
      
      'close the workbook - must close before attaching
      oWorkbook.Close True
      
      'send the workbook
      If blnEmail Then
        cptSendStatusSheet strFileName
      End If
      
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
            oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
          
          .lblStatus = "Creating Worksheet for " & strItem & "...done"

        End If
      Next lngItem
      
      'save the workbook
      strFileName = cptSaveStatusSheet(oWorkbook)
      
      'close the workbook - must save before attaching
      oWorkbook.Close True
      
      'send the workbook
      If blnEmail Then
        Call cptSendStatusSheet(strFileName)
      End If
      
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
            oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
                    
          'save the workbook
          strFileName = cptSaveStatusSheet(oWorkbook, strItem)
          .lblStatus = "Creating Workbook for " & strItem & "...done"
          
          'must close before attaching to email
          oWorkbook.Close True
          oWorkbook.Application.Wait Now + TimeValue("00:00:02")
          
          'send email
          If blnEmail Then
            .lblStatus = "Creating Email for " & strItem & "..."
            cptSendStatusSheet strFileName, strItem
            .lblStatus = "Creating Email for " & strItem & "...done"
          End If
          
        End If
      Next lngItem
    End If
    .lblStatus.Caption = Choose(.cboCreate + 1, "Workbook", "Workbook", "Workbooks") & " Complete"
  End With

  oExcel.Calculation = xlCalculationAutomatic
  oExcel.ScreenUpdating = True
  oExcel.Visible = True
  
  GoTo exit_here
  
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

exit_here:
  On Error Resume Next
'  Application.DefaultDateFormat = lngDateFormat
'  ActiveProject.SpaceBeforeTimeLabels = blnSpace
'  ActiveProject.DayLabelDisplay = lngDayLabelDisplay
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
  oExcel.Quit
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
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  lngItem = 0
  If cptStatusSheet_frm.lboExport.ListCount > 0 Then
    For lngItem = 0 To cptStatusSheet_frm.lboExport.ListCount - 1
      TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(cptStatusSheet_frm.lboExport.List(lngItem, 0)), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
    Next lngItem
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Name", Title:="Task Name / Scope", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Remaining Duration", Title:="", Width:=12, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Baseline Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Baseline Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False, ShowAddNewColumn:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Start", Title:="Forecast Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Finish", Title:="Forecast Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Actual Start", Title:="New Forecast/ Actual Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Actual Finish", Title:="New Forecast/ Actual Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  If cptStatusSheet_frm.cboEVT <> 0 Then
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=cptStatusSheet_frm.cboEVT.Value, Title:="EVT", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  End If
  If cptStatusSheet_frm.cboEVP <> 0 Then
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=cptStatusSheet_frm.cboEVP.Value, Title:="EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
    TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:=cptStatusSheet_frm.cboEVP.Value, Title:="New EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Baseline Work", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Remaining Work", Title:="Previous ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, newfieldname:="Remaining Work", Title:="Revised ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, headerautorowheightadjustment:=False, WrapText:=False
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
  Dim oUnlockedRange As Excel.Range
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
  Dim blnLOE As Boolean
  Dim blnLocked As Boolean
  Dim blnValidation As Boolean
  Dim blnConditionalFormats As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate
  blnValidation = cptStatusSheet_frm.chkValidation = True
  blnConditionalFormats = cptStatusSheet_frm.chkAddConditionalFormats = True
  blnLocked = cptStatusSheet_frm.chkLocked = True
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
  oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(0, 1).Value = "Reason / Action / Impact"
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
    'capture if task is LOE
    blnLOE = False
    With cptStatusSheet_frm
      If Not IsNull(.cboCostTool.Value) Then
        If .cboCostTool = "MPM" Then
          If oTask.GetField(FieldNameToFieldConstant(.cboEVT.Value)) = "6" Then blnLOE = True
        ElseIf cptStatusSheet_frm.cboCostTool = "COBRA" Then
          If oTask.GetField(FieldNameToFieldConstant(.cboEVT.Value)) = "A" Then blnLOE = True
        End If
      End If
    End With
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
    If oTask.Start < dtStatus And Not IsDate(oTask.ActualStart) Then 'should have started
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      If Not blnLOE Then Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
    End If
    If oTask.Finish <= dtStatus And Not IsDate(oTask.ActualFinish) Then 'should have finished
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
      If Not blnLOE Then Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
    End If
    If IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish) Then 'in progress
      'highlight EVP for discrete only
      If Not blnLOE Then
        If oInputRange Is Nothing Then
          Set oInputRange = oWorksheet.Cells(lngRow, lngEVPCol)
        Else
          Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngEVPCol))
        End If
      End If
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
    End If
    'two week window
    If oTask.Start > dtStatus And oTask.Start <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      If oUnlockedRange Is Nothing Then
        Set oUnlockedRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
    End If
    If oTask.Finish > dtStatus And oTask.Finish <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
      If oUnlockedRange Is Nothing Then
        Set oUnlockedRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
    End If
    'unstarted
    If Not IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish) Then 'unstarted
      If oUnlockedRange Is Nothing Then
        Set oUnlockedRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngAFCol))
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
    
    If oEVTRange Is Nothing Then 'todo: probably not needed
      Set oEVTRange = oWorksheet.Cells(lngRow, lngEVTCol)
    Else
      Set oEVTRange = oWorksheet.Application.Union(oEVTRange, oWorksheet.Cells(lngRow, lngEVTCol))
    End If
    
''    'add EVT comment - this is slow, and often fails
'    oWorksheet.Application.ScreenUpdating = True
'    Set oComment = oWorksheet.Cells(lngRow, lngEVTCol).AddComment(oEVTs.Item(oTask.GetField(FieldNameToFieldConstant(cptStatusSheet_frm.cboEVT.Value))))
'    oComment.Shape.TextFrame.Characters.Font.Bold = False
'    oComment.Shape.TextFrame.AutoSize = True
'    oWorksheet.Application.ScreenUpdating = False
    
    'todo: unlock comments
    If oUnlockedRange Is Nothing Then
      Set oUnlockedRange = oWorksheet.Cells(lngRow, lngLastCol)
    Else
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngLastCol))
    End If
    
    If oTask.Assignments.Count > 0 Then
      cptGetAssignmentData oTask, oWorksheet, lngRow, lngHeaderRow, lngNameCol, lngETCCol - 1
    End If
    
    'todo: capture conditional formatting range(s)
    
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
  If blnValidation And Not oDateValidationRange Is Nothing Then
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
  If Not oInputRange Is Nothing Then
    oInputRange.Style = "Input"
    oInputRange.Locked = False
  End If
  If blnLocked And Not oUnlockedRange Is Nothing Then oUnlockedRange.Locked = False
  If Not oTwoWeekWindowRange Is Nothing Then oTwoWeekWindowRange.Style = "Neutral"
  
  'add EVT gloassary - test comment
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
      strEVTList = strEVTList & "0 - No EVM required"
      strEVTList = strEVTList & "1 - 0/100"
      strEVTList = strEVTList & "2 - 25/75"
      strEVTList = strEVTList & "3 - 40/60"
      strEVTList = strEVTList & "4 - 50/50"
      strEVTList = strEVTList & "5 - % Complete"
      strEVTList = strEVTList & "6 - LOE"
      strEVTList = strEVTList & "7 - Earned Standards"
      strEVTList = strEVTList & "8 - Milestone Weights"
      strEVTList = strEVTList & "9 - BCWP Entry"
      strEVTList = strEVTList & "A - Apportioned"
      strEVTList = strEVTList & "P - Milestone Weights with % Complete"
      strEVTList = strEVTList & "K - Key Event"
    End If
    oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).Value = "Earned Value Techniques (EVT)"
    oWorksheet.Cells(lngHeaderRow, lngLastCol).Copy
    oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).PasteSpecial xlPasteFormats
    oWorksheet.Range(oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2), oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2).Offset(UBound(Split(strEVTList, ",")), 0)).Value = oWorksheet.Application.Transpose(Split(strEVTList, ","))
    oWorksheet.Columns(lngLastCol + 2).AutoFit
  End If
  
  If blnConditionalFormats Then
    'todo: conditional formats
  End If
  
exit_here:
  On Error Resume Next
  Set oComment = Nothing
  Set oUnlockedRange = Nothing
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
    vAssignment(1, 1) = oAssignment.UniqueID 'import assumes this is oAssignment.UniqueID
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

Sub cptFinalFormats(ByRef oWorksheet As Worksheet)
Dim lngHeaderRow As Long
Dim vBorder As Variant
  lngHeaderRow = 8
  oWorksheet.Cells(lngHeaderRow, 1).AutoFilter
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
'  'todo: format entry headers
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

Sub cptListQuickParts(Optional blnRefreshOutlook As Boolean = False)
'objects
Dim olApp As Outlook.Application
Dim oMailItem As MailItem
Dim objDoc As Word.Document
Dim objWord As Word.Application
Dim objSel As Word.Selection
Dim objETemp As Word.Template
Dim oBuildingBlockEntries As BuildingBlockEntries
Dim oBuildingBlock As BuildingBlock
'longs
Dim lngItem As Long
'strings
Dim strSQL As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If blnRefreshOutlook Then
    'refresh QuickParts in Outlook
    cptStatusSheet_frm.cboQuickParts.Clear
    'get outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
      Set olApp = CreateObject("Outlook.Application")
    End If
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    'create MailItem, insert quickparts, update links, dates
    Set oMailItem = olApp.CreateItem(olMailItem)
    'keep mailitem hidden
    If oMailItem.BodyFormat <> olFormatHTML Then oMailItem.BodyFormat = olFormatHTML
    Set objDoc = oMailItem.GetInspector.WordEditor
    Set objWord = objDoc.Application
    Set objSel = objDoc.Windows(1).Selection
    Set objETemp = objWord.Templates(1)
    Set oBuildingBlockEntries = objETemp.BuildingBlockEntries
    'loop through them
    For lngItem = 1 To oBuildingBlockEntries.Count
      Set oBuildingBlock = oBuildingBlockEntries(lngItem)
      If oBuildingBlock.Type.Name = "Quick Parts" Then
        cptStatusSheet_frm.cboQuickParts.AddItem oBuildingBlock.Name
      End If
    Next
    oMailItem.Close olDiscard
  End If
    
exit_here:
  On Error Resume Next
  Set olApp = Nothing
  Set oMailItem = Nothing
  Set objDoc = Nothing
  Set objWord = Nothing
  Set objSel = Nothing
  Set objETemp = Nothing
  Set oBuildingBlockEntries = Nothing
  Set oBuildingBlock = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptListQuickParts", Err)
  Resume exit_here
End Sub

Sub cptAddConditionalFormatting(ByRef oWorksheet As Excel.Worksheet)
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
  Call cptHandleErr("cptStatusSheet_bas", "cptAddConditionalFormatting", Err, Erl)
  Resume exit_here
End Sub


Function cptSaveStatusSheet(ByRef oWorkbook As Excel.Workbook, Optional strItem As String) As String
  'objects
  'strings
  Dim strMsg As String
  Dim strFileName As String
  Dim strDir As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate

  With cptStatusSheet_frm
    strDir = .lblDirSample.Caption
    'create the status date directory
    If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
    strFileName = .txtFileName.Value & ".xlsx"
    strFileName = Replace(strFileName, "[yyyy-mm-dd]", Format(dtStatus, "yyyy-mm-dd"))
    'todo: strFileName = cptRemoveIllegalCharacters(ActiveProject.Name)
    If Len(strItem) > 0 Then
      strFileName = Replace(strFileName, "[item]", strItem)
    End If
    On Error Resume Next
    If Dir(strDir & strFileName) <> vbNullString Then
      Kill strDir & strFileName
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
      DoEvents
    End If
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
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
    End If
  End With

  cptSaveStatusSheet = strDir & strFileName

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSaveStatusSheet", Err, Erl)
  Resume exit_here
End Function

Sub cptSendStatusSheet(strFullName As String, Optional strItem As String)
  'objects
  Dim oBuildingBlock As Word.BuildingBlock
  Dim oOutlook As Outlook.Application
  Dim oMailItem As Outlook.MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oSelection As Word.Selection
  Dim oEmailTemplate As Word.Template
  'strings
  Dim strSubject As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oOutlook = GetObject(, "Outlook.Application")
  If oOutlook Is Nothing Then
    Set oOutlook = CreateObject("Outlook.Application")
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set oMailItem = oOutlook.CreateItem(0) '0 = oloMailItem
  oMailItem.Display False
  oMailItem.Attachments.Add strFullName
  With cptStatusSheet_frm
    strSubject = .txtSubject
    strSubject = Replace(strSubject, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
    strSubject = Replace(strSubject, "[item]", strItem)
    oMailItem.Subject = strSubject
    oMailItem.CC = .txtCC
  
    If oMailItem.BodyFormat <> olFormatHTML Then oMailItem.BodyFormat = olFormatHTML
    If Not IsNull(.cboQuickParts.Value) Then
      Set oDocument = oMailItem.GetInspector.WordEditor
      Set oWord = oDocument.Application
      Set oSelection = oDocument.Windows(1).Selection
      Set oEmailTemplate = oWord.Templates(1)
      On Error Resume Next
      Set oBuildingBlock = oEmailTemplate.BuildingBlockEntries(.cboQuickParts)
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If oBuildingBlock Is Nothing Then
        MsgBox "Quick Part '" & .cboQuickParts & "' not found!", vbExclamation + vbOKOnly, "Missing Quick Part"
      Else
        oBuildingBlock.Insert oSelection.Range, True
      End If
      oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, "[STATUS_DATE]", Format(ActiveProject.StatusDate, "mm/dd/yyyy"))
    End If
  End With
exit_here:
  On Error Resume Next
  Set oBuildingBlock = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oSelection = Nothing
  Set oEmailTemplate = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSendStatusSheet", Err, Erl)
  Resume exit_here

End Sub

Sub cptSaveStatusSheetSettings()
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  With cptStatusSheet_frm
    'save settings
    cptSaveSetting "StatusSheet", "cboEVP", .cboEVP.Value
    cptSaveSetting "StatusSheet", "cboCostTool", .cboCostTool.Value
    cptSaveSetting "StatusSheet", "cboEVT", .cboEVT.Value
    cptSaveSetting "StatusSheet", "chkHide", IIf(.chkHide, 1, 0)
    cptSaveSetting "StatusSheet", "cboCreate", .cboCreate
    cptSaveSetting "StatusSheet", "txtDir", .txtDir
    cptSaveSetting "StatusSheet", "chkAppendStatusDate", IIf(.chkAppendStatusDate, 1, 0)
    If .cboEach.Value <> 0 Then
      cptSaveSetting "StatusSheet", "cboEach", .cboEach.Value
    Else
      cptSaveSetting "StatusSheet", "cboEach", "" 'todo: handle '<none>'
    End If
    cptSaveSetting "StatusSheet", "txtFileName", .txtFileName
    cptSaveSetting "StatusSheet", "chkAllItems", IIf(.chkAllItems, 1, 0)
    cptSaveSetting "StatusSheet", "chkDataValidation", IIf(.chkValidation, 1, 0)
    cptSaveSetting "StatusSheet", "chkLocked", IIf(.chkLocked, 1, 0)
    cptSaveSetting "StatusSheet", "chkConditionalFormatting", IIf(.chkAddConditionalFormats, 1, 0)
    cptSaveSetting "StatusSheet", "chkEmail", IIf(.chkSendEmails, 1, 0)
    If .chkSendEmails Then
      cptSaveSetting "StatusSheet", "txtSubject", .txtSubject
      cptSaveSetting "StatusSheet", "txtCC", .txtCC
      If Not IsNull(.cboQuickParts.Value) Then
        cptSaveSetting "StatusSheet", "cboQuickPart", .cboQuickParts.Value
      End If
    End If
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSaveStatusSheetSettings", Err, Erl)
  Resume exit_here
  
End Sub
