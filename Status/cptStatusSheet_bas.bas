Attribute VB_Name = "cptStatusSheet_bas"
'<cpt_version>v1.6.0</cpt_version>
Option Explicit
#If Win64 And VBA7 Then '<issue53>
  Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr '<issue53>
#Else '<issue53>
  Declare Function GetTickCount Lib "kernel32" () As Long
#End If '<issue53>
Private Const adVarChar As Long = 200
Private strStartingViewTopPane As String
Private strStartingViewBottomPane As String
Private strStartingTable As String
Private strStartingFilter As String
Private strStartingGroup As String
Private oAssignmentRange As Excel.Range
Private oNumberValidationRange As Excel.Range
Private oETCValidationRange As Excel.Range
Private oInputRange As Excel.Range
Private oUnlockedRange As Excel.Range
Private oEntryHeaderRange As Excel.Range
Public oEVTs As Scripting.Dictionary
Private Const lngForeColorValid As Long = -2147483630
Private Const lngBorderColorValid As Long = 8421504 '-2147483642

Sub cptShowStatusSheet_frm()
  'populate all outline codes, text, and number fields
  'populate UID,[user selections],Task Name,Duration,Forecast Start,Forecast Finish,Total Slack,[EVT],EV%,New EV%,BLW,Remaining Work,New ETC,BLS,BLF,Reason/Impact/Action
  'add pick list for EV% or default to Physical % Complete
  'objects
  Dim oRecordset As ADODB.Recordset 'Object
  Dim oShell As Object
  Dim oTasks As MSProject.Tasks
  Dim rstFields As ADODB.Recordset 'Object
  Dim rstEVT As ADODB.Recordset 'Object
  Dim rstEVP As ADODB.Recordset 'Object
  'longs
  Dim lngField As Long
  Dim lngItem As Long
  Dim lngSelectedItems As Long
  'integers
  Dim intField As Integer
  'strings
  Dim strNewCustomFieldName As String
  Dim strLOE As String
  Dim strIgnoreLOE As String
  Dim strLookahead As String
  Dim strLookaheadDays As String
  Dim strAssignments As String
  Dim strKeepOpen As String
  Dim strExportNotes As String
  Dim strAllowAssignmentNotes As String
  Dim strNotesColTitle As String
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
  Dim strFieldName As String
  Dim strFileName As String
  'dates
  Dim dtStatus As Date
  'variants
  Dim vFieldType As Variant

  'confirm existence of tasks to export
  On Error Resume Next
  Set oTasks = ActiveProject.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "This Project has no Tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  ElseIf oTasks.Count = 0 Then
    MsgBox "This Project has no Tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "No Status Date"
    If Not Application.ChangeStatusDate Then
      MsgBox "No Status Date. Exiting.", vbCritical + vbOKOnly, "No Status Date"
      GoTo exit_here
    End If
  End If
    
  'requires metrics settings
  If Not cptValidMap("EVP,EVT,LOE", blnConfirmationRequired:=True) Then
    MsgBox "No settings saved; cannot proceed.", vbExclamation + vbOKOnly, "Settings Required"
    GoTo exit_here
  End If
  
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
    .Caption = "Create Status Sheets (" & cptGetVersion("cptStatusSheet_frm") & ")"
    .lboFields.Clear
    .lboExport.Clear
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
    If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'it is a server project: default to Desktop
      Set oShell = CreateObject("WScript.Shell")
      .txtDir = oShell.SpecialFolders("Desktop") & "\Status Requests\" & IIf(.chkAppendStatusDate, "[yyyy-mm-dd]\", "")
    Else  'not a server project: use ActiveProject.Path
      .txtDir = ActiveProject.Path & "\Status Requests\" & IIf(.chkAppendStatusDate, "[yyyy-mm-dd]\", "")
    End If
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
  
  'cycle through and add all custom fields
  For Each vFieldType In Array("Text|30", "Outline Code|10", "Number|20") 'todo: start, finish, date, flag?
    Dim strFieldType As String
    Dim lngFieldCount As Long
    strFieldType = Split(vFieldType, "|")(0)
    lngFieldCount = Split(vFieldType, "|")(1)
    For intField = 1 To lngFieldCount
      lngField = FieldNameToFieldConstant(strFieldType & intField, pjTask)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        If strFieldType = "Number" Then
          rstFields.AddNew Array(0, 1, 2), Array(lngField, strFieldName, "Number")
        Else
          rstFields.AddNew Array(0, 1, 2), Array(lngField, strFieldName, "Text")
        End If
      End If
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
        If rstFields(1) = "Resources" Then
          .lboFields.List(.lboFields.ListCount - 1, 2) = FieldConstantToFieldName(rstFields(0))
        ElseIf FieldNameToFieldConstant(rstFields(1), pjTask) >= 188776000 Then
          .lboFields.List(.lboFields.ListCount - 1, 2) = "Enterprise"
        Else
          .lboFields.List(.lboFields.ListCount - 1, 2) = FieldConstantToFieldName(rstFields(0))
        End If
skip_fields:
        'add to Each
        If rstFields(1) <> "Physical % Complete" Then .cboEach.AddItem rstFields(1)
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
    strCreate = cptGetSetting("StatusSheet", "cboCreate")
    If strCreate <> "" Then .cboCreate.Value = CLng(strCreate)
    strHide = cptGetSetting("StatusSheet", "chkHide")
    If strHide <> "" Then
      .chkHide = CBool(strHide)
    Else
      .chkHide = False
    End If
    .txtHideCompleteBefore.Enabled = .chkHide
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
            If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            If Err.Number > 0 Then
              MsgBox "Unable to set 'For Each' Field to '" & rstFields(1) & "' - contact cpt@ClearPlanConsulting.com if you need assistance.", vbExclamation + vbOKOnly, "Cannot assign For Each"
              Err.Clear
            End If
          End If
        End If
      End If
      ActiveWindow.TopPane.Activate
      FilterClear
      strAllItems = cptGetSetting("StatusSheet", "chkAllItems")
      If strAllItems <> "" Then
        .chkAllItems = CBool(strAllItems)
      Else
        .chkAllItems = False
      End If
    Else
      ActiveWindow.TopPane.Activate
      FilterClear
    End If
    strDir = cptGetSetting("StatusSheet", "txtDir")
    If strDir <> "" Then .txtDir = strDir
    strFileNamingConvention = cptGetSetting("StatusSheet", "txtFileName")
    If strFileNamingConvention <> "" Then .txtFileName = strFileNamingConvention
    
    strEmail = cptGetSetting("StatusSheet", "chkEmail")
    If strEmail <> "" Then
      .chkSendEmails.Value = CBool(strEmail)  'this refreshes the quickparts list
      If .chkSendEmails Then
        .chkKeepOpen.Value = False
        .chkKeepOpen.Enabled = False
      End If
    Else
      .chkSendEmails.Value = False
      .chkKeepOpen.Enabled = True
    End If
    If .chkSendEmails Then
      strSubject = cptGetSetting("StatusSheet", "txtSubject")
      If strSubject <> "" Then
        .txtSubject.Value = strSubject
      Else
        .txtSubject = "Status Request WE [yyyy-mm-dd]"
      End If
      strCC = cptGetSetting("StatusSheet", "txtCC")
      If strCC <> "" Then .txtCC.Value = strCC
      'cboQuickParts updated when .chkSendEmails = true
    End If
    
    strConditionalFormats = cptGetSetting("StatusSheet", "chkConditionalFormatting")
    If strConditionalFormats <> "" Then
      .chkAddConditionalFormats.Value = CBool(strConditionalFormats)
    Else
      .chkAddConditionalFormats.Value = False
    End If
    
    strDataValidation = cptGetSetting("StatusSheet", "chkDataValidation")
    If strDataValidation <> "" Then
      .chkValidation = CBool(strDataValidation)
    Else
      .chkValidation = True
    End If
    
    strLocked = cptGetSetting("StatusSheet", "chkLocked")
    If strLocked <> "" Then
      .chkLocked.Value = CBool(strLocked)
    Else
      .chkLocked.Value = True
    End If
    
    strNotesColTitle = cptGetSetting("StatusSheet", "txtNotesColTitle")
    If Len(strNotesColTitle) > 0 Then
      .txtNotesColTitle.Value = strNotesColTitle
    Else
      .txtNotesColTitle = "Reason / Action / Impact"
    End If
    
    strExportNotes = cptGetSetting("StatusSheet", "chkExportNotes")
    If strExportNotes <> "" Then
      .chkExportNotes.Value = CBool(strExportNotes)
    Else
      .chkExportNotes.Value = False
    End If
        
    strKeepOpen = cptGetSetting("StatusSheet", "chkKeepOpen")
    If strKeepOpen <> "" Then
      .chkKeepOpen.Value = CBool(strKeepOpen)
      If .chkKeepOpen Then
        .chkSendEmails.Value = False
        .chkSendEmails.Enabled = False
      Else
        .chkSendEmails.Enabled = True
      End If
    Else
      .chkKeepOpen.Value = False
    End If
    
    strAssignments = cptGetSetting("StatusSheet", "chkAssignments")
    If strAssignments <> "" Then
      .chkAssignments.Value = CBool(strAssignments)
    Else
      .chkAssignments.Value = True 'default
    End If
    
    If .chkAssignments Then
      .chkAllowAssignmentNotes.Enabled = True
      strAllowAssignmentNotes = cptGetSetting("StatusSheet", "chkAllowAssignmentNotes")
      If strAllowAssignmentNotes <> "" Then
        .chkAllowAssignmentNotes.Value = CBool(strAllowAssignmentNotes)
      Else
        .chkAllowAssignmentNotes.Value = False 'default
      End If
    Else
      .chkAllowAssignmentNotes.Value = False
      .chkAllowAssignmentNotes.Enabled = False
    End If
    
    .txtLookaheadDays.Enabled = False
    .txtLookaheadDate.Enabled = False
    .lblLookaheadWeekday.Visible = False
    strLookahead = cptGetSetting("StatusSheet", "chkLookahead")
    If Len(strLookahead) > 0 Then
      .chkLookahead = CBool(strLookahead)
    Else
      .chkLookahead = False 'default
    End If
    
    If .chkLookahead Then
      .txtLookaheadDays.Enabled = True
      strLookaheadDays = cptGetSetting("StatusSheet", "txtLookaheadDays")
      If Len(strLookaheadDays) > 0 Then
        .txtLookaheadDays = CLng(strLookaheadDays)
        .lblLookaheadWeekday.Visible = True
      End If
      .txtLookaheadDate.Enabled = True
    End If
    
    .chkIgnoreLOE.Enabled = False
    strEVT = Split(cptGetSetting("Integration", "EVT"), "|")(1)
    strLOE = cptGetSetting("Integration", "LOE")
    If Len(strEVT) > 0 And Len(strLOE) > 0 Then
      .chkIgnoreLOE.Enabled = True
      .chkIgnoreLOE.ControlTipText = "Limit to tasks where " & strEVT & " <> " & strLOE
      strIgnoreLOE = cptGetSetting("StatusSheet", "chkIgnoreLOE")
      If Len(strIgnoreLOE) > 0 Then
        .chkIgnoreLOE = CBool(strIgnoreLOE)
      Else
        .chkIgnoreLOE = False
      End If
    End If
    
  End With

  'add saved export fields if they exist
  strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strFileName
      'todo: add program acronym field and filter for it?
      If .RecordCount > 0 Then
        .MoveFirst
        lngItem = 0
        Do While Not .EOF
          cptStatusSheet_frm.lboExport.AddItem
          cptStatusSheet_frm.lboExport.List(lngItem, 0) = .Fields(0) 'Field Constant
          cptStatusSheet_frm.lboExport.List(lngItem, 1) = .Fields(1) 'Custom Field Name
          cptStatusSheet_frm.lboExport.List(lngItem, 2) = .Fields(2) 'Local Field Name
          'todo: what was this for? no FieldConstantToFieldName(constant) returns "Custom"?
          'todo: was this for filtering out enterprise fields since CFGN = FCFN?
          'If cptRegEx(FieldConstantToFieldName(.Fields(0)), "[0-9]{1,}$") = "" Then GoTo next_item
          'If InStr("Custom", FieldConstantToFieldName(FieldNameToFieldConstant(.Fields(2)))) = 0 Then GoTo next_item
          If CustomFieldGetName(.Fields(0)) <> CStr(.Fields(1)) Then
            If Len(CustomFieldGetName(.Fields(0))) > 0 Then
              strNewCustomFieldName = CustomFieldGetName(.Fields(0))
            Else
              strNewCustomFieldName = "<unnamed>"
            End If
            'prompt user to accept changed name or remove from list
            If MsgBox("Saved field '" & .Fields(1) & "' has been renamed to '" & strNewCustomFieldName & "'." & vbCrLf & vbCrLf & "Click Yes to accept the name change." & vbCrLf & "Click No to remove from export list.", vbExclamation + vbYesNo, "Confirm Export Field") = vbYes Then
              'update export list
              cptStatusSheet_frm.lboExport.List(lngItem, 1) = CustomFieldGetName(.Fields(0))
              'update the adtg
              .Fields(1) = CustomFieldGetName(.Fields(0))
              .Update
            Else
              'remove from export list
              cptStatusSheet_frm.lboExport.RemoveItem (lngItem)
              'remove from adtg
              .Delete adAffectCurrent
              .Update
              lngItem = lngItem - 1
            End If
          End If
next_item:
          lngItem = lngItem + 1
          .MoveNext
        Loop
      End If
      .Filter = 0
      'overwrite in case of field name changes
      .Save strFileName, adPersistADTG
      .Close
    End With
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
  ActiveWindow.TopPane.Activate
  strStartingViewTopPane = ActiveWindow.TopPane.View.Name
  If Not ActiveWindow.BottomPane Is Nothing Then
    strStartingViewBottomPane = ActiveWindow.BottomPane.View.Name
    ActiveWindow.BottomPane.Activate
    Application.PaneClose
  Else
    strStartingViewBottomPane = "None"
  End If
  strStartingTable = ActiveProject.CurrentTable
  strStartingFilter = ActiveProject.CurrentFilter
  If ActiveProject.CurrentGroup = "Custom Group" Then
    MsgBox "An ad hoc Autofilter Group cannot be used." & vbCrLf & vbCrLf & "Please save the group and name it, or select another saved Group, before you proceed.", vbInformation + vbOKOnly, "Invalid Group"
    GoTo exit_here
  Else
    strStartingGroup = ActiveProject.CurrentGroup
  End If
  
  cptSpeed True
  ActiveWindow.TopPane.Activate
  If ActiveWindow.TopPane.View.Type <> pjTaskItem Then
    ViewApply "Gantt Chart"
    If ActiveProject.CurrentGroup <> "No Group" Then GroupApply "No Group" 'if not a task view, then group is irrelevant
  Else
    If strStartingGroup = "No Group" Then
      'no fake Group Summary UIDs will be used
    Else
      If CBool(cptGetSetting("StatusSheet", "chkAssignments")) Then
        If Not strStartingViewTopPane = "Task Usage" Then ViewApply "Task Usage"
      Else
        If Not strStartingViewTopPane = "Gantt Chart" Then ViewApply "Gantt Chart"
      End If
      'task usage view avoids fake Group Summary UIDs
      If ActiveProject.CurrentGroup <> strStartingGroup Then GroupApply strStartingGroup
    End If
  End If
  DoEvents
  
  OptionsViewEx DisplaySummaryTasks:=True, DisplayNameIndent:=True
  If strStartingGroup = "No Group" Then
    Sort "ID", , , , , , False, True 'OutlineShowAllTasks won't work without this
  Else
    If ActiveProject.CurrentGroup <> strStartingGroup Then
      On Error Resume Next
      GroupApply strStartingGroup
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    End If
  End If
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptRefreshStatusTable True  'this only runs when form is visible
  FilterClear 'added 9/28/2021
  FilterApply "cptStatusSheet Filter"
  If Len(strCreate) > 0 And Len(strEach) > 0 Then
    On Error Resume Next
    SetAutoFilter strEach, pjAutoFilterClear
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    DoEvents
  End If
  If strStartingGroup <> "No Group" Then
    On Error Resume Next
    GroupApply strStartingGroup
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  End If
  DoEvents
  Application.StatusBar = "Ready..."
  DoEvents
  cptStatusSheet_frm.txtStatusDate.SetFocus
  cptSpeed True
  cptStatusSheet_frm.Show 'Modal = True! Keep!
  
  'after user closes form, then:
  Application.StatusBar = "Restoring your view/table/filter/group..."
  DoEvents
  cptSpeed True
  ActiveWindow.TopPane.Activate
  ViewApply strStartingViewTopPane
  If strStartingViewBottomPane <> "None" Then
    If strStartingViewBottomPane = "Timeline" Then
      ViewApplyEx Name:="Timeline", applyto:=1
    Else
      PaneCreate
      ViewApplyEx strStartingViewBottomPane, applyto:=1
      ActiveWindow.TopPane.Activate
    End If
  End If
  If ActiveProject.CurrentTable <> strStartingTable Then TableApply strStartingTable
  If ActiveProject.CurrentFilter <> strStartingFilter Then FilterApply strStartingFilter
  If ActiveProject.CurrentGroup <> strStartingGroup Then GroupApply strStartingGroup
  
exit_here:
  On Error Resume Next
  Set cptStatusSheet_frm = Nothing
  Set oRecordset = Nothing
  Set oShell = Nothing
  Application.StatusBar = ""
  cptSpeed False
  Set oTasks = Nothing
  If rstFields.State Then rstFields.Close
  Set rstFields = Nothing
  Exit Sub

err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cptShowStatusSheet_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptCreateStatusSheet()
  'objects
  Dim oTasks As MSProject.Tasks, oTask As MSProject.Task, oAssignment As MSProject.Assignment
  'early binding:
  Dim oExcel As Excel.Application, oWorkbook As Excel.Workbook, oWorksheet As Excel.Worksheet, rng As Excel.Range
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
  Dim blnKeepOpen As Boolean
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
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure project has tasks
  If oTasks Is Nothing Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "Create Status Sheet"
    GoTo exit_here
  End If
    
  cptStatusSheet_frm.lblStatus.Caption = " Analyzing project..."
  Application.StatusBar = "Analyzing project..."
  DoEvents
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
  blnKeepOpen = cptStatusSheet_frm.chkKeepOpen
  'get task count
  If blnPerformanceTest Then t = GetTickCount
  SelectAll
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "There are no incomplete tasks in this schedule.", vbExclamation + vbOKOnly, "No Tasks Found"
    GoTo exit_here
  End If
  lngTaskCount = oTasks.Count
  If blnPerformanceTest Then Debug.Print "<=====PERFORMANCE TEST " & Now() & "=====>"

  cptStatusSheet_frm.lblStatus.Caption = " Setting up Workbook..."
  Application.StatusBar = "Setting up Workbook..."
  DoEvents
  'set up an excel Workbook
  If blnPerformanceTest Then t = GetTickCount
  Set oExcel = CreateObject("Excel.Application") 'do not use GetObject
  'oExcel.Visible = False
  oExcel.WindowState = xlMinimized
  '/=== debug ==\
  If Not cptErrorTrapping Then oExcel.Visible = True
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
      
      SelectAll
      On Error Resume Next
      Set oTasks = Nothing
      Set oTasks = ActiveSelection.Tasks
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oTasks Is Nothing Then
        .lblStatus.Caption = "No incomplete tasks ...skipped"
        Application.StatusBar = .lblStatus.Caption
        GoTo exit_here
      End If
      
      Set oWorkbook = oExcel.Workbooks.Add
      oExcel.Calculation = xlCalculationManual
      oExcel.ScreenUpdating = False
      Set oWorksheet = oWorkbook.Sheets(1)
      oWorksheet.Name = "Status Sheet"
      
      'copy data
      If blnPerformanceTest Then t = GetTickCount
      .lblStatus.Caption = "Creating Workbook..."
      Application.StatusBar = .lblStatus.Caption
      DoEvents
      cptCopyData oWorksheet, lngHeaderRow
      If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
      
      'add legend
      If blnPerformanceTest Then t = GetTickCount
      cptAddLegend oWorksheet, dtStatus
      If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
      
      'final formatting
      cptFinalFormats oWorksheet
      
      Set oInputRange = Nothing
      Set oNumberValidationRange = Nothing
      Set oETCValidationRange = Nothing
      Set oUnlockedRange = Nothing
      Set oAssignmentRange = Nothing
      
      oWorksheet.Calculate
      
      If blnLocked Then 'protect the sheet
        oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
        oWorksheet.EnableSelection = xlNoRestrictions
      End If
      
      .lblStatus.Caption = "Creating Workbook...done"
      Application.StatusBar = .lblStatus.Caption
      DoEvents
      
      'save the workbook
      .lblStatus.Caption = "Saving Workbook..."
      Application.StatusBar = .lblStatus.Caption
      DoEvents
      strFileName = cptSaveStatusSheet(oWorkbook)
      
      oExcel.Calculation = xlCalculationAutomatic
      oExcel.ScreenUpdating = True
      
      'send the workbook
      If blnEmail Then
        'close the workbook - must close before attaching
        oWorkbook.Close True
        oExcel.Wait Now + TimeValue("00:00:02")
        cptSendStatusSheet strFileName
      Else
        If Not blnKeepOpen Then
          oWorkbook.Close True
          oExcel.Wait Now + TimeValue("00:00:002")
        Else
          oExcel.Visible = True
        End If
      End If
      
    ElseIf .cboCreate.Value = "1" Then  'worksheet for each
      Set oWorkbook = oExcel.Workbooks.Add
      oExcel.Calculation = xlCalculationManual
      oExcel.ScreenUpdating = False
      For lngItem = 0 To .lboItems.ListCount - 1
        If .lboItems.Selected(lngItem) Then
          strItem = .lboItems.List(lngItem, 0)
          SetAutoFilter .cboEach.Value, pjAutoFilterCustom, "equals", strItem
          SelectAll
          Set oTasks = Nothing
          On Error Resume Next
          Set oTasks = ActiveSelection.Tasks
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTasks Is Nothing Then
            .lblStatus.Caption = "No incomplete tasks for " & strItem & "...skipped"
            Application.StatusBar = .lblStatus.Caption
            GoTo next_worksheet
          End If
          
          'create worksheet
          Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
          oWorksheet.Name = strItem
          'copy data
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Creating Worksheet for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          DoEvents
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
          Set oETCValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
          
next_worksheet:
          .lblStatus.Caption = "Creating Worksheet for " & strItem & "...done"
          Application.StatusBar = .lblStatus.Caption
          DoEvents
          
        End If
      Next lngItem
      
      Set oWorksheet = Nothing
      On Error Resume Next
      Set oWorksheet = oWorkbook.Sheets("Sheet1")
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If Not oWorksheet Is Nothing Then oWorksheet.Delete
      oWorkbook.Sheets(1).Activate
      
      'save the workbook
      strFileName = cptSaveStatusSheet(oWorkbook)
      
      'turn Excel back on
      oExcel.Calculation = xlCalculationAutomatic
      oExcel.ScreenUpdating = True
      
      'send the workbook
      If blnEmail Then
        'close the workbook - must close before attaching
        oWorkbook.Close True
        oExcel.Wait Now + TimeValue("00:00:02")
        cptSendStatusSheet strFileName
      Else
        If Not blnKeepOpen Then
          oWorkbook.Close True
          oExcel.Wait Now + TimeValue("00:00:002")
        Else
          oExcel.Visible = True
        End If
      End If
      
    ElseIf .cboCreate.Value = "2" Then  'workbook for each
      For lngItem = 0 To .lboItems.ListCount - 1
        If .lboItems.Selected(lngItem) Then
          strItem = .lboItems.List(lngItem, 0)
          SetAutoFilter .cboEach.Value, pjAutoFilterCustom, "equals", strItem
          SelectAll
          On Error Resume Next
          Set oTasks = Nothing
          Set oTasks = ActiveSelection.Tasks
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTasks Is Nothing Then
            .lblStatus.Caption = "No incomplete tasks for " & strItem & "...skipped"
            Application.StatusBar = .lblStatus.Caption
            GoTo next_workbook
          End If
          
          'get excel
          If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application") 'todo: added
          Set oWorkbook = oExcel.Workbooks.Add
          oExcel.Calculation = xlCalculationManual
          oExcel.ScreenUpdating = False
          Set oWorksheet = oWorkbook.Sheets(1)
          oWorksheet.Name = "Status Request"
          
          'copy data
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Creating Workbook for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          DoEvents
          cptCopyData oWorksheet, lngHeaderRow
          If blnPerformanceTest Then Debug.Print "copy data: " & (GetTickCount - t) / 1000
          
          'add legend
          If blnPerformanceTest Then t = GetTickCount
          .lblStatus.Caption = "Building legend for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          cptAddLegend oWorksheet, dtStatus
          .lblStatus.Caption = "Building legend for " & strItem & "...done."
          Application.StatusBar = .lblStatus.Caption
          If blnPerformanceTest Then Debug.Print "set up legend: " & (GetTickCount - t) / 1000
          
          'final formatting
          .lblStatus.Caption = "Formatting " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          cptFinalFormats oWorksheet
          .lblStatus.Caption = "Formatting " & strItem & "...done."
          Application.StatusBar = .lblStatus.Caption
          
          oWorksheet.Calculate
          
          If blnLocked Then 'protect the sheet
            .lblStatus.Caption = "Protecting " & strItem & "..."
            Application.StatusBar = .lblStatus.Caption
            oWorksheet.Protect Password:="NoTouching!", DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingRows:=True, AllowFormattingColumns:=True, AllowFormattingCells:=True
            oWorksheet.EnableSelection = xlNoRestrictions
            .lblStatus.Caption = "Protecting " & strItem & "...done."
            Application.StatusBar = .lblStatus.Caption
          End If
          
          Set oInputRange = Nothing
          Set oNumberValidationRange = Nothing
          Set oETCValidationRange = Nothing
          Set oUnlockedRange = Nothing
          Set oAssignmentRange = Nothing
                    
          'save the workbook
          .lblStatus.Caption = "Saving Workbook for " & strItem & "..."
          Application.StatusBar = .lblStatus.Caption
          strFileName = cptSaveStatusSheet(oWorkbook, strItem)
          .lblStatus.Caption = "Saving Workbook for " & strItem & "...done"
          DoEvents
          
          'send email
          oExcel.Calculation = xlCalculationAutomatic
          If blnEmail Then
            .lblStatus.Caption = "Creating Email for " & strItem & "..."
            Application.StatusBar = .lblStatus.Caption
            DoEvents
            'must close before attaching to email
            oWorkbook.Close True
            oExcel.Wait Now + TimeValue("00:00:02")
            oExcel.Quit 'todo: added
            Set oExcel = Nothing 'todo: added
            cptSendStatusSheet strFileName, strItem
            .lblStatus.Caption = "Creating Email for " & strItem & "...done"
            Application.StatusBar = .lblStatus.Caption
            DoEvents
          Else
            If Not blnKeepOpen Then
              oWorkbook.Close True
              oExcel.Wait Now + TimeValue("00:00:002")
            Else
              oExcel.Visible = True
              oWorkbook.Activate
            End If
          End If 'blnEmail
        End If '.lboItems.Selected(lngItem)
                
next_workbook:
        
      Next lngItem
      
      If Not blnEmail Then
        oExcel.ScreenUpdating = True
        oExcel.Visible = True
      End If
      
    End If
    .lblStatus.Caption = Choose(.cboCreate + 1, "Workbook", "Workbook", "Workbooks") & " Complete"
    Application.StatusBar = .lblStatus.Caption
    DoEvents
  End With

exit_here:
  On Error Resume Next
  If oExcel.Workbooks.Count > 0 Then oExcel.Calculation = xlAutomatic
  oExcel.ScreenUpdating = True
  oExcel.EnableEvents = True
  Application.StatusBar = ""
  cptSpeed False
  Set oTasks = Nothing
  Set oTask = Nothing
  Set oAssignment = Nothing
  If blnEmail Then oExcel.Quit
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  'todo: only quit if we had to create
'  If Not oExcel Is Nothing Then
'    oExcel.Visible = True
'    oExcel.Quit
    Set oExcel = Nothing
'  End If
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

Sub cptRefreshStatusTable(Optional blnOverride As Boolean = False, Optional blnFilterOnly As Boolean = False)
  'objects
  'strings
  Dim strLOE As String
  Dim strEVT As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtLookahead As Date

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not cptStatusSheet_frm.Visible And blnOverride = False Then GoTo exit_here

  If Not blnOverride Then cptSpeed True
  If blnFilterOnly Then GoTo filter_only
  
  'reset the view
  Application.StatusBar = "Resetting the cptStatusSheet View..."
  Application.ActiveWindow.TopPane.Activate
  If cptStatusSheet_frm.chkAssignments Then
    ViewApply "Task Usage"
  Else
    ViewApply "Gantt Chart"
  End If
  
  'reset the group
  Application.StatusBar = "Resetting the cptStatusSheet Group..."
  If ActiveProject.CurrentGroup <> "No Group" Then
    strStartingGroup = ActiveProject.CurrentGroup
    GroupApply "No Group"
  End If
  
  'reset the table
  Application.StatusBar = "Resetting the cptStatusSheet Table..."
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  lngItem = 0
  If cptStatusSheet_frm.lboExport.ListCount > 0 Then
    For lngItem = 0 To cptStatusSheet_frm.lboExport.ListCount - 1
      If Not IsNull(cptStatusSheet_frm.lboExport.List(lngItem, 0)) Then
        TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(cptStatusSheet_frm.lboExport.List(lngItem, 0)), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
      End If
    Next lngItem
  End If
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Name", Title:="Task Name / Scope", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Duration", Title:="", Width:=12, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Start", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Finish", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False, ShowAddNewColumn:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Start", Title:="Forecast Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Finish", Title:="Forecast Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Actual Start", Title:="New Forecast/ Actual Start", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Actual Finish", Title:="New Forecast/ Actual Finish", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=Split(cptGetSetting("Integration", "EVT"), "|")(1), Title:="EVT", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=Split(cptGetSetting("Integration", "EVP"), "|")(1), Title:="EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:=Split(cptGetSetting("Integration", "EVP"), "|")(1), Title:="New EV%", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Baseline Work", Title:="", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheet Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="New ETC", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableApply Name:="cptStatusSheet Table"

filter_only:
  'reset the filter
  Application.StatusBar = "Resetting the cptStatusSheet Filter..."
  FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Actual Finish", Test:="equals", Value:="NA", ShowInMenu:=False, ShowSummaryTasks:=True
  If cptStatusSheet_frm.chkHide And IsDate(cptStatusSheet_frm.txtHideCompleteBefore) Then
    FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Actual Finish", Test:="is greater than or equal to", Value:=cptStatusSheet_frm.txtHideCompleteBefore, Operation:="Or", ShowSummaryTasks:=True
  End If
  If Edition = pjEditionProfessional Then
    FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Active", Test:="equals", Value:="Yes", ShowInMenu:=False, ShowSummaryTasks:=True, Parenthesis:=True
  End If
  With cptStatusSheet_frm
    If .chkLookahead And .txtLookaheadDate.BorderColor <> 192 Then
      dtLookahead = CDate(.txtLookaheadDate) & " 5:00 PM"
      FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:="Start", Test:="is less than or equal to", Value:=dtLookahead, Operation:="And", Parenthesis:=False
    End If
    If .chkIgnoreLOE Then
      strEVT = Split(cptGetSetting("Integration", "EVT"), "|")(1)
      strLOE = cptGetSetting("Integration", "LOE")
      FilterEdit Name:="cptStatusSheet Filter", TaskFilter:=True, FieldName:="", NewFieldName:=strEVT, Test:="does not equal", Value:=strLOE, Operation:="And", Parenthesis:=False
    End If
  End With
  FilterApply "cptStatusSheet Filter"
  
  If Len(strStartingGroup) > 0 Then
    Application.StatusBar = "Restoring the cptStatusSheet Group..."
    GroupApply strStartingGroup
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  If Not blnOverride Then cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptRefreshStatusTable", Err, Erl)
  Err.Clear
  Resume exit_here
End Sub

Private Sub cptAddLegend(ByRef oWorksheet As Excel.Worksheet, dtStatus As Date)
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  oWorksheet.Cells(1, 1).Value = "Status Date:"
  oWorksheet.Cells(1, 1).Font.Bold = True
  oWorksheet.Cells(1, 2) = FormatDateTime(dtStatus, vbShortDate)
  oWorksheet.Names.Add "STATUS_DATE", oWorksheet.[B1]
  oWorksheet.Cells(1, 2).Font.Bold = True
  oWorksheet.Cells(1, 2).Font.Size = 14
  oWorksheet.Cells(1, 2).HorizontalAlignment = xlCenter
  oWorksheet.Cells(1, 2).Style = "Note"
  oWorksheet.Cells(1, 2).Columns.AutoFit
  'current
  oWorksheet.Cells(3, 1).Style = "Input" '<issue58>
  oWorksheet.Cells(3, 2) = "Task is active or within current status window. Cell requires update."
  'within two weeks
  oWorksheet.Cells(4, 1).Style = "Neutral" '<issue58>
  oWorksheet.Cells(4, 1).BorderAround xlContinuous, xlThin, , -8421505
  oWorksheet.Cells(4, 2) = "Task is within two week look-ahead. Please review forecast dates."
  'complete
  oWorksheet.Cells(5, 1).Style = "Explanatory Text"
'  oWorksheet.Cells(5, 1).Font.TintAndShade = 0
'  oWorksheet.Cells(5, 1).Interior.PatternColorIndex = -4105
'  oWorksheet.Cells(5, 1).Interior.Color = 15921906
'  oWorksheet.Cells(5, 1).Interior.TintAndShade = 0
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

Private Sub cptCopyData(ByRef oWorksheet As Excel.Worksheet, lngHeaderRow As Long)
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oFormatRange As Object
  Dim oDict As Scripting.Dictionary
  Dim oRecordset As ADODB.Recordset
  Dim oFirstCell As Excel.Range
  Dim oETCRange As Excel.Range
  Dim oEVPRange As Excel.Range
  Dim oNFRange As Excel.Range
  Dim oNSRange As Excel.Range
  Dim oComment As Excel.Comment
  Dim oEVTRange As Excel.Range
  Dim oCompleted As Excel.Range
  Dim oMilestoneRange As Excel.Range
  Dim oClearRange As Excel.Range
  Dim oSummaryRange As Excel.Range
  Dim oDateValidationRange As Excel.Range
  Dim oTwoWeekWindowRange As Excel.Range
  Dim oTask As MSProject.Task
  'strings
  Dim strItem As String
  Dim strFormula As String
  Dim strETC As String
  Dim strEVP As String
  Dim strNF As String
  Dim strAF As String
  Dim strFF As String
  Dim strNS As String
  Dim strAS As String
  Dim strFS As String
  Dim strNotesColTitle As String
  Dim strLOE As String
  Dim strEVT As String
  Dim strEVTList As String
  'longs
  Dim lngLastRow As Long
  Dim lngFormatCondition As Long
  Dim lngFormatConditions As Long
  Dim lngEVT As Long
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
  Dim blnAssignments As Boolean
  Dim blnAlerts As Boolean
  Dim blnLOE As Boolean
  Dim blnLocked As Boolean
  Dim blnValidation As Boolean
  Dim blnConditionalFormats As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate
  blnValidation = cptStatusSheet_frm.chkValidation = True
  blnConditionalFormats = cptStatusSheet_frm.chkAddConditionalFormats = True
  blnLocked = cptStatusSheet_frm.chkLocked = True
  ActiveWindow.TopPane.Activate
try_again:
  SelectAll
  EditCopy
  DoEvents
  oWorksheet.Application.Wait 5000
  On Error Resume Next
  oWorksheet.Paste oWorksheet.Cells(lngHeaderRow, 1), False
  If Err.Number = 1004 Then 'try again
    EditCopy
    oWorksheet.Application.Wait 5000
    oWorksheet.Paste oWorksheet.Cells(lngHeaderRow, 1), False
    oWorksheet.Application.Wait 5000
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  oWorksheet.Application.Wait 5000
  oWorksheet.Cells.WrapText = False
  oWorksheet.Application.ActiveWindow.Zoom = 85
  oWorksheet.Cells.Font.Name = "Calibri"
  oWorksheet.Cells.Font.Size = 11
  oWorksheet.Rows(lngHeaderRow).Font.Bold = True
  oWorksheet.Columns.AutoFit
  'format the colums
  blnAlerts = oWorksheet.Application.DisplayAlerts
  strItem = cptStatusSheet_frm.lboItems.List(cptStatusSheet_frm.lboItems.ListIndex, 0)
  If blnAlerts Then oWorksheet.Application.DisplayAlerts = False
  For lngCol = 1 To ActiveSelection.FieldIDList.Count
    oWorksheet.Columns(lngCol).ColumnWidth = ActiveProject.TaskTables("cptStatusSheet Table").TableFields(lngCol + 1).Width + 2
    oWorksheet.Cells(lngHeaderRow, lngCol).WrapText = True
    If InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Start") > 0 Then
      oWorksheet.Columns(lngCol).Replace "NA", ""
      oWorksheet.Columns(lngCol).NumberFormat = "m/d/yyyy"
    ElseIf InStr(oWorksheet.Cells(lngHeaderRow, lngCol), "Finish") > 0 Then
      oWorksheet.Columns(lngCol).Replace "NA", ""
      oWorksheet.Columns(lngCol).NumberFormat = "m/d/yyyy"
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
  Next lngCol
  oWorksheet.Application.DisplayAlerts = blnAlerts
  
  'format the header
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  If lngLastCol > ActiveProject.TaskTables("cptStatusSheet Table").TableFields.Count + 10 Then GoTo try_again
  oWorksheet.Columns(lngLastCol + 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
  oWorksheet.Columns(lngLastCol + 1).ColumnWidth = 40
  strNotesColTitle = cptGetSetting("StatusSheet", "txtNotesColTitle")
  If Len(strNotesColTitle) > 0 Then
    oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(0, 1).Value = strNotesColTitle
  Else
    oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Offset(0, 1).Value = "Reason / Action / Impact" 'default
  End If
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
  
  'get LOE settings
  strEVT = cptGetSetting("Integration", "EVT")
  If Len(strEVT) > 0 Then
    lngEVT = CLng(Split(strEVT, "|")(0))
  End If
  strLOE = cptGetSetting("Integration", "LOE")
  
  'format the data rows
  lngNameCol = oWorksheet.Rows(lngHeaderRow).Find("Task Name / Scope", lookat:=xlWhole).Column
  lngASCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Start", lookat:=xlPart).Column
  lngAFCol = oWorksheet.Rows(lngHeaderRow).Find("Actual Finish", lookat:=xlPart).Column
  lngEVPCol = oWorksheet.Rows(lngHeaderRow).Find("New EV%", lookat:=xlWhole).Column
  lngEVTCol = oWorksheet.Rows(lngHeaderRow).Find("EVT", lookat:=xlWhole).Column
  'todo: add Milestones EVT
  lngETCCol = oWorksheet.Rows(lngHeaderRow).Find("New ETC", lookat:=xlWhole).Column
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  lngTasks = ActiveSelection.Tasks.Count
  lngTask = 0
  For Each oTask In ActiveSelection.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    'find the row of the current task
    On Error Resume Next
    lngRow = 0
    lngRow = oWorksheet.Columns(1).Find(oTask.UniqueID, lookat:=xlWhole).Row
    If Err.Number = 91 Then
      MsgBox "UID " & oTask.UniqueID & " not found on worksheet!" & vbCrLf & vbCrLf & "You may need to re-run...", vbExclamation + vbOKOnly, "ERROR"
      GoTo next_task
    End If
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    'capture if task is LOE
    blnLOE = oTask.GetField(lngEVT) = strLOE
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
'      GoTo next_task 'don't skip - need to unlock foreceast dates for milestones, too
    End If
    'format completed todo: still needed?
    If IsDate(oTask.ActualFinish) Then
      If oCompleted Is Nothing Then
        Set oCompleted = oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol))
      Else
        Set oCompleted = oWorksheet.Application.Union(oCompleted, oWorksheet.Range(oWorksheet.Cells(lngRow, 1), oWorksheet.Cells(lngRow, lngLastCol)))
      End If
      GoTo get_assignments
    End If
    If blnLOE Then
      oWorksheet.Cells(lngRow, lngEVPCol - 1) = "'-"
      oWorksheet.Cells(lngRow, lngEVPCol) = "'-"
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
    End If
    If oTask.Finish > dtStatus And oTask.Finish <= DateAdd("d", 14, dtStatus) Then
      If oTwoWeekWindowRange Is Nothing Then
        Set oTwoWeekWindowRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oTwoWeekWindowRange = oWorksheet.Application.Union(oTwoWeekWindowRange, oWorksheet.Cells(lngRow, lngAFCol))
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
      If Not blnLOE Then Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngEVPCol))
      'Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngETCCol))
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
        'allow incomplete tasks to have EVP updated
        If Not blnLOE Then
          If oNumberValidationRange Is Nothing Then
            Set oNumberValidationRange = oWorksheet.Cells(lngRow, lngEVPCol)
          Else
            Set oNumberValidationRange = oWorksheet.Application.Union(oNumberValidationRange, oWorksheet.Cells(lngRow, lngEVPCol))
          End If
        End If
      End If
    End If 'blnValidation
    
    'capture conditional formatting ranges
    blnConditionalFormats = cptStatusSheet_frm.chkAddConditionalFormats
    If Not blnLOE And blnConditionalFormats Then
      If oNSRange Is Nothing Then
        Set oNSRange = oWorksheet.Cells(lngRow, lngASCol)
      Else
        Set oNSRange = oWorksheet.Application.Union(oNSRange, oWorksheet.Cells(lngRow, lngASCol))
      End If
      If oNFRange Is Nothing Then
        Set oNFRange = oWorksheet.Cells(lngRow, lngAFCol)
      Else
        Set oNFRange = oWorksheet.Application.Union(oNFRange, oWorksheet.Cells(lngRow, lngAFCol))
      End If
      If oEVPRange Is Nothing Then
        Set oEVPRange = oWorksheet.Cells(lngRow, lngEVPCol)
      Else
        Set oEVPRange = oWorksheet.Application.Union(oEVPRange, oWorksheet.Cells(lngRow, lngEVPCol))
      End If
      If oEVTRange Is Nothing Then
        Set oEVTRange = oWorksheet.Cells(lngRow, lngEVTCol)
      Else
        Set oEVTRange = oWorksheet.Application.Union(oEVTRange, oWorksheet.Cells(lngRow, lngEVTCol))
      End If
      'todo: assignments vs task
      If oETCRange Is Nothing Then
        Set oETCRange = oWorksheet.Cells(lngRow, lngETCCol)
      Else
        Set oETCRange = oWorksheet.Application.Union(oETCRange, oWorksheet.Cells(lngRow, lngETCCol))
      End If
    End If
        
''    'add EVT comment - this is slow, and often fails
'    oWorksheet.Application.ScreenUpdating = True
'    Set oComment = oWorksheet.Cells(lngRow, lngEVTCol).AddComment(oEVTs.Item(oTask.GetField(FieldNameToFieldConstant(cptStatusSheet_frm.cboEVT.Value))))
'    oComment.Shape.TextFrame.Characters.Font.Bold = False
'    oComment.Shape.TextFrame.AutoSize = True
'    oWorksheet.Application.ScreenUpdating = False
    
    'unlock comment column
    If oUnlockedRange Is Nothing Then
      Set oUnlockedRange = oWorksheet.Cells(lngRow, lngLastCol)
    Else
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow, lngLastCol))
    End If
    
    'export notes
    If cptStatusSheet_frm.chkExportNotes Then
      oWorksheet.Cells(lngRow, lngLastCol) = Trim(Replace(oTask.Notes, vbCr, vbLf))
    End If
    
    'format comments column
    oWorksheet.Cells(lngRow, lngLastCol).HorizontalAlignment = xlLeft
    oWorksheet.Cells(lngRow, lngLastCol).NumberFormat = "General"
    oWorksheet.Cells(lngRow, lngLastCol).WrapText = True
    
get_assignments:
    blnAssignments = CBool(cptGetSetting("StatusSheet", "chkAssignments"))
    If blnAssignments Then
      If oTask.Assignments.Count > 0 And Not IsDate(oTask.ActualFinish) Then
        cptGetAssignmentData oTask, oWorksheet, lngRow, lngHeaderRow, lngNameCol, lngETCCol - 1
      ElseIf IsDate(oTask.ActualFinish) Then
        For Each oAssignment In oTask.Assignments
          Set oAssignment = Nothing
          On Error Resume Next
          Set oAssignment = oTask.Assignments.UniqueID(oWorksheet.Cells(lngRow + 1, 1).Value)
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If Not oAssignment Is Nothing Then
            oWorksheet.Rows(lngRow + 1).EntireRow.Delete
          End If
        Next oAssignment
        Set oAssignment = Nothing
      End If
    Else
      oWorksheet.Cells(lngRow, lngETCCol) = oTask.RemainingWork / 60
      oWorksheet.Cells(lngRow, lngETCCol - 1) = oTask.RemainingWork / 60
      oWorksheet.Cells(lngRow, lngBLWCol) = oTask.BaselineWork / 60
      'add ETC to inputrange
      If oInputRange Is Nothing Then
        Set oInputRange = oWorksheet.Cells(lngRow, lngETCCol)
      Else
        Set oInputRange = oWorksheet.Application.Union(oInputRange, oWorksheet.Cells(lngRow, lngETCCol))
      End If
      'add to ETC Validation Range
      If oETCValidationRange Is Nothing Then
        Set oETCValidationRange = oWorksheet.Cells(lngRow, lngETCCol)
      Else
        Set oETCValidationRange = oWorksheet.Application.Union(oETCValidationRange, oWorksheet.Cells(lngRow, lngETCCol))
      End If
    End If
        
    oWorksheet.Columns(1).AutoFit
    oWorksheet.Rows(lngRow).AutoFit

next_task:
    lngTask = lngTask + 1
    cptStatusSheet_frm.lblProgress.Width = (lngTask / lngTasks) * cptStatusSheet_frm.lblStatus.Width
  Next oTask
  
  'clear out group summary stuff
  If ActiveProject.CurrentGroup <> "No Group" Then
    lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row
    For lngRow = lngHeaderRow + 1 To lngLastRow
      If Not blnAssignments Then
        'remove UID on Group Summaries
        Set oTask = Nothing
        On Error Resume Next
        Set oTask = ActiveProject.Tasks.UniqueID(oWorksheet.Cells(lngRow, 1))
        If oTask Is Nothing Then
          oWorksheet.Cells(lngRow, 1).ClearContents
        Else
          If Trim(oTask.Name) <> Trim(oWorksheet.Cells(lngRow, lngNameCol).Value) Then
            oWorksheet.Cells(lngRow, 1).ClearContents
          End If
        End If
      End If
      If Len(oWorksheet.Cells(lngRow, 1)) = 0 Then
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
      End If
    Next lngRow
  End If
  
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
      .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=FormatDateTime(ActiveProject.ProjectStart, vbShortDate), Formula2:="12/31/2149"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Date Only"
      .ErrorTitle = "Date Only"
      .InputMessage = "Please enter a date between " & FormatDateTime(ActiveProject.ProjectStart, vbShortDate) & " and 12/31/2149 in 'm/d/yyyy' format."
      .ErrorMessage = "Please enter a date between " & FormatDateTime(ActiveProject.ProjectStart, vbShortDate) & " and 12/31/2149 in 'm/d/yyyy' format."
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If blnValidation And Not oNumberValidationRange Is Nothing Then
    'number validation range (contains EV% only)
    With oNumberValidationRange.Validation
      .Delete
      .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="1"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Number Only"
      .ErrorTitle = "Number Only"
      .InputMessage = "Please enter a percentage between 0% and 100%."
      .ErrorMessage = "Please enter a percentage between 0% and 100%."
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If blnValidation And Not oETCValidationRange Is Nothing Then
    'ETC validation range (contains ETC only)
    With oETCValidationRange.Validation
      .Delete
      .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
      .IgnoreBlank = True
      .InCellDropdown = True
      .InputTitle = "Number Only"
      .ErrorTitle = "Number Only"
      .InputMessage = "Please enter a number greater than, or equal to, zero (0)."
      .ErrorMessage = "Please enter a number greater than, or equal to, zero (0)"
      .ShowInput = True
      .ShowError = True
    End With
  End If
  If Not blnAssignments And Not oETCValidationRange Is Nothing Then
    oETCValidationRange.Locked = False
    oETCValidationRange.HorizontalAlignment = xlCenter
  End If
  
  'format the Assignment Rows
  If Not oAssignmentRange Is Nothing Then
    With oAssignmentRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -4.99893185216834E-02
      .PatternTintAndShade = 0
    End With
  End If
  'unlock the input cells
  If Not oInputRange Is Nothing Then
    'oInputRange.Style = "Input" todo
    oInputRange.Locked = False
  End If
  If blnLocked And Not oUnlockedRange Is Nothing Then oUnlockedRange.Locked = False
  If Not oTwoWeekWindowRange Is Nothing Then
    'oTwoWeekWindowRange.Style = "Neutral" todo
    oTwoWeekWindowRange.Locked = False
  End If
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
      strEVTList = strEVTList & "0 - No EVM required,"
      strEVTList = strEVTList & "1 - 0/100,"
      strEVTList = strEVTList & "'2 - 25/75,"
      strEVTList = strEVTList & "'3 - 40/60,"
      strEVTList = strEVTList & "'4 - 50/50,"
      strEVTList = strEVTList & "5 - % Complete,"
      strEVTList = strEVTList & "6 - LOE,"
      strEVTList = strEVTList & "7 - Earned Standards,"
      strEVTList = strEVTList & "8 - Milestone Weights,"
      strEVTList = strEVTList & "9 - BCWP Entry,"
      strEVTList = strEVTList & "A - Apportioned,"
      strEVTList = strEVTList & "P - Milestone Weights with % Complete,"
      strEVTList = strEVTList & "K - Key Event,"
    End If
    If Len(strEVTList) > 0 Then
      oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).Value = "Earned Value Techniques (EVT)"
      oWorksheet.Cells(lngHeaderRow, lngLastCol).Copy
      oWorksheet.Cells(lngHeaderRow, lngLastCol + 2).PasteSpecial xlPasteFormats
      oWorksheet.Range(oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2), oWorksheet.Cells(lngHeaderRow + 1, lngLastCol + 2).Offset(UBound(Split(strEVTList, ",")), 0)).Value = oWorksheet.Application.Transpose(Split(strEVTList, ","))
      oWorksheet.Columns(lngLastCol + 2).AutoFit
    End If
    
  End If
  
  If blnConditionalFormats Then
    
    'entry required cells:
    'green by nature or green as last condition
    'if empty then "input"
    'if invalid then "red"
    'if valid then "green"
    
    oNSRange.Select
    Set oFirstCell = oWorksheet.Application.ActiveCell
    oFirstCell.Select
    strNS = oFirstCell.Address(False, True)
    lngNSCol = lngASCol 'new start
    lngNFCol = lngAFCol 'new finish
    lngCSCol = oWorksheet.Cells(lngHeaderRow).Find(what:="Forecast Start", lookat:=xlWhole).Column
    lngCFCol = oWorksheet.Cells(lngHeaderRow).Find(what:="Forecast Finish", lookat:=xlWhole).Column
    lngCEVPCol = oWorksheet.Cells(lngHeaderRow).Find(what:="EV%", lookat:=xlWhole).Column
    lngCETCCol = oWorksheet.Cells(lngHeaderRow).Find(what:="ETC", lookat:=xlWhole).Column
    strCS = oWorksheet.Cells(oFirstCell.Row, lngCSCol).Address(False, True)
    strCF = oWorksheet.Cells(oFirstCell.Row, lngCFCol).Address(False, True)
    strNF = oWorksheet.Cells(oFirstCell.Row, lngNFCol).Address(False, True)
    strEVP = oWorksheet.Cells(oFirstCell.Row, lngEVPCol).Address(False, True)
    strCEVP = oWorksheet.Cells(oFirstCell.Row, lngCEVPCol).Address(False, True)
    strETC = oWorksheet.Cells(oFirstCell.Row, lngETCCol).Address(False, True)
    strCETC = oWorksheet.Cells(oFirstCell.Row, lngCETCCol).Address(False, True)
    strEVT = oWorksheet.Cells(oFirstCell.Row, lngEVTCol).Address(False, True)
    'set up derived addresses for ease of formula writing
    'AS = (NS>0,NS<=SD)
    strAS = strNS & ">0," & strNS & "<=STATUS_DATE"
    'AF = (NF>0,NF<=SD)
    strAF = strNF & ">0," & strNF & "<=STATUS_DATE"
    'FS = (NS>0,NS>SD)
    strFS = strNS & ">0," & strNS & ">STATUS_DATE"
    'FF = (NF>0,NF>SD)
    strFF = strNF & ">0," & strNF & ">STATUS_DATE"
    
    'create map of ranges
    Set oDict = CreateObject("Scripting.Dictionary")
    Set oDict.Item("NS") = oNSRange
    oNSRange.FormatConditions.Delete
    Set oDict.Item("NF") = oNFRange
    oNFRange.FormatConditions.Delete
    Set oDict.Item("EVP") = oEVPRange
    oEVPRange.FormatConditions.Delete
    Set oDict.Item("EVT") = oEVTRange
    oEVTRange.FormatConditions.Delete
    Set oDict.Item("ETC") = oETCRange
    oETCRange.FormatConditions.Delete
    Dim oAssignmentETCRange As Excel.Range
    Set oAssignmentETCRange = oWorksheet.Application.Intersect(oAssignmentRange, oWorksheet.Columns(lngETCCol))
    Set oDict.Item("AssignmentETC") = oAssignmentETCRange
    oAssignmentETCRange.FormatConditions.Delete
    
    'capture list of formulae
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Fields.Append "RANGE", adVarChar, 13
    oRecordset.Fields.Append "FORMULA", adVarChar, 255
    oRecordset.Fields.Append "FORMAT", adVarChar, 10
    oRecordset.Fields.Append "STOP", adInteger
    oRecordset.Open
    'todo: formula = "=IF('{breadcrumb}','{format}',{formula}')
    'todo: output list of format conditions to header comment or text file or something
    
    '<cpt-breadcrumbs:format-conditions>
    'VARIABLES:
    'SD = Status Date
    'CS = Current Start
    'CF = Current Finish
    'NS = New Start
    'NF = New Finish
    'EVT = Earned Value Technique
    'EVP = Earned Value Percent
    'ETC = Estimate to Complete
    '
    'DERIVED VARIABLES:
    'AS = (NS>0,NS<=SD)
    'AF = (NF>0,NF<=SD)
    'FS = (NS>0,NS>SD)
    'FF = (NF>0,NF>SD)
    '
    'todo: change COMPLETE to gray italics
    'todo: reorder this section to control exact order of application and use STOP
    'NS:AND(NS>0,NS<=SD) -> COMPLETE
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strAS & ")", "COMPLETE")
    'NS:AND(NS>0,NS>SD) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & ">0," & strNS & ">STATUS_DATE)", "GOOD")
    'NS:AND(CS<=(SD+14),NS=0) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strCS & "<=(STATUS_DATE+14)," & strNS & "=0)", "NEUTRAL")
    'NS:AND(CS<=SD,NS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strCS & "<=STATUS_DATE," & strNS & "=0)", "BAD") 'should have started
    'NS:AND(NS>0,NF>0,NS>NF) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & ">0," & strNF & ">0," & strNS & ">" & strNF & ")", "BAD")
    'todo:oRecordset.AddNew Array(0, 1, 2), Array("NS", "=IF(""NS>0,NF>0,NS>NF,"",""BAD"",AND(" & strNS & ">0," & strNF & ">0," & strNS & ">" & strNF & "))","BAD")
    'NS:AND(NS=0,AF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & "=0," & strAF & ")", "BAD")
    'NS:AND(FS>0,EVP>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strFS & "," & strEVP & ">0)", "BAD")
    'NS:AND(NS=0,EVP>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & "=0," & strEVP & ">0)", "BAD")
    'NS:AND(FS>0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strFS & "," & strETC & "=0)", "BAD")
    'NS:AND(AS=0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NS", "=AND(" & strNS & "=0," & strETC & "=0)", "BAD")
    
    'NF:AND(NF>0,NF<=SD) -> COMPLETE
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & ")", "COMPLETE")
    'NF:AND(NF>0,NF>SD) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & ">0," & strNF & ">STATUS_DATE)", "GOOD")
    'NF:AND(CF<=(SD+14),NF=0) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strCF & "<=(STATUS_DATE+14)," & strNF & "=0)", "NEUTRAL")
    'NF:AND(CF<=SD,NF=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strCF & "<=STATUS_DATE," & strNF & "=0)", "BAD") 'should have finished
    'NF:AND(AS,NF=0) -> INPUT
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
    'NF:AND(NF>0,NS>0,NF<NS) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & ">0," & strNS & ">0," & strNS & ">" & strNF & ")", "BAD")
    'NF:AND(AF>0,NS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & "," & strNS & "=0)", "BAD")
    'NF:AND(FF>0,EVP=1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strFF & "," & strEVP & "=1)", "BAD")
    'NF:AND(AF,EVP<1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & "," & strEVP & "<1)", "BAD")
    'NF:AND(NF=0,EVP=1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & "=0," & strEVP & "=1)", "BAD")
    'NF:AND(FF>0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strFF & "," & strETC & "=0)", "BAD")
    'NF:AND(AF>0,ETC>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strAF & "," & strETC & ">0)", "BAD")
    'NF:AND(NF=0,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("NF", "=AND(" & strNF & "=0," & strETC & "=0)", "BAD")
    
    'EVP:AND(FF,NEW EVP>EVP) -> GOOD
    'todo: keeping FF forces FF before all good
    'todo: remove FF; add last and add stop if true to isolate this good update
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strFF & "," & strEVP & ">" & strCEVP & ")", "GOOD")
    'EVP:AND(AF,EVP=1) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strAF & "," & strEVP & "=1)", "GOOD")
    'EVP:AND(FF,EVP=PREVIOUS) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strFF & "," & strEVP & "=" & strCEVP & ")", "NEUTRAL")
    'EVP:AND(AS,NF=0) -> INPUT
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
    'EVP:AND(EVP>0,FS>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & ">0," & strFS & ")", "BAD")
    'EVP:AND(EVP=1,FF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "=1," & strFF & ")", "BAD")
    'EVP:AND(EVP=1,NF=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "=1," & strNF & "=0)", "BAD")
    'EVP:AND(EVP<1,AF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "<1," & strAF & ")", "BAD")
    'EVP:AND(EVP=1,ETC>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "=1," & strETC & ">0)", "BAD")
    'EVP:AND(EVP<1,ETC=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "<1," & strETC & "=0)", "BAD")
    'EVP:AND(EVP>0,AS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & ">0," & strNS & "=0)", "BAD")
    'EVP:AND(EVP<PREVIOUS) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("EVP", "=AND(" & strEVP & "<" & strCEVP & ")", "BAD")
    
    'ETC:AND(FF,NEW ETC<>ETC) -> GOOD
    'todo: keeping FF forces FF before all good
    'todo: remove FF; add last and add stop if true to isolate this good update
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strFF & "," & strETC & "<>" & strCETC & ")", "GOOD")
    'ETC:AND(AF,ETC=0) -> GOOD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strAF & "," & strETC & "=0," & strEVP & "=1)", "GOOD")
    'ETC:AND(FF,ETC=PREVIOUS) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strFF & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
    'ETC:AND(FS,ETC=PREVIOUS) -> NEUTRAL
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strFS & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
    'todo: what is oAssignmentRange
    If Not blnAssignments Then
      'ETC:AND(AS,NF=0) -> INPUT
      oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
    
    Else
      'ETC:AND(FF,ETC=PREVIOUS) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFF & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
      'ETC:AND(FS,ETC=PREVIOUS) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFS & "," & strETC & "=" & strCETC & ")", "NEUTRAL")
      'ETC:AND(FF,ETC=0) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFF & "," & strETC & "=0)", "NEUTRAL")
      'ETC:AND(FS,ETC=0) -> NEUTRAL (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strFS & "," & strETC & "=0)", "NEUTRAL")
      'ETC:AND(AS,NF=0) -> INPUT (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strAS & "," & strNF & "=0)", "INPUT") 'in progress
      'ETC:AND(ETC>0,AF>0) -> BAD (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strETC & ">0," & strAF & ")", "BAD")
      'ETC:AND(ETC>0,EVP=1) -> BAD (ASSIGNMENT)
      oRecordset.AddNew Array(0, 1, 2), Array("AssignmentETC", "=AND(" & strETC & ">0," & strEVP & "=1)", "BAD")
    End If
    
    'ETC:AND(ETC=0,FS>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strFS & ")", "BAD")
    'ETC:AND(ETC=0,FF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strFF & ")", "BAD")
    'ETC:AND(ETC>0,AF>0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & ">0," & strAF & ")", "BAD")
    'ETC:AND(ETC>0,EVP=1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & ">0," & strEVP & "=1)", "BAD")
    'ETC:AND(ETC=0,EVP<1) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strEVP & "<1)", "BAD")
    'ETC:AND(ETC=0,EVP=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strEVP & "=0)", "BAD")
    'ETC:AND(ETC=0,AF=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strNF & "=0)", "BAD")
    'ETC:AND(ETC=0,AS=0) -> BAD
    oRecordset.AddNew Array(0, 1, 2), Array("ETC", "=AND(" & strETC & "=0," & strNS & "=0)", "BAD")
    
    Dim blnMilestones As Boolean
    If blnMilestones Then 'assumes COBRA and field values = COBRA codes
      'todo: AS>0,EVT='E',EVP<>50
      'todo: oRecordset.AddNew Array(0,1),Array("NS", "=AND(" & strAS & "," & strEVT & "='E'," & strEVP & "<>.5)")
      'todo: AS>0,EVT='F',EVP>0
      'todo: oRecordset.AddNew Array(0,1),Array("NS", "=AND(" & strAS & "," & strEVT & "='F'," & strEVP & ">0)")
      'todo: AS>0,EVT='G',EVP<>1
      'todo: EVP<>.5,EVT='E",AS>0
      'todo: oRecordset.AddNew Array(0,1),Array("EVP", "=AND(" & strEVP & "<>.5," & strEVT & "='E'," & strAS & ">0)")
      'todo: EVP>0,EVT='F',AS>0
      'todo: oRecordset.AddNew Array(0,1),Array("EVP", "=AND(" & strEVP & ">0," & strEVT & "='F'," & strAS & ">0)")
      'todo: EVP<1,EVT='G',AS>0

    End If
    '</cpt-breadcrumbs:format-conditions>
skip_working:
    lngFormatCondition = 0
    With oRecordset
      'for the progress bar
      lngFormatConditions = .RecordCount
      .MoveFirst
      Do While Not .EOF
        'race is on
        lngFormatCondition = lngFormatCondition + 1
        cptStatusSheet_frm.lblStatus.Caption = "Applying Conditional Formatting [" & strItem & "]...(" & Format(lngFormatCondition / lngFormatConditions, "0%") & ")"
        cptStatusSheet_frm.lblProgress.Width = (lngFormatCondition / lngFormatConditions) * cptStatusSheet_frm.lblStatus.Width
        Application.StatusBar = "Applying Conditional Formatting [" & strItem & "]...(" & Format(lngFormatCondition / lngFormatConditions, "0%") & ")"
        Set oFormatRange = oDict.Item(CStr(.Fields(0)))
        oFormatRange.Select
        oFormatRange.FormatConditions.Add Type:=xlExpression, Formula1:=CStr(.Fields(1))
        oFormatRange.FormatConditions(oFormatRange.FormatConditions.Count).SetFirstPriority
        If .Fields(2) = "BAD" Then
          oFormatRange.FormatConditions(1).Font.Color = -16383844
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
          oFormatRange.FormatConditions(1).Interior.Color = 13551615
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "CALCULATION" Then
          oFormatRange.FormatConditions(1).Font.Color = 32250
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = -4105
          oFormatRange.FormatConditions(1).Interior.Color = 15921906
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "COMPLETE" Then
          oFormatRange.FormatConditions(1).Font.Color = 8355711
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = -4105
          oFormatRange.FormatConditions(1).Interior.Color = 15921906
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "GOOD" Then
          oFormatRange.FormatConditions(1).Font.Color = -16752384
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
          oFormatRange.FormatConditions(1).Interior.Color = 13561798
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
        ElseIf .Fields(2) = "INPUT" Then
          oFormatRange.FormatConditions(1).Font.Color = 7749439
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = -4105
          oFormatRange.FormatConditions(1).Interior.Color = 10079487
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
          'oFormatRange.FormatConditions(1).BorderAround xlContinuous, xlThin, , Color:=RGB(127, 127, 127)
        ElseIf .Fields(2) = "NEUTRAL" Then
          oFormatRange.FormatConditions(1).Font.Color = -16754788
          oFormatRange.FormatConditions(1).Font.TintAndShade = 0
          oFormatRange.FormatConditions(1).Interior.PatternColorIndex = xlAutomatic
          oFormatRange.FormatConditions(1).Interior.Color = 10284031
          oFormatRange.FormatConditions(1).Interior.TintAndShade = 0
          'oFormatRange.FormatConditions(1).BorderAround xlContinuous, xlThin, , Color:=RGB(127, 127, 127)
        End If
        oFormatRange.FormatConditions(1).StopIfTrue = False 'CBool(oRecordset(3)) 'todo?
        .MoveNext
      Loop
      'race is over - notify
      cptStatusSheet_frm.lblStatus.Caption = "Applying Conditional Formatting [" & strItem & "]...done."
      cptStatusSheet_frm.lblProgress.Width = cptStatusSheet_frm.lblStatus.Width
      Application.StatusBar = "Applying Conditional Formatting [" & strItem & "]...done."
      oDict.RemoveAll
      .Close
    End With
            
  End If
  
exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Set oFormatRange = Nothing
  Set oDict = Nothing
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  Set oFirstCell = Nothing
  Set oETCRange = Nothing
  Set oEVPRange = Nothing
  Set oNFRange = Nothing
  Set oNSRange = Nothing
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

Private Sub cptGetAssignmentData(ByRef oTask As MSProject.Task, ByRef oWorksheet As Excel.Worksheet, lngRow As Long, lngHeaderRow As Long, lngNameCol As Long, lngRemainingWorkCol As Long)
  'objects
  Dim oAssignment As Assignment
  'strings
  Dim strAllowAssignmentNotes As String
  Dim strProtect As String
  Dim strDataValidation As String
  'longs
  Dim lngEVPCol As Long
  Dim lngEVTCol  As Long
  Dim lngNFCol As Long
  Dim lngNSCol As Long
  Dim lngFFCol As Long
  Dim lngFSCol As Long
  Dim lngBaselineCostCol As Long
  Dim lngBaselineWorkCol As Long
  Dim lngIndent As Long
  Dim lngItem As Long
  Dim lngLastCol As Long
  Dim lngLastRow As Long
  Dim lngCol As Long
  'integers
  'doubles
  'booleans
  Dim blnAllowAssignmentNotes As Boolean
  'variants
  Dim vCol As Variant
  Dim vAssignment As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  lngIndent = Len(cptRegEx(oWorksheet.Cells(lngRow, lngNameCol).Value, "^\s*"))
  lngLastCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
  lngLastRow = oWorksheet.Cells(1048576, 1).End(xlUp).Row
  'get column for FS,FF,NS,NF,EVT,EVP
  
  lngFSCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Forecast Start", lookat:=xlWhole).Column
  lngFFCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Forecast Finish", lookat:=xlWhole).Column
  lngNSCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Actual Start", lookat:=xlPart).Column
  lngNFCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Actual Finish", lookat:=xlPart).Column
  'todo: lngEVTCol = oWorksheet.Rows(lngHeaderRow).Find(what:="EVT", lookat:=xlWhole).Column - Milestone EVT?
  lngEVPCol = oWorksheet.Rows(lngHeaderRow).Find(what:="New EV%", lookat:=xlWhole).Column
  
  lngItem = 0
  For Each oAssignment In oTask.Assignments
    lngItem = lngItem + 1
    If ActiveProject.CurrentView <> "Task Usage" Or IsDate(oAssignment.ActualFinish) Then
      If Trim(oWorksheet.Cells(lngRow + lngItem, lngNameCol).Value) <> oAssignment.ResourceName Then
        oWorksheet.Rows(lngRow + lngItem).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Font.ColorIndex = xlAutomatic
      Else
        oWorksheet.Rows(lngRow + lngItem).ClearContents
      End If
    Else
      oWorksheet.Rows(lngRow + lngItem).ClearContents
    End If
    'this fills down task custom fields to assignments
    For lngCol = 2 To lngNameCol
      If lngCol <> lngNameCol Then oWorksheet.Cells(lngRow + lngItem, lngCol) = oWorksheet.Cells(lngRow, lngCol)
    Next lngCol
    
    oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Font.Italic = True 'todo: limit to columns
    vAssignment = oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Value
    vAssignment(1, 1) = oAssignment.UniqueID 'import assumes this is oAssignment.UniqueID
    vAssignment(1, lngNameCol) = String(lngIndent + 3, " ") & oAssignment.ResourceName
    If oAssignment.ResourceType = pjWork Then
      lngBaselineWorkCol = oWorksheet.Rows(lngHeaderRow).Find("Baseline Work", lookat:=xlWhole).Column
      vAssignment(1, lngBaselineWorkCol) = oAssignment.BaselineWork / 60
      vAssignment(1, lngRemainingWorkCol) = oAssignment.RemainingWork / 60
      vAssignment(1, lngRemainingWorkCol + 1) = oAssignment.RemainingWork / 60
    Else
      lngBaselineCostCol = oWorksheet.Rows(lngHeaderRow).Find("Baseline Work", lookat:=xlWhole).Column
      vAssignment(1, lngBaselineCostCol) = oAssignment.BaselineCost
      vAssignment(1, lngRemainingWorkCol) = oAssignment.RemainingCost
      vAssignment(1, lngRemainingWorkCol + 1) = oAssignment.RemainingCost
    End If

    'fill down NS,NF,EVP todo: add lngETCCol also
    For Each vCol In Array(lngFSCol, lngFFCol, lngNSCol, lngNFCol, lngEVPCol)
      vAssignment(1, vCol) = "=" & oWorksheet.Cells(lngRow, vCol).AddressLocal(False, True)
      oWorksheet.Cells(lngRow + lngItem, vCol).Font.ThemeColor = xlThemeColorDark1
      oWorksheet.Cells(lngRow + lngItem, vCol).Font.TintAndShade = -4.99893185216834E-02
      If vCol = lngEVPCol Then oWorksheet.Cells(lngRow + lngItem, lngEVPCol).NumberFormat = "0%"
    Next vCol

    'add validation
    If oETCValidationRange Is Nothing Then
      Set oETCValidationRange = oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1)
    Else
      Set oETCValidationRange = oWorksheet.Application.Union(oETCValidationRange, oWorksheet.Cells(lngRow + lngItem, lngRemainingWorkCol + 1))
    End If
    'allow input on ETC if task is unstarted or incomplete - i.e., in progress
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
    
    'export assignment notes
    If cptStatusSheet_frm.chkExportNotes And Len(oAssignment.Notes) > 0 Then
      vAssignment(1, lngLastCol) = Trim(Replace(oAssignment.Notes, vbCr, vbLf))
    End If
    'allow notes at the assignment level?
    strAllowAssignmentNotes = cptGetSetting("StatusSheet", "chkAllowAssignmentNotes")
    If strAllowAssignmentNotes <> "" Then
      blnAllowAssignmentNotes = CBool(strAllowAssignmentNotes)
    Else
      blnAllowAssignmentNotes = False
    End If
    If blnAllowAssignmentNotes Then
      Set oUnlockedRange = oWorksheet.Application.Union(oUnlockedRange, oWorksheet.Cells(lngRow + lngItem, lngLastCol))
    End If
    oWorksheet.Cells(lngRow + lngItem, lngLastCol).HorizontalAlignment = xlLeft
    oWorksheet.Cells(lngRow + lngItem, lngLastCol).NumberFormat = "General"
    oWorksheet.Cells(lngRow + lngItem, lngLastCol).WrapText = True
    
    'enter the values
    oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)).Value = vAssignment
    
    If oAssignmentRange Is Nothing Then
      Set oAssignmentRange = oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol))
    Else
      Set oAssignmentRange = oWorksheet.Application.Union(oAssignmentRange, oWorksheet.Range(oWorksheet.Cells(lngRow + lngItem, 1), oWorksheet.Cells(lngRow + lngItem, lngLastCol)))
    End If
    oWorksheet.Rows(lngRow + lngItem).AutoFit
  Next oAssignment
  'add formulae
  If oTask.Assignments.Count > 0 Then
    'baseline work
    oWorksheet.Cells(lngRow, lngRemainingWorkCol - 1).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C" & lngRemainingWorkCol - 1 & ":R" & lngRow + lngItem & "C" & lngRemainingWorkCol - 1 & ")"
    'prev etc
    oWorksheet.Cells(lngRow, lngRemainingWorkCol).FormulaR1C1 = "=SUM(R" & lngRow + 1 & "C" & lngRemainingWorkCol & ":R" & lngRow + lngItem & "C" & lngRemainingWorkCol & ")"
    'new etc
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

Sub cptFinalFormats(ByRef oWorksheet As Excel.Worksheet)
  Dim lngHeaderRow As Long
  Dim vBorder As Variant
  
  lngHeaderRow = 8
  oWorksheet.Cells(lngHeaderRow, 1).AutoFilter
  oWorksheet.Columns(1).AutoFit
  With oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight), oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp))
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
  oWorksheet.Application.WindowState = xlNormal 'cannot apply certain settings below if window is minimized...like data validation
  oWorksheet.Application.Calculation = xlCalculationAutomatic
  oWorksheet.Application.ScreenUpdating = True
  oWorksheet.Application.ActiveWindow.DisplayGridlines = False
  oWorksheet.[B2].Select
  oWorksheet.Application.ActiveWindow.SplitRow = 8
  oWorksheet.Application.ActiveWindow.SplitColumn = 0
  oWorksheet.Application.ActiveWindow.FreezePanes = True
  oWorksheet.Application.ActiveWindow.DisplayHorizontalScrollBar = True
  oWorksheet.Application.ActiveWindow.DisplayVerticalScrollBar = True
  oWorksheet.Application.WindowState = xlMinimized
  Set oEntryHeaderRange = Nothing
End Sub

Sub cptListQuickParts(Optional blnRefreshOutlook As Boolean = False)
  'objects
  Dim oOutlook As Outlook.Application
  Dim oMailItem As MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oTemplate As Word.Template
  Dim oBuildingBlockEntries As Word.BuildingBlockEntries
  Dim oBuildingBlock As Word.BuildingBlock
  'longs
  Dim lngItem As Long
  'strings
  Dim strQuickPartList As String
  Dim strSQL As String
  'variants
  Dim vQuickPart As Variant
  Dim vQuickParts As Variant

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If blnRefreshOutlook Then
    'refresh QuickParts in Outlook
    cptStatusSheet_frm.cboQuickParts.Clear
    'get Outlook
    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application") 'this works even if Outlook isn't open
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oOutlook Is Nothing Then
      Set oOutlook = CreateObject("Outlook.Application")
    End If
    'create MailItem, insert quickparts, update links, dates
    Set oMailItem = oOutlook.CreateItem(olMailItem)
    If oMailItem.BodyFormat <> olFormatHTML Then oMailItem.BodyFormat = olFormatHTML
    On Error Resume Next
    Set oDocument = oMailItem.GetInspector.WordEditor
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oDocument Is Nothing Then
      'try again with MailItem displayed
      oMailItem.Display False
      On Error Resume Next
      Set oDocument = oMailItem.GetInspector.WordEditor
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oDocument Is Nothing Then
        'todo: try again by accessing Word directly
        cptStatusSheet_frm.cboQuickParts.Enabled = False
        oMailItem.Close olDiscard
        GoTo exit_here
      Else
        oMailItem.GetInspector.WindowState = olMinimized
      End If
    End If
    Set oWord = oDocument.Application
    Set oTemplate = oWord.Templates(1)
    Set oBuildingBlockEntries = oTemplate.BuildingBlockEntries
    'loop through them
    For lngItem = 1 To oBuildingBlockEntries.Count
      Set oBuildingBlock = oBuildingBlockEntries(lngItem)
      If oBuildingBlock.Type.Name = "Quick Parts" Then
        strQuickPartList = strQuickPartList & oBuildingBlock.Name & ","
      End If
    Next lngItem
    'sort them
    If Len(strQuickPartList) > 0 Then
      strQuickPartList = Left(strQuickPartList, Len(strQuickPartList) - 1)
      vQuickParts = Split(strQuickPartList, ",")
      cptQuickSort vQuickParts, 0, UBound(vQuickParts)
      For Each vQuickPart In vQuickParts
        cptStatusSheet_frm.cboQuickParts.AddItem vQuickPart
      Next vQuickPart
    End If
    oMailItem.Close olDiscard
  End If
    
exit_here:
  On Error Resume Next
  Set oWord = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oTemplate = Nothing
  Set oBuildingBlockEntries = Nothing
  Set oBuildingBlock = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptListQuickParts", Err)
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
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  dtStatus = ActiveProject.StatusDate

  With cptStatusSheet_frm
    strDir = .lblDirSample.Caption
    'create the status date directory
    If Dir(strDir, vbDirectory) = vbNullString Then
      MkDir strDir
      oWorkbook.Application.Wait Now + TimeValue("00:00:03")
    End If
    strFileName = .txtFileName.Value & ".xlsx"
    strFileName = Replace(strFileName, "[yyyy-mm-dd]", Format(dtStatus, "yyyy-mm-dd"))
    strFileName = Replace(strFileName, "[program]", cptGetProgramAcronym)
    If Len(strItem) > 0 Then
      strFileName = Replace(strFileName, "[item]", strItem)
    End If
    strFileName = cptRemoveIllegalCharacters(strFileName)
    On Error Resume Next
    If Dir(strDir & strFileName) <> vbNullString Then
      Kill strDir & strFileName
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
      DoEvents
    End If
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    'account for if the file exists and is open in the background
    If Dir(strDir & strFileName) <> vbNullString Then  'delete failed, rename with timestamp
      strMsg = "'" & strFileName & "' already exists, and is likely open." & vbCrLf
      strFileName = Replace(strFileName, ".xlsx", "_" & Format(Now, "hh-nn-ss") & ".xlsx")
      strMsg = strMsg & "The file you are now creating will be named '" & strFileName & "'"
      MsgBox strMsg, vbExclamation + vbOKOnly, "NOTA BENE"
      oWorkbook.SaveAs strDir & strFileName, 51
      oWorkbook.Application.Wait Now + TimeValue("00:00:02")
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
  Dim oInspector As Outlook.Inspector
  Dim oBuildingBlock As Word.BuildingBlock
  Dim oOutlook As Outlook.Application
  Dim oMailItem As Outlook.MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oSelection As Word.Selection
  Dim oEmailTemplate As Word.Template
  'strings
  Dim strTempItem As String
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
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set oMailItem = oOutlook.CreateItem(0) '0 = olMailItem
  oMailItem.Display False
  oMailItem.Attachments.Add strFullName
  With cptStatusSheet_frm
    strSubject = .txtSubject
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[status\_date\]"), FormatDateTime(ActiveProject.StatusDate, vbShortDate))
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[yyyy\-mm\-dd\]"), Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[item\]"), strItem)
    strSubject = Replace(strSubject, cptRegEx(strSubject, "\[program\]"), cptGetProgramAcronym)
    oMailItem.Subject = strSubject
    oMailItem.CC = .txtCC
    If oMailItem.BodyFormat <> 2 Then oMailItem.BodyFormat = 2 '2=olFormatHTML
    If Not IsNull(.cboQuickParts.Value) And .cboQuickParts.Enabled Then
      If Len(.cboQuickParts.Value) = 0 Then GoTo skip_QuickPart
      Set oDocument = oMailItem.GetInspector.WordEditor
      Set oWord = oDocument.Application
      Set oSelection = oDocument.Windows(1).Selection
      Set oEmailTemplate = oWord.Templates(1)
      On Error Resume Next
      Set oBuildingBlock = oEmailTemplate.BuildingBlockEntries(.cboQuickParts)
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oBuildingBlock Is Nothing Then
        MsgBox "Quick Part '" & .cboQuickParts & "' not found!", vbExclamation + vbOKOnly, "Missing Quick Part"
      Else
        oBuildingBlock.Insert oSelection.Range, True
      End If
      'only do replacements if QuickPart is used
      'clean date format
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(Y|y){1,}-(M|m){1,}-(D|d){1,}\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
      End If
      'clean status date
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(S|s)(T|t)(A|a)(T|t)(U|u)(S|s).(D|d)(A|a)(T|t)(E|e)\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, FormatDateTime(ActiveProject.StatusDate, vbShortDate))
      End If
      'clean program
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(P|p)(R|r)(O|o)(G|g)(R|r)(A|a)(M|m)\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, cptGetProgramAcronym)
      End If
      'clean item
      strTempItem = cptRegEx(oMailItem.HTMLBody, "\[(I|i)(T|t)(E|e)(M|m)\]")
      If Len(strTempItem) > 0 Then
        oMailItem.HTMLBody = Replace(oMailItem.HTMLBody, strTempItem, strItem)
      End If
    End If
skip_QuickPart:
    On Error Resume Next
    Set oInspector = oMailItem.GetInspector
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oInspector Is Nothing Then
      oInspector.WindowState = 1 '1=olMinimized
    Else
      'todo: how to minimize?
    End If
      
  End With
  
exit_here:
  On Error Resume Next
  Set oInspector = Nothing
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
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFileName As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With cptStatusSheet_frm
    'save settings
    cptDeleteSetting "StatusSheet", "cboEVP" 'moved to Integration
    cptSaveSetting "StatusSheet", "cboCostTool", .cboCostTool.Value
    cptDeleteSetting "StatusSheet", "cboEVT" 'moved to Integration
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
    cptSaveSetting "StatusSheet", "chkAssignments", IIf(.chkAssignments, 1, 0)
    cptSaveSetting "StatusSheet", "chkConditionalFormatting", IIf(.chkAddConditionalFormats, 1, 0)
    cptSaveSetting "StatusSheet", "chkEmail", IIf(.chkSendEmails, 1, 0)
    If .chkSendEmails Then
      cptSaveSetting "StatusSheet", "txtSubject", .txtSubject
      cptSaveSetting "StatusSheet", "txtCC", .txtCC
      If Not IsNull(.cboQuickParts.Value) Then
        cptSaveSetting "StatusSheet", "cboQuickPart", .cboQuickParts.Value
      End If
    End If
    If Len(.txtNotesColTitle.Value) > 0 Then
      cptSaveSetting "StatusSheet", "txtNotesColTitle", .txtNotesColTitle.Value
    Else
      cptSaveSetting "StatusSheet", "txtNotesColTitle", "Reason / Action / Impact"
    End If
    cptSaveSetting "StatusSheet", "chkExportNotes", IIf(.chkExportNotes, 1, 0)
    cptSaveSetting "StatusSheet", "chkAllowAssignmentNotes", IIf(.chkAllowAssignmentNotes, 1, 0)
    cptSaveSetting "StatusSheet", "chkKeepOpen", IIf(.chkKeepOpen, 1, 0)
    cptSaveSetting "StatusSheet", "chkConditionalFormatting", IIf(.chkAddConditionalFormats, 1, 0)
    cptSaveSetting "StatusSheet", "chkLookahead", IIf(.chkLookahead, 1, 0)
    If .chkLookahead And Len(.txtLookaheadDays) > 0 Then
      cptSaveSetting "StatusSheet", "txtLookaheadDays", CLng(.txtLookaheadDays.Value)
    End If
    cptSaveSetting "StatusSheet", "chkIgnoreLOE", IIf(.chkIgnoreLOE, 1, 0)
    'save user fields - overwrite
    strFileName = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Fields.Append "Field Constant", adInteger
    oRecordset.Fields.Append "Custom Field Name", adVarChar, 255
    oRecordset.Fields.Append "Local Field Name", adVarChar, 100
    oRecordset.Open
    If .lboExport.ListCount > 0 Then
      For lngItem = 0 To .lboExport.ListCount - 1
        oRecordset.AddNew Array(0, 1, 2), Array(.lboExport.List(lngItem, 0), .lboExport.List(lngItem, 1), .lboExport.List(lngItem, 2))
      Next lngItem
    End If
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    oRecordset.Save strFileName, adPersistADTG
    oRecordset.Close

  End With

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptSaveStatusSheetSettings", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptAdvanceStatusDate()
  Application.ChangeStatusDate
End Sub

Sub cptCaptureJournal()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strProgram As String
  Dim strFile As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strProgram = cptGetProgramAcronym
  
  dtStatus = FormatDateTime(ActiveProject.StatusDate, vbGeneralDate)
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  
  strFile = cptDir & "\settings\cpt-journal.adtg"
  If Dir(strFile) = vbNullString Then
    With oRecordset
      .Fields.Append "PROGRAM", adVarChar, 50
      .Fields.Append "STATUS_DATE", adDate
      .Fields.Append "TASK_UID", adInteger
      .Fields.Append "TASK_NOTE", adVarChar, 255
      .Fields.Append "ASSIGNMENT_UID", adInteger
      .Fields.Append "ASSIGNMENT_NOTE", adVarChar, 255
      .Open
    End With
  Else
    oRecordset.Open strFile
  End If
  Dim oTask As MSProject.Task, oTasks As MSProject.Tasks
  Set oTasks = ActiveProject.Tasks
  lngTasks = oTasks.Count
  lngTask = 0
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Len(oTask.Notes) > 0 Then
      oRecordset.AddNew Array(0, 1, 2, 3), Array(strProgram, dtStatus, oTask.UniqueID, Chr(34) & oTask.Notes & Chr(34))
    End If
    Dim oAssignment As Assignment
    For Each oAssignment In oTask.Assignments
      If Len(oAssignment.Notes) > 0 Then
        oRecordset.AddNew Array(0, 1, 2, 3, 4, 5), Array(strProgram, dtStatus, oTask.UniqueID, oTask.Notes, oAssignment.UniqueID, oAssignment.Notes)
      End If
    Next
next_task:
    lngTask = lngTask + 1
    Debug.Print Format(lngTask / lngTasks, "0%")
  Next oTask
  
  oRecordset.Save strFile
  oRecordset.Close
  
exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCaptureJournal", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCompletedWork()
  'objects
  Dim oCDP As DocumentProperty
  Dim oAssignment As Assignment
  Dim oWorksheet As Object 'Excel.Worksheet
  Dim oWorkbook As Object 'Excel.Workbook
  Dim oExcel As Object 'Excel.Application
  Dim oRecordset As Object 'ADODB.Recordset
  Dim oTask As MSProject.Task
  'strings
  Dim strCA As String
  Dim strEVP As String
  Dim strEVT As String
  Dim strLC As String
  Dim strWPM As String
  Dim strWP As String
  Dim strCAM As String
  Dim strOBS As String
  Dim strWBS As String
  Dim strProgram As String
  Dim strRecord As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  'longs
  Dim lngEVPCol As Long
  Dim lngCA As Long
  Dim lngLC As Long
  Dim lngEVP As Long
  Dim lngEVT As Long
  Dim lngItem As Long
  Dim lngWPM As Long
  Dim lngWP As Long
  Dim lngCAM As Long
  Dim lngOBS As Long
  Dim lngWBS As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  Dim blnMissing As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  Dim dtAF As Date
  
  If Not cptValidMap("WBS,OBS,CAM,CAM,WP,WPM,EVT,EVP", False, False, True) Then
    MsgBox "Settings required. Exiting.", vbExclamation + vbOKOnly, "Invalid Settings"
    GoTo exit_here
  End If
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngWBS = Split(cptGetSetting("Integration", "WBS"), "|")(0)
  strWBS = CustomFieldGetName(lngWBS)
  lngOBS = Split(cptGetSetting("Integration", "OBS"), "|")(0)
  strOBS = CustomFieldGetName(lngOBS)
  lngCA = Split(cptGetSetting("Integration", "CA"), "|")(0)
  strCA = CustomFieldGetName(lngCA)
  lngCAM = Split(cptGetSetting("Integration", "CAM"), "|")(0)
  strCAM = CustomFieldGetName(lngCAM)
  lngWP = Split(cptGetSetting("Integration", "WP"), "|")(0)
  strWP = CustomFieldGetName(lngWP)
  lngWPM = Split(cptGetSetting("Integration", "WPM"), "|")(0)
  strWPM = CustomFieldGetName(lngWPM)
  On Error Resume Next
  Set oCDP = ActiveProject.CustomDocumentProperties("fResID")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oCDP Is Nothing Then
    strLC = ActiveProject.CustomDocumentProperties("fResID")
    lngLC = FieldNameToFieldConstant(strLC, pjResource)
  End If
  lngEVT = Split(cptGetSetting("Integration", "EVT"), "|")(0)
  strEVT = CustomFieldGetName(lngEVT)
  lngEVP = Split(cptGetSetting("Integration", "EVP"), "|")(0)
  strEVP = CustomFieldGetName(lngEVP)
  
  'create Schema
  strFile = Environ("tmp") & "\Schema.ini"
  lngFile = FreeFile
  Open strFile For Output As #lngFile
  Print #lngFile, "[wp.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=UID Long"
  Print #lngFile, "Col2=WBS Text"
  Print #lngFile, "Col3=OBS Text"
  Print #lngFile, "Col4=CA Text"
  Print #lngFile, "Col5=CAM Text"
  Print #lngFile, "Col6=WP Text"
  Print #lngFile, "Col7=WPM Text"
  Print #lngFile, "Col8=LC Text"
  Print #lngFile, "Col9=AF DateTime"
  Print #lngFile, "Col10=PercentComplete Long"
  Close #lngFile
  
  strFile = Environ("tmp") & "\wp.csv"
  lngFile = FreeFile
  Open strFile For Output As #lngFile
  Print #lngFile, "UID,WBS,OBS,CA,CAM,WP,WPM,LC,AF,PercentComplete,"
  
  lngTasks = ActiveProject.Tasks.Count
    
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    For Each oAssignment In oTask.Assignments
      strRecord = oTask.UniqueID & ","
      strRecord = strRecord & oTask.GetField(lngWBS) & ","
      strRecord = strRecord & oTask.GetField(lngOBS) & ","
      strRecord = strRecord & oTask.GetField(lngCA) & ","
      strRecord = strRecord & oTask.GetField(lngCAM) & ","
      strRecord = strRecord & oTask.GetField(lngWP) & ","
      strRecord = strRecord & oTask.GetField(lngWPM) & ","
      If lngLC > 0 Then
        strRecord = strRecord & oAssignment.Resource.GetField(lngLC) & ","
      Else
        strRecord = strRecord & oAssignment.ResourceName & ","
      End If
      If IsDate(oTask.ActualFinish) Then
        dtAF = FormatDateTime(oTask.ActualFinish, vbShortDate)
      Else
        dtAF = #1/1/1984#
      End If
      strRecord = strRecord & dtAF & ","
      strRecord = strRecord & CLng(Replace(oTask.GetField(lngEVP), "%", "")) & "," 'assumes 100
      Print #lngFile, strRecord
    Next oAssignment
    
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = Format(lngTask, "#,##0") & " / " & Format(lngTasks, "#,##0") & "...(" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask
  Close #lngFile
  
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT WP,MAX(AF),AVG(PercentComplete) AS EV "
  strSQL = strSQL & "FROM [wp.csv] "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING AVG(PercentComplete)=100 "
  strSQL = strSQL & "ORDER BY MAX(AF) Desc "
  'strSQL = "SELECT * FROM wp.csv"
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strSQL, strCon, 1, 1 '1=adOpenKeyset, 1=adLockReadOnly
  If oRecordset.RecordCount > 0 Then
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = "COMPLETED WPs"
    oWorksheet.[A1:C1] = Array("WP", "AF", "EV")
    oWorksheet.[A1:C1].Font.Bold = True
    oWorksheet.[A2].CopyFromRecordset oRecordset
    oRecordset.Close
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Columns(2).HorizontalAlignment = -4108 'xlCenter
    oWorksheet.Rows(1).HorizontalAlignment = -4131 'xlLeft
    oWorksheet.[A1].AutoFilter
    oWorksheet.Columns.AutoFit
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    'get details
    If oWorkbook.Sheets.Count >= 2 Then
      Set oWorksheet = oWorkbook.Sheets(2)
    Else
      Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
    End If
    oWorksheet.Name = "DETAILS"
    strSQL = "SELECT * FROM [wp.csv] ORDER BY WP,PercentComplete"
    oRecordset.Open strSQL, strCon, 1, 1 '1=adOpenKeyset, 1=adLockReadOnly
    For lngItem = 0 To oRecordset.Fields.Count - 1
      oWorksheet.Cells(1, lngItem + 1) = oRecordset.Fields(lngItem).Name
    Next lngItem
    oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(-4161)).Font.Bold = True 'xlToRight
    oWorksheet.[A2].CopyFromRecordset oRecordset
    oWorksheet.Columns(8).Replace #1/1/1984#, "NA"
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Columns(8).HorizontalAlignment = -4108 'xlCenter
    oWorksheet.Rows(1).HorizontalAlignment = -4131 'xlLeft
    oWorksheet.[A1].AutoFilter
    oWorksheet.Columns.AutoFit
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 0
    oExcel.ActiveWindow.FreezePanes = True
    lngEVPCol = oWorksheet.Rows(1).Find("PercentComplete", lookat:=xlWhole).Column
    oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).AutoFilter Field:=lngEVPCol, Criteria1:="100"
    oRecordset.Close
    oWorkbook.Sheets("COMPLETED WPs").Activate
    oExcel.Visible = True
    oExcel.ActiveWindow.WindowState = -4143 'xlNormal
    Application.ActivateMicrosoftApp pjMicrosoftExcel
  Else
    MsgBox "No records found!", vbExclamation + vbOKOnly, "Completed Work"
  End If
    
exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  Set oAssignment = Nothing
  Application.StatusBar = ""
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Kill Environ("tmp") & "\Schema.ini"
  Kill Environ("tmp") & "\wp.csv"
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptCompletedWork", Err, Erl)
  Resume exit_here
End Sub

Sub cptFindUnstatusedTasks()
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  'strings
  Dim strMsg As String
  Dim strUnstatused As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngUnstatused As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  On Error Resume Next
  Set oTasks = ActiveProject.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oTasks Is Nothing Then
    MsgBox "This project has no tasks.", vbCritical + vbOKOnly, "No tasks"
    GoTo exit_here
  End If
  
  If oTasks.Count = 0 Then
    MsgBox "This project has no tasks.", vbCritical + vbOKOnly, "No tasks"
    GoTo exit_here
  End If
  
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Status Date is required.", vbCritical + vbOKOnly, "No Status Date"
    If Not Application.ChangeStatusDate Then
      GoTo exit_here
    Else
      dtStatus = ActiveProject.StatusDate
    End If
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  If Not IsDate(dtStatus) Then GoTo exit_here
  
  'Updating Task status updates resource status
  If Not ActiveProject.AutoTrack Then
    strMsg = "> Updating Task status updates resource status = True" & vbCrLf
  End If
  
  'Actual costs are always calculated by Project
  If Not ActiveProject.AutoCalcCosts Then
    strMsg = strMsg & "> Actual costs are always calculated by Project = True" & vbCrLf
  End If
  
  'prompt user to apply recommended settings
  If Len(strMsg) > 0 Then
    strMsg = "File > Options > Schedule > Calculation options for this project:" & vbCrLf & vbCrLf & strMsg & vbCrLf & "Apply now?"
    If MsgBox(strMsg, vbInformation + vbYesNo, "Recommended settings:") = vbYes Then
      ActiveProject.AutoTrack = True
      ActiveProject.AutoCalcCosts = True
    End If
  End If
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  FilterClear
  OptionsViewEx DisplaySummaryTasks:=True
  OutlineShowAllTasks
  
  lngTasks = oTasks.Count
  
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If oTask.Summary Then GoTo next_task 'skip summary tasks
    If oTask.ExternalTask Then GoTo next_task 'skip external tasks
    If Not oTask.Active Then GoTo next_task 'skip inactive tasks
    If oTask.Start < dtStatus And Not IsDate(oTask.ActualStart) Then 'unstarted
      strUnstatused = strUnstatused & oTask.UniqueID & vbTab
    ElseIf oTask.Finish <= dtStatus And Not IsDate(oTask.ActualFinish) Then 'unfinished
      strUnstatused = strUnstatused & oTask.UniqueID & vbTab
    ElseIf IsDate(oTask.ActualStart) And Not IsDate(oTask.ActualFinish) Then
      If oTask.Stop <> dtStatus Then  'unstatused
        strUnstatused = strUnstatused & oTask.UniqueID & vbTab
      End If
    End If
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Processing..." & Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & " (" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask
  'report results
  lngUnstatused = UBound(Split(strUnstatused, vbTab))
  If lngUnstatused > 0 Then
    strMsg = Format(lngUnstatused, "#,##0") & " unstatused task" & IIf(lngUnstatused = 1, ".", "s.") & vbCrLf & vbCrLf
    strMsg = strMsg & "Unstatused means:" & vbCrLf
    strMsg = strMsg & "> Forecast Start prior to Status Date" & vbCrLf
    strMsg = strMsg & "> Forecast Finish prior to Status Date" & vbCrLf
    strMsg = strMsg & "> In progress but not statused through Status Date"
    MsgBox strMsg, vbExclamation + vbOKCancel, "Unstatused Tasks"
    strUnstatused = Left(strUnstatused, Len(strUnstatused) - 1) 'hack off trailing tab
    OptionsViewEx DisplaySummaryTasks:=False
    SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strUnstatused
    SelectAll
    SetRowHeight "1"
    SelectBeginning
  Else
    MsgBox "No unstatused tasks.", vbInformation + vbOKOnly, "Well Done"
  End If

  Application.StatusBar = "Complete."

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oTasks = Nothing
  Set oTask = Nothing
  cptSpeed False
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_bas", "cptFindUnstatusedTasks", Err, Erl)
  Resume exit_here
End Sub
