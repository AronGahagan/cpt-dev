Attribute VB_Name = "cptStatusSheetImport_bas"
'<cpt_version>v1.1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowStatusSheetImport_frm()
'objects
Dim rst As Object 'ADODB.Recordset
'strings
Dim strAppendTo As String
Dim strETC As String
Dim strEVP As String
Dim strFF As String
Dim strFS As String
Dim strAF As String
Dim strAS As String
Dim strGUID As String
Dim strSettings As String
Dim strCustomFieldName As String
'longs
Dim lngField As Long
'integers
Dim intField As Integer
'doubles
'booleans
'variants
Dim vField As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'populate comboboxes
  With cptStatusSheetImport_frm
    .cboAppendTo.Clear
    .cboAppendTo.AddItem "Bottom of Task Note"
    .cboAppendTo.AddItem "Top of Task Note"
    
    'start
    For Each vField In Array("Start", "Date")
      For intField = 1 To 10
        lngField = FieldNameToFieldConstant(vField & intField, pjTask)
        strCustomFieldName = CustomFieldGetName(lngField)
        .cboAS.AddItem
        .cboAS.List(.cboAS.ListCount - 1, 0) = lngField
        .cboAS.List(.cboAS.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
        .cboFS.AddItem
        .cboFS.List(.cboFS.ListCount - 1, 0) = lngField
        .cboFS.List(.cboFS.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
      Next intField
    Next vField
    'add AS to bottom of list
    .cboAS.AddItem
    .cboAS.List(.cboAS.ListCount - 1, 0) = FieldNameToFieldConstant("Actual Start")
    .cboAS.List(.cboAS.ListCount - 1, 1) = "Actual Start"
    
    'finish
    For Each vField In Array("Finish", "Date")
      For intField = 1 To 10
        lngField = FieldNameToFieldConstant(vField & intField, pjTask)
        strCustomFieldName = CustomFieldGetName(lngField)
        .cboAF.AddItem
        .cboAF.List(.cboAF.ListCount - 1, 0) = lngField
        .cboAF.List(.cboAF.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
        .cboFF.AddItem
        .cboFF.List(.cboFF.ListCount - 1, 0) = lngField
        .cboFF.List(.cboFF.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
      Next intField
    Next vField
    'add AF to bottom of list
    .cboAF.AddItem
    .cboAF.List(.cboAF.ListCount - 1, 0) = FieldNameToFieldConstant("Actual Finish")
    .cboAF.List(.cboAF.ListCount - 1, 1) = "Actual Finish"
    
    'ev% and etc
    For Each vField In Array("Number")
      For intField = 1 To 20
        lngField = FieldNameToFieldConstant(vField & intField, pjTask)
        strCustomFieldName = CustomFieldGetName(lngField)
        .cboEV.AddItem
        .cboEV.List(.cboEV.ListCount - 1, 0) = lngField
        .cboEV.List(.cboEV.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
        .cboETC.AddItem
        .cboETC.List(.cboETC.ListCount - 1, 0) = lngField
        .cboETC.List(.cboETC.ListCount - 1, 1) = FieldConstantToFieldName(lngField) & IIf(Len(strCustomFieldName) > 0, " (" & strCustomFieldName & ")", "")
      Next intField
    Next vField
    
    'todo: add enterprise custom fields?
    
    'get project guid
    If Application.Version < 12 Then
      strGUID = ActiveProject.DatabaseProjectUniqueID
    Else
      strGUID = ActiveProject.GetServerProjectGuid
    End If
    
    'convert legacy user settings
    strSettings = cptDir & "\settings\cpt-status-sheet-import.adtg"
    If Dir(strSettings) <> vbNullString Then
      'import user settings
      Set rst = CreateObject("ADODB.Recordset")
      rst.Open strSettings
      'rst.Filter = "GUID='" & strGUID & "'"
      If Not rst.EOF Then
        cptSaveSetting "StatusSheetImport", "cboAS", CStr(rst("AS"))
        cptSaveSetting "StatusSheetImport", "cboAF", CStr(rst("AF"))
        cptSaveSetting "StatusSheetImport", "cboFS", CStr(rst("FS"))
        cptSaveSetting "StatusSheetImport", "cboFF", CStr(rst("FF"))
        cptSaveSetting "StatusSheetImport", "cboEV", CStr(rst("EV"))
        cptSaveSetting "StatusSheetImport", "cboETC", CStr(rst("ETC"))
        cptSaveSetting "StatusSheetImport", "chkAppend", CStr(rst("Append"))
        If rst("AppendTo") <> "" Then
          cptSaveSetting "StatusSheetImport", "cboAppendTo", CStr(rst("AppendTo"))
        Else
          cptSaveSetting "StatusSheetImport", "cboAppendTo", "Top of Task Note"
        End If
      End If
      Kill strSettings
    Else
      'default settings
      .cboAppendTo.Value = "Top of Task Note"
    End If

    'todo: import user settings
    strAS = cptGetSetting("StatusSheetImport", "cboAS")
    If Len(strAS) > 0 Then .cboAS.Value = CLng(strAS)
    strAF = cptGetSetting("StatusSheetImport", "cboAF")
    If Len(strAF) > 0 Then .cboAF.Value = CLng(strAF)
    strFS = cptGetSetting("StatusSheetImport", "cboFS")
    If Len(strFS) > 0 Then .cboFS.Value = CLng(strFS)
    strFF = cptGetSetting("StatusSheetImport", "cboFF")
    If Len(strFF) > 0 Then .cboFF.Value = CLng(strFF)
    strEVP = cptGetSetting("StatusSheetImport", "cboEV")
    If Len(strEVP) > 0 Then .cboEV.Value = CLng(strEVP)
    strETC = cptGetSetting("StatusSheetImport", "cboETC")
    If Len(strETC) > 0 Then .cboETC.Value = CLng(strETC)
    .chkAppend = CBool(cptGetSetting("StatusSheetImport", "chkAppend"))
    strAppendTo = cptGetSetting("StatusSheetImport", "cboAppendTo")
    If Len(strAppendTo) > 0 Then
      .cboAppendTo.Value = strAppendTo
    Else
      .cboAppendTo.Value = "Top of Task Note"
    End If
  
    'show the form
    .Show False

  End With
  
  ActiveWindow.TopPane.Activate
  ViewApply "Task Usage"
  Call cptRefreshStatusImportTable

exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptShowStatusSheetImport_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptStatusSheetImport()
'objects
Dim SubProject As Object
Dim Task As Task
Dim Resource As Resource
Dim Assignment As Assignment
Dim xlApp As Object 'Excel.Application
Dim Workbook As Object 'Workbook
Dim Worksheet As Object 'Worksheet
Dim ListObject As Object 'ListObject
Dim rng As Object 'Excel.Range
Dim c As Object 'Excel.Range
Dim cbo As ComboBox
Dim rst As Object 'ADODB.Recordset
'strings
Dim strImportLog As String
Dim strAppendTo As String
Dim strSettings As String
Dim strGUID As String
'longs
Dim lngTask As Long
Dim lngTasks As Long
Dim lngTaskNameCol As Long
Dim lngEVCol As Long
Dim lngUIDCol As Long
Dim lngFile As Long
Dim lngRow As Long
Dim lngCommentsCol As Long
Dim lngETCCol As Long
Dim lngAFCol As Long
Dim lngASCol As Long
Dim lngHeaderRow As Long
Dim lngLastRow As Long
Dim lngFiles As Long
Dim lngETC As Long
Dim lngEV As Long
Dim lngFF As Long
Dim lngFS As Long
Dim lngAF As Long
Dim lngAS As Long
'integers
'doubles
'booleans
Dim blnTask As Boolean
Dim blnValid As Boolean
'variants
Dim vField As Variant
Dim vControl As Variant
'dates
Dim dtNewDate As Date
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: save current dates to selected fields
  'todo: age the dates X periods - carry back names
  'todo: view should be gantt on top and task usage below (but with custom table having only UID,{user fields},Task/Resource Name, Remaining Work, New ETC
  
  'validate choices for all
  With cptStatusSheetImport_frm
    
    blnValid = True
    
    'ensure file(s) are added to import list
    If .TreeView1.Nodes.Count = 0 Then
      MsgBox "Please select one or more files to import.", vbInformation + vbOKOnly, "No Files Found"
      blnValid = False
      GoTo exit_here
    End If
    
    'ensure import fields are selected
    For Each vControl In Array("cboAS", "cboAF", "cboFS", "cboFF", "cboEV", "cboETC", "cboAppendTo")
      'reset border color
      Set cbo = .Controls(vControl)
      cbo.BorderColor = -2147483642
      If IsNull(cbo) And cbo.Enabled Then
        cbo.BorderColor = 192 'red
        blnValid = False
      End If
    Next vControl
    
    'warn if invalid
    If Not blnValid Then
      MsgBox "Please select import fields.", vbExclamation + vbOKOnly, "Invalid Import Fields"
      GoTo exit_here
    End If
    
  End With
  
  'speed up
  cptSpeed True
  
  'capture import fields and settings
  With cptStatusSheetImport_frm
    lngAS = .cboAS.Value
    lngAF = .cboAF.Value
    lngFS = .cboFS.Value
    lngFF = .cboFF.Value
    lngEV = .cboEV.Value
    lngETC = .cboETC.Value
    strAppendTo = .cboAppendTo
  End With
  
  'save user settings
  strSettings = cptDir & "\settings\cpt-status-sheet-import.adtg"
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  Set rst = CreateObject("ADODB.Recordset")
  
  If Dir(strSettings) <> vbNullString Then
    'update the settings if different
    rst.Open strSettings
    rst.Find "GUID='" & strGUID & "'"
    If rst.EOF Then
      If Application.Version < 12 Then
        strGUID = ActiveProject.DatabaseProjectUniqueID
      Else
        strGUID = ActiveProject.GetServerProjectGuid
      End If
      rst.AddNew
      rst("GUID") = strGUID
    End If
  Else
    'create it
    With rst
      .Fields.Append "GUID", adGUID
      .Fields.Append "AS", adInteger
      .Fields.Append "AF", adInteger
      .Fields.Append "FS", adInteger
      .Fields.Append "FF", adInteger
      .Fields.Append "EV", adInteger
      .Fields.Append "ETC", adInteger
      .Fields.Append "Append", adBoolean
      .Fields.Append "AppendTo", adVarChar, 50
      .Open
      .AddNew Array("GUID"), Array(strGUID)
    End With
  End If
  rst.Fields("AS") = lngAS
  rst.Fields("AF") = lngAF
  rst.Fields("FS") = lngFS
  rst.Fields("FF") = lngFF
  rst.Fields("EV") = lngEV
  rst.Fields("ETC") = lngETC
  rst.Fields("Append") = IIf(cptStatusSheetImport_frm.chkAppend, -1, 0)
  If cptStatusSheetImport_frm.chkAppend Then
    If IsNull(cptStatusSheetImport_frm.cboAppendTo) Then strAppendTo = "Top of Task Note"
    rst.Fields("AppendTo") = strAppendTo
  End If
  rst.Update
  rst.Save strSettings, adPersistADTG
  rst.Close
  
  'set up import log file
  strImportLog = Environ("USERPROFILE") & "\CP_Status_Sheets\cpt-import-log-" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".txt"
  lngFile = FreeFile
  Open strImportLog For Output As #lngFile

  'log action
  Print #lngFile, "START STATUS SHEET IMPORT - " & Format(Now(), "mm/dd/yyyy hh:nn:ss")

  'clear existing values from selected import fields -- but not Task.ActualStart or Task.ActualFinish
  cptStatusSheetImport_frm.lblStatus = "Clearing existing values..."
  cptSpeed True
  If ActiveProject.Subprojects.Count > 0 Then
    For Each SubProject In ActiveProject.Subprojects
      lngTasks = lngTasks + SubProject.SourceProject.Tasks.Count
    Next
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  For Each Task In ActiveProject.Tasks
    lngTask = lngTask + 1
    If Task Is Nothing Then GoTo next_task
    If Task.Summary Then GoTo next_task
    If Task.ExternalTask Then GoTo next_task
    If Not Task.Active Then GoTo next_task
    'clear dates
    For Each vField In Array(lngAS, lngAF, lngFS, lngFF)
      If vField = 188743721 Then GoTo next_field 'DO NOT clear out Actual Start
      If vField = 188743722 Then GoTo next_field 'DO NOT clear out Actual Finish
      If Not Task.GetField(vField) = "NA" Then
        Task.SetField vField, ""
      End If
next_field:
    Next vField
    'clear EV
    Task.SetField lngEV, CStr(0)
    'clear ETC
    For Each Assignment In Task.Assignments
      If lngETC = FieldNameToFieldConstant("Number1") Then Assignment.Number1 = 0
      If lngETC = FieldNameToFieldConstant("Number2") Then Assignment.Number2 = 0
      If lngETC = FieldNameToFieldConstant("Number3") Then Assignment.Number3 = 0
      If lngETC = FieldNameToFieldConstant("Number4") Then Assignment.Number4 = 0
      If lngETC = FieldNameToFieldConstant("Number5") Then Assignment.Number5 = 0
      If lngETC = FieldNameToFieldConstant("Number6") Then Assignment.Number6 = 0
      If lngETC = FieldNameToFieldConstant("Number7") Then Assignment.Number7 = 0
      If lngETC = FieldNameToFieldConstant("Number8") Then Assignment.Number8 = 0
      If lngETC = FieldNameToFieldConstant("Number9") Then Assignment.Number9 = 0
      If lngETC = FieldNameToFieldConstant("Number10") Then Assignment.Number10 = 0
      If lngETC = FieldNameToFieldConstant("Number11") Then Assignment.Number11 = 0
      If lngETC = FieldNameToFieldConstant("Number12") Then Assignment.Number12 = 0
      If lngETC = FieldNameToFieldConstant("Number13") Then Assignment.Number13 = 0
      If lngETC = FieldNameToFieldConstant("Number14") Then Assignment.Number14 = 0
      If lngETC = FieldNameToFieldConstant("Number15") Then Assignment.Number15 = 0
      If lngETC = FieldNameToFieldConstant("Number16") Then Assignment.Number16 = 0
      If lngETC = FieldNameToFieldConstant("Number17") Then Assignment.Number17 = 0
      If lngETC = FieldNameToFieldConstant("Number18") Then Assignment.Number18 = 0
      If lngETC = FieldNameToFieldConstant("Number19") Then Assignment.Number19 = 0
      If lngETC = FieldNameToFieldConstant("Number20") Then Assignment.Number20 = 0
    Next Assignment
next_task:
    cptStatusSheetImport_frm.lblStatus.Caption = "Clearing Previous Values...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptStatusSheetImport_frm.lblProgress.Width = (lngTask / lngTasks) * cptStatusSheetImport_frm.lblStatus.Width
    DoEvents
  Next Task
    
  'set up excel
  Set xlApp = CreateObject("Excel.Application")
  With cptStatusSheetImport_frm
    For lngFiles = 1 To .TreeView1.Nodes.Count
      Set Workbook = xlApp.Workbooks.Open(.TreeView1.Nodes(lngFiles).Text, ReadOnly:=True)
      cptStatusSheetImport_frm.lblStatus.Caption = "Importing " & Workbook.Name & "..."
      DoEvents
      Print #lngFile, String(25, "=")
      Print #lngFile, "IMPORTING WORKBOOK: " & .TreeView1.Nodes(lngFiles).Text & " (" & Workbook.Sheets.Count & " worksheets)"
      Print #lngFile, String(25, "-")
      For Each Worksheet In Workbook.Sheets
        Print #lngFile, "IMPORTING WORKSHEET: " & Worksheet.Name
        cptStatusSheetImport_frm.lblStatus.Caption = "Importing Worksheets...(" & Format(Worksheet.Index / Workbook.Sheets.Count, "0%") & ")"
        cptStatusSheetImport_frm.lblProgress.Width = (Worksheet.Index / Workbook.Sheets.Count) * cptStatusSheetImport_frm.lblStatus.Width
        DoEvents
        'get status date
        On Error Resume Next
        dtStatus = Worksheet.Range("STATUS_DATE")
        If Err.Number = 1004 Then 'invalid workbook
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          Print #lngFile, "INVALID WORKSHEET - UID HEADER NOT FOUND IN COLUMN 1 OF WORKSHEET"
          GoTo next_worksheet
        End If
        If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
        'get header row
        lngUIDCol = 1
        lngHeaderRow = Worksheet.Columns(lngUIDCol).Find(what:="UID").Row
        'get header columns
        lngTaskNameCol = Worksheet.Rows(lngHeaderRow).Find(what:="Task Name", lookat:=xlWhole).Column
        lngASCol = Worksheet.Rows(lngHeaderRow).Find(what:="Actual Start", lookat:=xlPart).Column
        lngAFCol = Worksheet.Rows(lngHeaderRow).Find(what:="Actual Finish", lookat:=xlPart).Column
        lngEVCol = Worksheet.Rows(lngHeaderRow).Find(what:="New EV%", lookat:=xlWhole).Column
        lngETCCol = Worksheet.Rows(lngHeaderRow).Find(what:="Revised ETC", lookat:=xlWhole).Column
        lngCommentsCol = Worksheet.Rows(lngHeaderRow).Find(what:="Reason / Action / Impact", lookat:=xlWhole).Column
        'get last row
        lngLastRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Row
        'pull in the data
        For lngRow = lngHeaderRow + 1 To lngLastRow
          'only summaries have interior color
          If Worksheet.Cells(lngRow, lngUIDCol).Interior.Color <> 16777215 Then GoTo next_row
          'determine if row is a task or an assignment
          If Worksheet.Cells(lngRow, lngUIDCol).Font.Color = 0 Then
            blnTask = False
          ElseIf Worksheet.Cells(lngRow, lngUIDCol).Font.Color = 6567712 Then
            blnTask = True
          Else
            GoTo next_row
          End If
          'set task
          On Error Resume Next
          Set Task = ActiveProject.Tasks.UniqueID(Worksheet.Cells(lngRow, lngUIDCol))
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          If Task Is Nothing Then
            Print #lngFile, "UID " & Worksheet.Cells(lngRow, lngUIDCol) & " not found in IMS."
            GoTo next_row
          End If
          If blnTask Then
            'new start date
            If Worksheet.Cells(lngRow, lngASCol).Value > 0 And Not Worksheet.Cells(lngRow, lngASCol).Locked Then
              dtNewDate = CDate(Worksheet.Cells(lngRow, lngASCol).Value)
              'determine actual or forecast
              If dtNewDate <= dtStatus Then 'actual start
                Task.SetField lngAS, dtNewDate
              ElseIf dtNewDate > dtStatus Then 'forecast start
                Task.SetField lngFS, dtNewDate
              End If
            End If
            'new finish date
            If Worksheet.Cells(lngRow, lngAFCol).Value > 0 And Not Worksheet.Cells(lngRow, lngAFCol).Locked Then
              dtNewDate = CDate(Worksheet.Cells(lngRow, lngAFCol))
              If dtNewDate <= dtStatus Then 'actual finish
                Task.SetField lngAF, CDate(Worksheet.Cells(lngRow, lngAFCol))
              ElseIf dtNewDate > dtStatus Then 'forecast finish
                Task.SetField lngFF, CDate(Worksheet.Cells(lngRow, lngAFCol))
              End If
            End If
            'ev
            If (Worksheet.Cells(lngRow, lngEVCol) * 100) <> Task.GetField(lngEV) Then
              Task.SetField lngEV, Worksheet.Cells(lngRow, lngEVCol) * 100
            End If
            'comments
            If .chkAppend And Worksheet.Cells(lngRow, lngCommentsCol).Value <> "" Then
              If .cboAppendTo = "Top of Task Note" Then
                Task.Notes = Format(Now, "mm/dd/yyyy") & " - " & Worksheet.Cells(lngRow, lngCommentsCol) & vbCrLf & String(25, "-") & vbCrLf & Task.Notes
              ElseIf .cboAppendTo = "Bottom of Task Note" Then
                Task.AppendNotes String(25, "-") & vbCrLf & Format(Now, "mm/dd/yyyy") & " - " & Worksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
              End If
            End If
            'if user is importing to AS and AF then mark task on track
'            'todo: decide if this is a good idea
'            If lngAS = FieldNameToFieldConstant("Actual Start") And lngAF = FieldNameToFieldConstant("Actual Finish") Then
'              'todo: what if user has it filtered? or collapsed?
'              EditGoTo Task.ID
'              UpdateProject All:=False, UpdateDate:=CStr(dtStatus & " 5 PM"), Action:=1
'            End If
          ElseIf Not blnTask Then 'it's an assignment
            On Error Resume Next
            Set Assignment = Task.Assignments.UniqueID(Worksheet.Cells(lngRow, lngUIDCol))
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            If Assignment Is Nothing Then
              Print #lngFile, "ASSIGNMENT MISSING: " & Worksheet.Cells(lngRow, lngTaskNameCol).Value
            Else
              If lngETC = FieldNameToFieldConstant("Number1") Then Assignment.Number1 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number2") Then Assignment.Number2 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number3") Then Assignment.Number3 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number4") Then Assignment.Number4 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number5") Then Assignment.Number5 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number6") Then Assignment.Number6 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number7") Then Assignment.Number7 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number8") Then Assignment.Number8 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number9") Then Assignment.Number9 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number10") Then Assignment.Number10 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number11") Then Assignment.Number11 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number12") Then Assignment.Number12 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number13") Then Assignment.Number13 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number14") Then Assignment.Number14 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number15") Then Assignment.Number15 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number16") Then Assignment.Number16 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number17") Then Assignment.Number17 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number18") Then Assignment.Number18 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number19") Then Assignment.Number19 = Worksheet.Cells(lngRow, lngETCCol)
              If lngETC = FieldNameToFieldConstant("Number20") Then Assignment.Number20 = Worksheet.Cells(lngRow, lngETCCol)
              Set Assignment = Nothing
            End If
          End If
next_row:
        Next lngRow
next_worksheet:
        Print #lngFile, String(25, "-")
      Next Worksheet
next_file:
      Workbook.Close False
    Next lngFiles
  End With 'cptStatusSheetImport_frm
  
  'reset view
  ActiveWindow.TopPane.Activate
  ViewApply "Task Usage"
  Call cptRefreshStatusImportTable
  
exit_here:
  On Error Resume Next
  Set SubProject = Nothing
  cptStatusSheetImport_frm.lblStatus.Caption = "Import Complete."
  cptStatusSheetImport_frm.lblProgress.Width = cptStatusSheetImport_frm.lblStatus.Width
  DoEvents
  If blnValid Then
    'close log for output
    Print #lngFile, String(25, "=")
    Print #lngFile, "IMPORT COMPLETE."
    Close #lngFile
    'open log in notepad
    Shell "C:\WINDOWS\notepad.exe " & strImportLog, vbNormalFocus
  End If
  cptStatusSheetImport_frm.lblStatus.Caption = "Ready..."
  cptSpeed False
  Set Assignment = Nothing
  Set Resource = Nothing
  Set Task = Nothing
  For lngFile = 1 To FreeFile
    Close #lngFile
  Next lngFile
  Set c = Nothing
  Set rng = Nothing
  Set ListObject = Nothing
  Set Worksheet = Nothing
  If Not Workbook Is Nothing Then Workbook.Close False
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set cbo = Nothing
  If rst.State = 1 Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptStatusSheetImport", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshStatusImportTable()
'objects
Dim rst As Object 'ADODB.Recordset 'Object
'strings
Dim strEVP As String
Dim strSettings As String
Dim strGUID As String
'longs
Dim lngETC As Long
Dim lngEVP As Long
Dim lngNewEVP As Long
Dim lngFF As Long
Dim lngAF As Long
Dim lngFS As Long
Dim lngAS As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not cptStatusSheetImport_frm.Visible Then GoTo exit_here

  cptSpeed True
  
  'get saved settings
  'get project guid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'get EVP
  strSettings = cptDir & "\settings\cpt-status-sheet.adtg"
  If Dir(strSettings) <> vbNullString Then
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open strSettings
    If Not rst.EOF Then
      'does field name still match?
      strEVP = rst("cboEVP")
    End If
    rst.Close
  End If
  
  'reset the table
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  
  'import user fields
  strSettings = cptDir & "\settings\cpt-status-sheet-userfields.adtg"
  If Dir(strSettings) <> vbNullString Then
    'import user settings
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open strSettings
    If Not rst.EOF Then
      rst.MoveFirst
      Do While Not rst.EOF
        'does field name still match?
        If CustomFieldGetName(rst(0)) = rst(1) Then
          TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=rst(1), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
        End If
        rst.MoveNext
      Loop
    End If
    rst.Close
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Name", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Remaining Duration", Title:="", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Start", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Actual Start", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(cptStatusSheetImport_frm.cboAS.Value) Then
    lngAS = cptStatusSheetImport_frm.cboAS.Value
    If lngAS <> FieldNameToFieldConstant("Actual Start") Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(lngAS), Title:="New Actual Start", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  If Not IsNull(cptStatusSheetImport_frm.cboFS.Value) Then
    lngFS = cptStatusSheetImport_frm.cboFS.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(lngFS), Title:="New Forecast Start", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Finish", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Actual Finish", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(cptStatusSheetImport_frm.cboAF.Value) Then
    lngAF = cptStatusSheetImport_frm.cboAF.Value
    If lngAF <> FieldNameToFieldConstant("Actual Finish") Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(lngAF), Title:="New Actual Finish", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  If Not IsNull(cptStatusSheetImport_frm.cboFF.Value) Then
    lngFF = cptStatusSheetImport_frm.cboFF.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(lngFF), Title:="New Forecast Finish", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  'existing EV%
  If Len(strEVP) > 0 Then
    'does field still exist?
    On Error Resume Next
    lngEVP = FieldNameToFieldConstant(strEVP)
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If lngEVP > 0 Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=strEVP, Title:="EV%", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  'imported EV
  If Not IsNull(cptStatusSheetImport_frm.cboEV.Value) Then
    lngNewEVP = cptStatusSheetImport_frm.cboEV.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(lngNewEVP), Title:="New EV%", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  'existing ETC (remaining work)
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:="Remaining Work", Title:="ETC", Width:=20, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'imported ETC
  If Not IsNull(cptStatusSheetImport_frm.cboETC.Value) Then
    lngETC = cptStatusSheetImport_frm.cboETC.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, newfieldname:=FieldConstantToFieldName(lngETC), Title:="New ETC", Width:=20, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  ActiveWindow.TopPane.Activate
  TableApply Name:="Entry"
  TableApply Name:="cptStatusSheetImport Table"

  'reset the filter
'  FilterEdit Name:="cptStatusSheetImport Filter", Taskfilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Actual Finish", test:="equals", Value:="NA", ShowInMenu:=False, ShowSummaryTasks:=True
'  If cptStatusSheetImport_frm.chkHide And IsDate(cptStatusSheetImport_frm.txtHideCompleteBefore) Then
'    FilterEdit Name:="cptStatusSheetImport Filter", Taskfilter:=True, FieldName:="", newfieldname:="Actual Finish", test:="is greater than or equal to", Value:=cptStatusSheetImport_frm.txtHideCompleteBefore, operation:="Or", ShowSummaryTasks:=True
'  End If
'  FilterApply "cptStatusSheetImport Filter"

exit_here:
  On Error Resume Next
  If rst.State = 1 Then rst.Close
  Set rst = Nothing
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptRefreshStatusImportTable", Err, Erl)
  Err.Clear
  Resume exit_here
End Sub

