Attribute VB_Name = "cptStatusSheetImport_bas"
'<cpt_version>v1.2.5</cpt_version>
Option Explicit

Sub cptShowStatusSheetImport_frm()
  'objects
  Dim rst As Object 'ADODB.Recordset
  'strings
  Dim strAppend As String
  Dim strTaskUsage As String
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
  Dim blnTaskUsageBelow As Boolean
  'variants
  Dim vField As Variant
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'todo: start Excel in the background if not already open; close on form close
  'todo: new options: log{true|false}; verbose{true|false}

  'populate comboboxes
  With cptStatusSheetImport_frm
    .Caption = "Import Status Sheets (" & cptGetVersion("cptStatusSheetImport_frm") & ")"
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
    'direct import to Actual Start removed in v1.3.0
    
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
    'direct import to Actual Finish removed in v1.3.0
    
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
        
    'convert legacy user settings
    strSettings = cptDir & "\settings\cpt-status-sheet-import.adtg"
    If Dir(strSettings) <> vbNullString Then
      'import user settings
      Set rst = CreateObject("ADODB.Recordset")
      rst.Open strSettings
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

    'import user settings
    strAS = cptGetSetting("StatusSheetImport", "cboAS")
    If Len(strAS) > 0 Then
      If strAS = CStr(FieldNameToFieldConstant("Actual Start")) Then
        MsgBox "Direct import to Actual Start is no longer supported. Please select a different field.", vbExclamation + vbOKOnly, "Actual Start"
      Else
        .cboAS.Value = CLng(strAS)
      End If
    End If
    strAF = cptGetSetting("StatusSheetImport", "cboAF")
    If Len(strAF) > 0 Then
      If strAF = CStr(FieldNameToFieldConstant("Actual Finish")) Then
        MsgBox "Direct import to Actual Finish is no longer supported. Please select a different field.", vbExclamation + vbOKOnly, "Actual Finish"
      Else
        .cboAF.Value = CLng(strAF)
      End If
    End If
    strFS = cptGetSetting("StatusSheetImport", "cboFS")
    If Len(strFS) > 0 Then .cboFS.Value = CLng(strFS)
    strFF = cptGetSetting("StatusSheetImport", "cboFF")
    If Len(strFF) > 0 Then .cboFF.Value = CLng(strFF)
    strEVP = cptGetSetting("StatusSheetImport", "cboEV")
    If Len(strEVP) > 0 Then .cboEV.Value = CLng(strEVP)
    strETC = cptGetSetting("StatusSheetImport", "cboETC")
    If Len(strETC) > 0 Then .cboETC.Value = CLng(strETC)
    strAppend = cptGetSetting("StatusSheetImport", "chkAppend")
    If Len(strAppend) > 0 Then .chkAppend = CBool(strAppend)
    strAppendTo = cptGetSetting("StatusSheetImport", "cboAppendTo")
    If Len(strAppendTo) > 0 Then .cboAppendTo.Value = strAppendTo
    
    'refresh which view
    strTaskUsage = cptGetSetting("StatusSheetImport", "optTaskUsage")
    If Len(strTaskUsage) > 0 Then
      If strTaskUsage = "above" Then
        .optAbove = True
        blnTaskUsageBelow = False
      ElseIf strTaskUsage = "below" Then
        .optBelow = True
        blnTaskUsageBelow = True
      End If
    Else
      .optBelow = True
      blnTaskUsageBelow = True
    End If
    .cmdRemove.Enabled = False
    'show the form
    .Show False
    cptRefreshStatusImportTable blnTaskUsageBelow

  End With
  
'  ActiveWindow.TopPane.Activate
'  If blnTaskUsageBelow Then
'    ViewApply "Task Entry"
'  Else
'    ViewApply "Task Usage"
'  End If
'  Call cptRefreshStatusImportTable(blnTaskUsageBelow)
  

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
  Dim oDict As Scripting.Dictionary
  Dim oShell As Object
  Dim oRecordset As ADODB.Recordset
  Dim oSubproject As Subproject
  Dim oTask As MSProject.Task
  Dim oResource As Resource
  Dim oAssignment As Assignment
  Dim oExcel As Object 'Excel.Application
  Dim oWorkbook As Object 'Excel.Workbook
  Dim oWorksheet As Object 'Excel.Worksheet
  Dim oListObject As Object 'Excel.ListObject
  Dim oRange As Object 'Excel.Range
  Dim oCell As Object 'Excel.Range
  Dim oComboBox As MSForms.ComboBox
  Dim rst As Object 'ADODB.Recordset
  'strings
  Dim strUIDList As String
  Dim strLOE As String
  Dim strLOEField As String
  Dim strHeader As String
  Dim strCon As String
  Dim strSQL As String
  Dim strDeconflictionFile As String
  Dim strSchema As String
  Dim strEVP As String
  Dim strFile As String
  Dim strNotesColTitle As String
  Dim strImportLog As String
  Dim strAppendTo As String
  Dim strSettings As String
  Dim strGUID As String
  'longs
  Dim lngEVT As Long
  Dim lngMultiplier As Long
  Dim lngDeconflictionFile As Long
  Dim lngEVP As Long
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
  Dim lngItem As Long
  Dim lngETC As Long
  Dim lngEV As Long
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngAF As Long
  Dim lngAS As Long
  'integers
  'doubles
  Dim dblWas As Double
  Dim dblETC As Double
  'booleans
  Dim blnAppend As Boolean
  Dim blnTask As Boolean
  Dim blnValid As Boolean
  'variants
  Dim vKey As Variant
  Dim vField As Variant
  Dim vControl As Variant
  'dates
  Dim dtNewDate As Date
  Dim dtStatus As Date

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: do we need to import actual and forecast dates separately? Why not simply 'New Start'?
  
  'validate choices for all
  With cptStatusSheetImport_frm
    
    blnValid = True
    
    'ensure file(s) are added to import list
    If .lboStatusSheets.ListCount = 0 Then
      MsgBox "Please select one or more files to import.", vbInformation + vbOKOnly, "No Files Found"
      blnValid = False
      GoTo exit_here
    End If
    
    'ensure import fields are selected
    For Each vControl In Array("cboAS", "cboAF", "cboFS", "cboFF", "cboEV", "cboETC", "cboAppendTo")
      'reset border color
      Set oComboBox = .Controls(vControl)
      oComboBox.BorderColor = -2147483642
      If IsNull(oComboBox) And oComboBox.Enabled Then
        oComboBox.BorderColor = 192 'red
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
    blnAppend = .chkAppend
    strAppendTo = .cboAppendTo
  End With
  
  'save user settings
  cptSaveSetting "StatusSheetImport", "cboAS", CStr(lngAS)
  cptSaveSetting "StatusSheetImport", "cboAF", CStr(lngAF)
  cptSaveSetting "StatusSheetImport", "cboFS", CStr(lngFS)
  cptSaveSetting "StatusSheetImport", "cboFF", CStr(lngFF)
  cptSaveSetting "StatusSheetImport", "cboEV", CStr(lngEV)
  cptSaveSetting "StatusSheetImport", "cboETC", CStr(lngETC)
  cptSaveSetting "StatusSheetImport", "chkAppend", IIf(blnAppend, 1, 0)
  cptSaveSetting "StatusSheetImport", "cboAppendTo", strAppendTo
  
  'ensure metrics settings exist
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "Settings not saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings Required"
      GoTo exit_here
    End If
  End If
  
  'get LOE settings
  strLOEField = cptGetSetting("Metrics", "cboLOEField")
  strLOE = cptGetSetting("Metrics", "txtLOE")
    
  'set up import log file
  If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'server project: default to Desktop
    Set oShell = CreateObject("WScript.Shell")
    strImportLog = oShell.SpecialFolders("Desktop") & "\cpt-import-log-" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".txt"
  Else  'not a server project: use ActiveProject.Path
    strImportLog = ActiveProject.Path & "\cpt-import-log-" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".txt"
  End If
  
  lngFile = FreeFile
  Open strImportLog For Output As #lngFile
  'log action
  Print #lngFile, "START STATUS SHEET IMPORT - " & Format(Now(), "mm/dd/yyyy hh:nn:ss")
  
  'set up deconfliction db
  strSchema = Environ("temp") & "\Schema.ini"
  lngDeconflictionFile = FreeFile
  Open strSchema For Output As lngDeconflictionFile
  Print #lngDeconflictionFile, "[imported.csv]"
  Print #lngDeconflictionFile, "Format=CSVDelimited"
  Print #lngDeconflictionFile, "ColNameHeaders=True"
  Print #lngDeconflictionFile, "Col1=FILE Text Width 255"
  Print #lngDeconflictionFile, "Col2=TASK_UID Integer"
  Print #lngDeconflictionFile, "Col3=FIELD Text Width 100"
  Print #lngDeconflictionFile, "Col4=RESOURCE_NAME Text Width 150"
  Print #lngDeconflictionFile, "Col5=WAS Text Width 50"
  Print #lngDeconflictionFile, "Col6=IS Text Width 50"
  Close #lngDeconflictionFile
  strDeconflictionFile = Environ("temp") & "\imported.csv"
  lngDeconflictionFile = FreeFile
  Open strDeconflictionFile For Output As #lngDeconflictionFile
  Print #lngDeconflictionFile, "FILE,TASK_UID,FIELD,RESOURCE_NAME,WAS,IS"
  
  'clear existing values from selected import fields -- but not oTask.ActualStart or oTask.ActualFinish
  cptStatusSheetImport_frm.lblStatus = "Clearing existing values..."
  cptSpeed True
  If ActiveProject.Subprojects.Count > 0 Then
    For Each oSubproject In ActiveProject.Subprojects
      lngTasks = lngTasks + oSubproject.SourceProject.Tasks.Count
    Next
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  
  For Each oTask In ActiveProject.Tasks
    lngTask = lngTask + 1
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    'clear dates
    For Each vField In Array(lngAS, lngAF, lngFS, lngFF)
      If vField = 188743721 Then GoTo next_field 'DO NOT clear out Actual Start
      If vField = 188743722 Then GoTo next_field 'DO NOT clear out Actual Finish
      If Not oTask.GetField(vField) = "NA" Then
        oTask.SetField vField, ""
      End If
next_field:
    Next vField
    'clear EV
    oTask.SetField lngEV, CStr(0)
    'clear ETC
    For Each oAssignment In oTask.Assignments
      If lngETC = pjTaskNumber1 Then
        oAssignment.Number1 = 0
        oTask.Number1 = 0
      ElseIf lngETC = pjTaskNumber2 Then
        oAssignment.Number2 = 0
        oTask.Number2 = 0
      ElseIf lngETC = pjTaskNumber3 Then
        oAssignment.Number3 = 0
        oTask.Number3 = 0
      ElseIf lngETC = pjTaskNumber4 Then
        oAssignment.Number4 = 0
        oTask.Number4 = 0
      ElseIf lngETC = pjTaskNumber5 Then
        oAssignment.Number5 = 0
        oTask.Number5 = 0
      ElseIf lngETC = pjTaskNumber6 Then
        oAssignment.Number6 = 0
        oTask.Number6 = 0
      ElseIf lngETC = pjTaskNumber7 Then
        oAssignment.Number7 = 0
        oTask.Number7 = 0
      ElseIf lngETC = pjTaskNumber8 Then
        oAssignment.Number8 = 0
        oTask.Number8 = 0
      ElseIf lngETC = pjTaskNumber9 Then
        oAssignment.Number9 = 0
        oTask.Number9 = 0
      ElseIf lngETC = pjTaskNumber10 Then
        oAssignment.Number10 = 0
        oTask.Number10 = 0
      ElseIf lngETC = pjTaskNumber11 Then
        oAssignment.Number11 = 0
        oTask.Number11 = 0
      ElseIf lngETC = pjTaskNumber12 Then
        oAssignment.Number12 = 0
        oTask.Number12 = 0
      ElseIf lngETC = pjTaskNumber13 Then
        oAssignment.Number13 = 0
        oTask.Number13 = 0
      ElseIf lngETC = pjTaskNumber14 Then
        oAssignment.Number14 = 0
        oTask.Number14 = 0
      ElseIf lngETC = pjTaskNumber15 Then
        oAssignment.Number15 = 0
        oTask.Number15 = 0
      ElseIf lngETC = pjTaskNumber16 Then
        oAssignment.Number16 = 0
        oTask.Number16 = 0
      ElseIf lngETC = pjTaskNumber17 Then
        oAssignment.Number17 = 0
        oTask.Number17 = 0
      ElseIf lngETC = pjTaskNumber18 Then
        oAssignment.Number18 = 0
        oTask.Number18 = 0
      ElseIf lngETC = pjTaskNumber19 Then
        oAssignment.Number19 = 0
        oTask.Number19 = 0
      ElseIf lngETC = pjTaskNumber20 Then
        oAssignment.Number20 = 0
        oTask.Number20 = 0
      End If
    Next oAssignment
next_task:
    cptStatusSheetImport_frm.lblStatus.Caption = "Clearing Previous Values...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptStatusSheetImport_frm.lblProgress.Width = (lngTask / lngTasks) * cptStatusSheetImport_frm.lblStatus.Width
    DoEvents
  Next oTask
  
  'set up array of updated  UIDs
  Set oDict = CreateObject("Scripting.Dictionary")
  
  'set up excel
  Set oExcel = CreateObject("Excel.Application")
  With cptStatusSheetImport_frm
    .lblStatus.Caption = "Importing..."
    For lngItem = 0 To .lboStatusSheets.ListCount - 1
      strFile = .lboStatusSheets.List(lngItem, 0) & .lboStatusSheets.List(lngItem, 1)
      Set oWorkbook = oExcel.Workbooks.Open(strFile, ReadOnly:=True)
      .lboStatusSheets.Selected(lngItem) = True
      DoEvents
      Print #lngFile, String(25, "=")
      Print #lngFile, "IMPORTING Workbook: " & strFile & " (" & oWorkbook.Sheets.Count & " Worksheets)"
      Print #lngFile, String(25, "-")
      For Each oWorksheet In oWorkbook.Sheets
        Print #lngFile, "IMPORTING Worksheet: " & oWorksheet.Name
'        cptStatusSheetImport_frm.lblStatus.Caption = "Importing...(" & Format(oWorksheet.Index / oWorkbook.Sheets.Count, "0%") & ")"
'        cptStatusSheetImport_frm.lblProgress.Width = (oWorksheet.Index / oWorkbook.Sheets.Count) * cptStatusSheetImport_frm.lblStatus.Width
        DoEvents
        
        'unhide columns and rows (sort is blocked by sheet protection...)
        oWorksheet.Columns.Hidden = False
        oWorksheet.Rows.Hidden = False
        
        'get status date
        On Error Resume Next
        dtStatus = oWorksheet.Range("STATUS_DATE")
        If Err.Number = 1004 Then 'invalid oWorkbook
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          Print #lngFile, "INVALID Worksheet - range 'STATUS_DATE' not found"
          GoTo next_worksheet
        End If
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        'get header row
        lngUIDCol = 1
        On Error Resume Next
        lngHeaderRow = oWorksheet.Columns(lngUIDCol).Find(what:="UID").Row
        If Err.Number = 1004 Then 'invalid oWorkbook
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          Print #lngFile, "INVALID Worksheet - UID column not found"
          GoTo next_worksheet
        End If
        'get header columns
        lngTaskNameCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Task Name", lookat:=xlPart).Column
        lngASCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Actual Start", lookat:=xlPart).Column
        lngAFCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Actual Finish", lookat:=xlPart).Column
        lngEVCol = oWorksheet.Rows(lngHeaderRow).Find(what:="New EV%", lookat:=xlWhole).Column
        lngETCCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Revised ETC", lookat:=xlWhole).Column
        strNotesColTitle = cptGetSetting("StatusSheet", "txtNotesColTitle")
        On Error Resume Next
        If Len(strNotesColTitle) > 0 Then
          lngCommentsCol = oWorksheet.Rows(lngHeaderRow).Find(what:=strNotesColTitle, lookat:=xlWhole).Column
        Else
          lngCommentsCol = oWorksheet.Rows(lngHeaderRow).Find(what:="Reason / Action / Impact", lookat:=xlWhole).Column
        End If
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If lngCommentsCol = 0 Then
          lngCommentsCol = oWorksheet.Cells(lngHeaderRow, 1).End(xlToRight).Column
        End If
        'get last row
        lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row
        'pull in the data
        For lngRow = lngHeaderRow + 1 To lngLastRow
          If oWorksheet.Cells(lngRow, lngUIDCol).Value = 0 Then GoTo next_row
          
          'determine if row is a oTask or an oAssignment
          On Error Resume Next
          Set oTask = ActiveProject.Tasks.UniqueID(oWorksheet.Cells(lngRow, lngUIDCol).Value)
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
          If oTask Is Nothing Or oTask.UniqueID <> oWorksheet.Cells(lngRow, lngUIDCol).Value Then 'check if it's a resource
            Set oResource = Nothing
            On Error Resume Next
            Set oResource = ActiveProject.Resources(oWorksheet.Cells(lngRow, lngTaskNameCol).Value)
            If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            If Not oResource Is Nothing Then
              blnTask = False
            Else
              Print #lngFile, "UID " & oWorksheet.Cells(lngRow, lngUIDCol) & " not found in IMS."
              If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
              GoTo next_row
            End If
          Else
            blnTask = True
          End If
          
          'set Task
          If blnTask Then
                    
            'skip completed tasks (which are also italicized)
            If IsDate(oTask.ActualFinish) Then
              If FormatDateTime(oTask.ActualFinish, vbShortDate) = FormatDateTime(oWorksheet.Cells(lngRow, lngAFCol).Value, vbShortDate) Then GoTo next_row
            End If

            'todo: do we really need to separate AS/FS on the form?
            'new start date
            If oWorksheet.Cells(lngRow, lngASCol).Value > 0 Then 'And Not oWorksheet.Cells(lngRow, lngASCol).Locked
              dtNewDate = FormatDateTime(CDate(oWorksheet.Cells(lngRow, lngASCol).Value), vbShortDate)
              'determine actual or forecast
              If dtNewDate <= FormatDateTime(dtStatus, vbShortDate) Then 'actual start
                If IsDate(oTask.ActualStart) Then
                  If FormatDateTime(oTask.ActualStart, vbShortDate) <> dtNewDate Then oTask.SetField lngAS, CDate(dtNewDate & " 08:00 AM")
                Else
                  oTask.SetField lngAS, CDate(dtNewDate & " 08:00 AM")
                End If
                Print #lngFile, "UID " & oTask.UniqueID & " AS > " & FormatDateTime(dtNewDate, vbShortDate)
                If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
              ElseIf dtNewDate > dtStatus Then 'forecast start
                If FormatDateTime(oTask.Start, vbShortDate) <> dtNewDate Then oTask.SetField lngFS, CDate(dtNewDate & " 08:00 AM")
                Print #lngFile, "UID " & oTask.UniqueID & " FS > " & FormatDateTime(dtNewDate, vbShortDate)
                If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
              End If
              If FormatDateTime(dtNewDate, vbShortDate) <> FormatDateTime(oTask.Start, vbShortDate) Then
                Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, "START", "", CStr(FormatDateTime(oTask.Start, vbShortDate)), CStr(FormatDateTime(dtNewDate, vbShortDate))), ",")
              End If
            End If
            'new finish date
            If oWorksheet.Cells(lngRow, lngAFCol).Value > 0 Then 'And Not oWorksheet.Cells(lngRow, lngAFCol).Locked
              dtNewDate = FormatDateTime(CDate(oWorksheet.Cells(lngRow, lngAFCol)), vbShortDate)
              If dtNewDate <= dtStatus Then 'actual finish
                If IsDate(oTask.ActualFinish) Then
                  If FormatDateTime(oTask.ActualFinish, vbShortDate) <> dtNewDate Then oTask.SetField lngAF, CDate(dtNewDate & " 05:00 PM")
                Else
                  oTask.SetField lngAF, CDate(dtNewDate & " 05:00 PM")
                End If
                Print #lngFile, "UID " & oTask.UniqueID & " AF > " & FormatDateTime(dtNewDate, vbShortDate)
                If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
              ElseIf dtNewDate > dtStatus Then 'forecast finish
                If FormatDateTime(oTask.Finish, vbShortDate) <> FormatDateTime(dtNewDate, vbShortDate) Then
                  If dtNewDate > #12/31/2149# Then
                    MsgBox "Workbook: " & oWorkbook.Name & vbCrLf & "UID: " & oTask.UniqueID & vbCrLf & "Forecast Finish: " & FormatDateTime(dtNewDate, vbShortDate) & vbCrLf & vbCrLf & "...skipped.", vbCritical + vbOKOnly, "Invalid Date"
                    Print #lngFile, "UID " & oTask.UniqueID & " FF > ERROR (" & FormatDateTime(dtNewDate, vbShortDate) & " is outside allowable range)"
                  Else
                    oTask.SetField lngFF, CDate(dtNewDate & " 05:00 PM")
                    Print #lngFile, "UID " & oTask.UniqueID & " FF > " & FormatDateTime(dtNewDate, vbShortDate)
                    If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
                  End If
                End If
              End If
              If FormatDateTime(dtNewDate, vbShortDate) <> FormatDateTime(oTask.Finish, vbShortDate) Then
                Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, "FINISH", "", CStr(FormatDateTime(oTask.Finish, vbShortDate)), CStr(FormatDateTime(dtNewDate, vbShortDate))), ",")
              End If
            End If
            
            'evp
            'skip LOE
            If Len(strLOEField) > 0 And Len(strLOE) > 0 Then
              lngEVT = CLng(strLOEField)
              If oTask.GetField(lngEVT) = strLOE Then GoTo evp_skipped
            End If
            'secondary catch to skip LOE
            If oWorksheet.Cells(lngRow, lngEVCol).Value <> "-" Then
              lngEVP = Round(oWorksheet.Cells(lngRow, lngEVCol).Value * 100, 0)
              strEVP = cptGetSetting("StatusSheet", "cboEVP")
              If Len(strEVP) > 0 Then 'compare
                If CLng(cptRegEx(oTask.GetField(FieldNameToFieldConstant(strEVP)), "[0-9]{1,}")) <> lngEVP Then
                  oTask.SetField lngEV, lngEVP
                  Print #lngFile, "UID " & oTask.UniqueID & " EV% > " & lngEVP & "%"
                  If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
                  Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, strEVP, "", cptRegEx(oTask.GetField(FieldNameToFieldConstant(strEVP)), "[0-9]{1,}"), CStr(lngEVP)), ",")
                End If
              Else 'log
                oTask.SetField lngEV, lngEVP
                Print #lngFile, "UID " & oTask.UniqueID & " EV% > " & lngEVP & "%"
                If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
                Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, "EV%", "", "<unknown>", CStr(lngEVP)), ",")
              End If
            End If
            
evp_skipped:
            'comments todo: only import if different
            If .chkAppend And oWorksheet.Cells(lngRow, lngCommentsCol).Value <> "" Then
              If .cboAppendTo = "Top of Task Note" Then
                oTask.Notes = Format(dtStatus, "mm/dd/yyyy") & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf & String(25, "-") & vbCrLf & vbCrLf & oTask.Notes
              'todo: replace task note
              ElseIf .cboAppendTo = "Overwrite Note" Then
                oTask.Notes = Format(dtStatus, "mm/dd/yyyy") & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
              ElseIf .cboAppendTo = "Bottom of Task Note" Then
                oTask.AppendNotes vbCrLf & String(25, "-") & vbCrLf & Format(dtStatus, "mm/dd/yyyy") & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
              End If
            End If
          ElseIf Not blnTask Then 'it's an Assignment
            On Error Resume Next
            Set oAssignment = oTask.Assignments.UniqueID(oWorksheet.Cells(lngRow, lngUIDCol).Value)
            If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
            If oAssignment Is Nothing Then
              Print #lngFile, "ASSIGNMENT MISSING: TASK " & oTask.UniqueID & " ASSIGNMENT: " & oWorksheet.Cells(lngRow, lngUIDCol).Value
            Else
              'If Not oWorksheet.Cells(lngRow, lngETCCol).Locked Then
                dblETC = oWorksheet.Cells(lngRow, lngETCCol).Value 'get the new value
                dblWas = 0 'reset was
                If oAssignment.ResourceType = pjResourceTypeWork Then
                  dblWas = Val(oAssignment.RemainingWork) / 60
                Else
                  dblWas = Val(oAssignment.RemainingCost)
                End If
                'only import if updated
                If Round(dblWas, 2) <> Round(dblETC, 2) Then
                  If lngETC = pjTaskNumber1 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number1 = dblETC
                    oTask.Number1 = oTask.Number1 + dblETC
                  ElseIf lngETC = pjTaskNumber2 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number2 = dblETC
                    oTask.Number2 = oTask.Number2 + dblETC
                  ElseIf lngETC = pjTaskNumber3 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number3 = dblETC
                    oTask.Number3 = oTask.Number3 + dblETC
                  ElseIf lngETC = pjTaskNumber4 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number4 = dblETC
                    oTask.Number4 = oTask.Number4 + dblETC
                  ElseIf lngETC = pjTaskNumber5 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number5 = dblETC
                    oTask.Number5 = oTask.Number5 + dblETC
                  ElseIf lngETC = pjTaskNumber6 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number6 = dblETC
                    oTask.Number6 = oTask.Number6 + dblETC
                  ElseIf lngETC = pjTaskNumber7 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number7 = dblETC
                    oTask.Number7 = oTask.Number7 + dblETC
                  ElseIf lngETC = pjTaskNumber8 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number8 = dblETC
                    oTask.Number8 = oTask.Number8 + dblETC
                  ElseIf lngETC = pjTaskNumber9 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number9 = dblETC
                    oTask.Number9 = oTask.Number9 + dblETC
                  ElseIf lngETC = pjTaskNumber10 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number10 = dblETC
                    oTask.Number10 = oTask.Number10 + dblETC
                  ElseIf lngETC = pjTaskNumber11 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number11 = dblETC
                    oTask.Number11 = oTask.Number11 + dblETC
                  ElseIf lngETC = pjTaskNumber12 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number12 = dblETC
                    oTask.Number12 = oTask.Number12 + dblETC
                  ElseIf lngETC = pjTaskNumber13 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number13 = dblETC
                    oTask.Number13 = oTask.Number13 + dblETC
                  ElseIf lngETC = pjTaskNumber14 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number14 = dblETC
                    oTask.Number14 = oTask.Number14 + dblETC
                  ElseIf lngETC = pjTaskNumber15 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number15 = dblETC
                    oTask.Number15 = oTask.Number15 + dblETC
                  ElseIf lngETC = pjTaskNumber16 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number16 = dblETC
                    oTask.Number16 = oTask.Number16 + dblETC
                  ElseIf lngETC = pjTaskNumber17 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number17 = dblETC
                    oTask.Number17 = oTask.Number17 + dblETC
                  ElseIf lngETC = pjTaskNumber18 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number18 = dblETC
                    oTask.Number18 = oTask.Number18 + dblETC
                  ElseIf lngETC = pjTaskNumber19 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number19 = dblETC
                    oTask.Number19 = oTask.Number19 + dblETC
                  ElseIf lngETC = pjTaskNumber20 Then
                    Print #lngDeconflictionFile, Join(Array(strFile, oTask.UniqueID, FieldConstantToFieldName(lngETC), oAssignment.ResourceName, dblWas, dblETC), ",")
                    oAssignment.Number20 = dblETC
                    oTask.Number20 = oTask.Number20 + dblETC
                  End If
                  Print #lngFile, "UID " & oTask.UniqueID & " [" & oAssignment.Resource.Name & "] ETC > " & dblETC
                  If Not oDict.Exists(oTask.UniqueID) Then oDict.Add oTask.UniqueID, oTask.UniqueID
                End If
                If .chkAppend And Len(oWorksheet.Cells(lngRow, lngCommentsCol)) > 0 Then
                  If .cboAppendTo = "Top of Task Note" Then
                    oAssignment.Notes = Format(dtStatus, "mm/dd/yyyy") & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf & String(25, "-") & vbCrLf & vbCrLf & oAssignment.Notes
                  'todo: replace assignment note
                  ElseIf .cboAppendTo = "Overwrite Note" Then
                    oAssignment.Notes = Format(dtStatus, "mm/dd/yyyy") & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
                  ElseIf .cboAppendTo = "Bottom of Task Note" Then
                    oAssignment.AppendNotes vbCrLf & String(25, "-") & vbCrLf & Format(dtStatus, "mm/dd/yyyy") & " - " & oWorksheet.Cells(lngRow, lngCommentsCol) & vbCrLf
                  End If
                End If
              'End If
              'todo: consolidate Assignment Notes into Task Notes?
              Set oAssignment = Nothing
            End If
          End If
next_row:
          cptStatusSheetImport_frm.lblStatus.Caption = "Importing...(" & Format(lngRow / lngLastRow, "0%") & ")"
          DoEvents
        Next lngRow
next_worksheet:
        Print #lngFile, String(25, "-")
      Next oWorksheet
next_file:
      .lblStatus.Caption = "Importing...(" & lngItem + 1 & " of " & .lboStatusSheets.ListCount & ")"
      .lboStatusSheets.Selected(lngItem) = False
      oWorkbook.Close False
      DoEvents
    Next lngItem
  End With 'cptStatusSheetImport_frm
  
  'were there any conflicts?
  Close #lngDeconflictionFile
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("temp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  strSQL = "SELECT T1.TASK_UID,T1.RESOURCE_NAME,T1.FIELD,T2.WAS,T2.[IS],T2.FILE "
  strSQL = strSQL & "FROM ((SELECT TASK_UID,RESOURCE_NAME,FIELD,COUNT(FILE) FROM [imported.csv] GROUP BY TASK_UID,RESOURCE_NAME,FIELD HAVING COUNT(FILE)>1) AS T1) "
  strSQL = strSQL & "LEFT JOIN [imported.csv] AS T2 ON T2.TASK_UID=T1.TASK_UID AND T2.FIELD=T1.FIELD  " 'AND T2.RESOURCE_NAME=T1.RESOURCE_NAME
  strSQL = strSQL & "ORDER BY T1.TASK_UID,T1.FIELD" 'todo: refine query
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  If oRecordset.RecordCount > 0 Then
    Print #lngFile, ">>> " & oRecordset.RecordCount & " POTENTIAL CONFLICTS IDENTIFIED <<<"
    If MsgBox("Potential conflicts found!" & vbCrLf & vbCrLf & "Review in Excel?", vbExclamation + vbYesNo, "Please Review") = vbYes Then
      oExcel.Visible = True
      Set oWorkbook = oExcel.Workbooks.Add
      Set oWorksheet = oWorkbook.Sheets(1)
      For lngItem = 1 To oRecordset.Fields.Count
        oWorksheet.Cells(1, lngItem).Value = oRecordset.Fields(lngItem - 1).Name
      Next lngItem
      oWorksheet.[A2].CopyFromRecordset oRecordset
      oExcel.ActiveWindow.Zoom = 85
      oWorksheet.Columns.AutoFit
    Else
      For lngItem = 0 To oRecordset.Fields.Count - 1
        strHeader = strHeader & oRecordset.Fields(lngItem).Name & ","
      Next lngItem
      Print #lngFile, strHeader
      Print #lngFile, oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      Print #lngFile, "...conflicts not reviewed."
    End If
  End If
  
  'get the string of UIDs updated
  If oDict.Count > 0 Then
    For Each vKey In oDict.Keys
      strUIDList = strUIDList & vKey & ","
    Next vKey
    strUIDList = Left(strUIDList, Len(strUIDList) - 1)
  End If
  
exit_here:
  On Error Resume Next
  Set oDict = Nothing
  Set oShell = Nothing
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Set oSubproject = Nothing
  cptStatusSheetImport_frm.lblStatus.Caption = "Import Complete."
  cptStatusSheetImport_frm.lblProgress.Width = cptStatusSheetImport_frm.lblStatus.Width
  DoEvents
  If blnValid Then
    'close log for output
    Print #lngFile, String(25, "=")
    Print #lngFile, "IMPORT COMPLETE." & vbCrLf
    Print #lngFile, "UPDATED UIDs:"
    If Len(strUIDList) = 0 Then
      Print #lngFile, "< no updates >"
    Else
      Print #lngFile, strUIDList
    End If
    Close #lngFile
    'open log in notepad
    Shell "notepad.exe """ & strImportLog & """", vbNormalFocus
  End If
  cptStatusSheetImport_frm.lblStatus.Caption = "Ready..."
  cptSpeed False
  Set oAssignment = Nothing
  Set oResource = Nothing
  Set oTask = Nothing
  For lngFile = 1 To FreeFile
    Close #lngFile
  Next lngFile
  Close #lngDeconflictionFile
  If Dir(Environ("tmp") & "\Schema.ini") <> vbNullString Then Kill Environ("tmp") & "\Schema.ini"
  If Dir(Environ("tmp") & "\imported.csv") <> vbNullString Then Kill Environ("tmp") & "\imported.csv"
  Set oRange = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  'If Not oWorkbook Is Nothing Then oWorkbook.Close False
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oComboBox = Nothing
  If Not rst Is Nothing Then
    If rst.State = 1 Then rst.Close
  End If
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_bas", "cptStatusSheetImport", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshStatusImportTable(Optional blnUsageBelow As Boolean = False)
  'objects
  Dim rst As Object 'ADODB.Recordset 'Object
  'strings
  Dim strBottomPaneViewName As String
  Dim strEVT As String
  Dim strEVP As String
  Dim strSettings As String
  'longs
  Dim lngEVT As Long
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not cptStatusSheetImport_frm.Visible Then GoTo exit_here

  cptSpeed True
  
  'get saved settings
  'get EVP and EVT
  strSettings = cptDir & "\settings\cpt-status-sheet.adtg" 'todo: keep for a few more versions
  If Dir(strSettings) <> vbNullString Then
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open strSettings
    If Not rst.EOF Then
      'todo: does field name still match?
      strEVP = rst("cboEVP")
      strEVT = rst("cboEVT")
    End If
    rst.Close
    'convert to ini
    cptSaveSetting "StatusSheet", "cboEVP", strEVP
    cptSaveSetting "StatusSheet", "cboEVT", strEVT
    'todo: don't kill the file here, kill it on Status Sheet Creation
  End If
  
  strEVP = cptGetSetting("StatusSheet", "cboEVP")
  strEVT = cptGetSetting("StatusSheet", "cboEVT")
  
  'reset the table
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  
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
          TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=rst(1), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
        End If
        rst.MoveNext
      Loop
    End If
    rst.Close
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Name", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Total Slack", Title:="", Width:=8, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Actual Start", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(cptStatusSheetImport_frm.cboAS.Value) Then
    lngAS = cptStatusSheetImport_frm.cboAS.Value
    If lngAS <> FieldNameToFieldConstant("Actual Start") Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngAS), Title:="New Actual Start", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Start", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(cptStatusSheetImport_frm.cboFS.Value) Then
    lngFS = cptStatusSheetImport_frm.cboFS.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngFS), Title:="New Forecast Start", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Actual Finish", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  If Not IsNull(cptStatusSheetImport_frm.cboAF.Value) Then
    lngAF = cptStatusSheetImport_frm.cboAF.Value
    If lngAF <> FieldNameToFieldConstant("Actual Finish") Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngAF), Title:="New Actual Finish", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Remaining Duration", Title:="", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Finish", Title:="", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  If Not IsNull(cptStatusSheetImport_frm.cboFF.Value) Then
    lngFF = cptStatusSheetImport_frm.cboFF.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngFF), Title:="New Forecast Finish", Width:=14, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  'EVT
  If Len(strEVT) > 0 Then
    On Error Resume Next
    lngEVT = FieldNameToFieldConstant(strEVT)
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If lngEVT > 0 Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=strEVT, Title:="EVT", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If
  'existing EV%
  If Len(strEVP) > 0 Then
    'does field still exist?
    On Error Resume Next
    lngEVP = FieldNameToFieldConstant(strEVP)
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If lngEVP > 0 Then
      TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=strEVP, Title:="EV%", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
  End If



  'imported EV
  If Not IsNull(cptStatusSheetImport_frm.cboEV.Value) Then
    lngNewEVP = cptStatusSheetImport_frm.cboEV.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngNewEVP), Title:="New EV%", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  'keep these here so user can filter on changes above, make edits below
  'Type
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Type", Width:=17, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'Effort Driven
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Effort Driven", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'existing ETC (remaining work)
  TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  'imported ETC
  If Not IsNull(cptStatusSheetImport_frm.cboETC.Value) Then
    lngETC = cptStatusSheetImport_frm.cboETC.Value
    TableEditEx Name:="cptStatusSheetImport Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngETC), Title:="New ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
  End If
  
  If blnUsageBelow Then
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, Create:=True, OverwriteExisting:=True, FieldName:="ID", Title:="", Width:=10, Align:=1, ShowInMenu:=False, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Unique ID", Title:="UID", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    
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
            TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:=rst(1), Title:="", Width:=10, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
          End If
          rst.MoveNext
        Loop
      End If
      rst.Close
    End If
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Name", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'Type
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Type", Width:=17, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'Effort Driven
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Effort Driven", Width:=10, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'existing ETC (remaining work)
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Remaining Work", Title:="ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    'imported ETC
    If Not IsNull(cptStatusSheetImport_frm.cboETC.Value) Then
      lngETC = cptStatusSheetImport_frm.cboETC.Value
      TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:=FieldConstantToFieldName(lngETC), Title:="New ETC", Width:=15, Align:=1, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    End If
    TableEditEx Name:="cptStatusSheetImportDetails Table", TaskTable:=True, NewFieldName:="Notes", Title:="", Width:=60, Align:=0, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, AlignTitle:=1, HeaderAutoRowHeightAdjustment:=False, WrapText:=False
    ActiveWindow.TopPane.Activate
    'If ActiveProject.CurrentView <> "Gantt Chart" Then ViewApply Name:="Gantt Chart"
    ViewApply Name:="Gantt Chart"
    'If ActiveProject.CurrentTable <> "cptStatusSheetImport Table" Then TableApply Name:="cptStatusSheetImport Table"
    TableApply Name:="cptStatusSheetImport Table"
    'todo: reapply group?
    
    On Error Resume Next
    strBottomPaneViewName = ActiveWindow.BottomPane.View.Name
    If Err.Number = 91 Then
      Err.Clear
      Application.FormViewShow
      'Application.ToggleTaskDetails
    End If
    ActiveWindow.BottomPane.Activate
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    'If ActiveProject.CurrentView <> "Task Usage" Then ViewApply "Task Usage"
    ViewApply "Task Usage"
    'If ActiveProject.CurrentTable <> "cptStatusSheetImportDetails Table" Then TableApply "cptStatusSheetImportDetails Table"
    TableApply "cptStatusSheetImportDetails Table"
    ActiveWindow.TopPane.Activate
  Else
    ActiveWindow.TopPane.Activate
    'If ActiveProject.CurrentView <> "Task Usage" Then ViewApply "Task Usage"
    ViewApply "Task Usage"
    DoEvents
    'If ActiveProject.CurrentTable <> "cptStatusSheetImport Table" Then TableApply Name:="cptStatusSheetImport Table"
    TableApply Name:="cptStatusSheetImport Table"
    On Error Resume Next
    strBottomPaneViewName = ActiveWindow.BottomPane.View.Name
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Len(strBottomPaneViewName) > 0 Then
      DetailsPaneToggle
    End If
    'todo: reapply group?
  End If
  
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

