Attribute VB_Name = "cptDECM_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit

Private strWBS As String
Private strOBS As String
Private strCA As String
Private strCAM As String
Private strWP As String
Private strEVT As String
Private strEVP As String
Private strPass As String
Private strFail As String

Private lngWBS As Long
Private lngOBS As Long
Private lngCA As Long
Private lngCAM As Long
Private lngWP As Long
Private lngEVT As Long
Private lngEVP As Long

Function ValidMap() As Boolean
  'objects
  Dim oComboBox As ComboBox
  'strings
  Dim strSetting As String
  'longs
  Dim lngItem  As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  Dim blnValid As Boolean
  'variants
  Dim vAddField  As Variant
  Dim vFields As Variant
  Dim vControl As Variant
  'dates
  
  'todo: validate cptIntegration_frm.cboEVP
  'todo: validate cptIntegration_frm.cboEOC
  'todo: validate cptIntegration_frm.txtRollingWaveDate
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnValid = True
  
  With cptIntegration_frm
    
    .Caption = "Integration (" & cptGetVersion("cptIntegration_frm") & ")"
    
    For Each vControl In Split("WBS,OBS,CA,CAM,WP,WPM,EVT,LOE,EVP,EOC", ",")
      If vControl = "WBS" Then vControl = "CWBS" 'todo: fix saved setting name
      If vControl = "WP" Then vControl = "WPCN"  'todo: fix saved setting name
      strSetting = cptGetSetting("Integration", CStr(vControl))
      If Len(strSetting) = 0 Then
        If vControl = "EVP" Then
          strSetting = cptGetSetting("Metrics", "cboEVP") & "|" & FieldConstantToFieldName(strSetting)
          cptSaveSetting "Integration", "EVP", strSetting
        ElseIf vControl = "EVT" Then
          strSetting = cptGetSetting("Metrics", "cboLOEField") & "|" & FieldConstantToFieldName(strSetting)
          cptSaveSetting "Integration", "EVT", strSetting
        ElseIf vControl = "LOE" Then
          strSetting = cptGetSetting("Metrics", "txtLOE")
          cptSaveSetting "Integration", "LOE", strSetting
          .cboLOE.Value = strSetting
        End If
      End If
      If vControl = "CWBS" Then vControl = "WBS"  'todo: fix saved setting name
      If vControl = "WPCN" Then vControl = "WP"   'todo: fix saved setting name
      Set oComboBox = .Controls("cbo" & vControl)
      oComboBox.BorderColor = -2147483642
      If Len(strSetting) = 0 Then
        blnValid = False
        lngField = 0
        oComboBox.BorderColor = 192
      Else
        If vControl <> "LOE" Then
          lngField = CLng(Split(strSetting, "|")(0))
        Else
          Dim strLOE As String
          strLOE = strSetting
        End If
      End If
      If vControl = "WBS" Then
        oComboBox.List = cptGetCustomFields("t", "Outline Code,Text", "c,cfn", False)
        If IsEmpty(oComboBox.List(oComboBox.ListCount - 1, 0)) Then oComboBox.RemoveItem (oComboBox.ListCount - 1)
      ElseIf vControl = "CAM" Or vControl = "WPM" Then
        For Each vAddField In Split("Contact", ",")
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant(vAddField)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vAddField
        Next vAddField
        vFields = cptGetCustomFields("t", "Text,Outline Code", "c,cfn", False)
        For lngItem = 0 To UBound(vFields) - 1
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
        Next lngItem
      ElseIf vControl = "EVP" Then
        For Each vAddField In Split("Physical % Complete,% Complete", ",")
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant(vAddField)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vAddField
        Next vAddField
        vFields = cptGetCustomFields("t", "Number", "c,cfn", False)
        For lngItem = 0 To UBound(vFields) - 1
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
        Next lngItem
      ElseIf vControl = "EOC" Then
        For Each vAddField In Split("Code,Group,Initials,Type", ",")
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = FieldNameToFieldConstant(vAddField, pjResource)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vAddField
        Next vAddField
        vFields = cptGetCustomFields("r", "Text", "c,cfn", False)
        For lngItem = 0 To UBound(vFields) - 1
          oComboBox.AddItem
          oComboBox.List(oComboBox.ListCount - 1, 0) = vFields(lngItem, 0)
          oComboBox.List(oComboBox.ListCount - 1, 1) = vFields(lngItem, 1)
        Next lngItem
      ElseIf vControl = "LOE" Then
        .cboLOE.Value = strLOE
        GoTo next_control
      Else 'WP
        oComboBox.List = cptGetCustomFields("t", "Text,Outline Code", "c,cfn", False)
        If IsEmpty(oComboBox.List(oComboBox.ListCount - 1, 0)) Then oComboBox.RemoveItem (oComboBox.ListCount - 1)
      End If
      If lngField > 0 Then oComboBox.Value = lngField
next_control:
    Next vControl
    
    .txtFiscalCalendar.BorderColor = 192
    If cptCalendarExists("cptFiscalCalendar") Then
      .txtFiscalCalendar.Value = "cptFiscalCalendar"
      .txtFiscalCalendar.BorderColor = -2147483642
    End If
    
    'todo: rolling wave date
    
    .Show
    'todo: validate selections
    'todo: save setting and update cbo border after selection
    ValidMap = .blnValidIntegrationMap
    'todo: save/overwrite new settings
  End With

exit_here:
  On Error Resume Next
  Set oComboBox = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptDECM", "ValidMap", Err, Erl)
  Resume exit_here
    
End Function

Sub cptDECM_GET_DATA()
'Optional blnIncompleteOnly As Boolean = True, Optional blnDiscreteOnly As Boolean = True
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oDict As Scripting.Dictionary
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRecordset As ADODB.Recordset
  Dim oLink As MSProject.TaskDependency
  Dim oTask As MSProject.Task
  'strings
  Dim strLinks  As String
  Dim strPE As String
  Dim strSE As String
  Dim strRecord As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  Dim strLOE As String
  Dim strList As String
  'longs
  Dim lngAssignmentFile As Long
  Dim lngTS As Long
  Dim lngConst As Long
  Dim lngX As Long
  Dim lngY As Long
  Dim lngSummary As Long
  Dim lngDur As Long
  Dim lngBDur As Long
  Dim lngAF As Long
  Dim lngAS As Long
  Dim lngBLF As Long
  Dim lngBLS As Long
  Dim lngFF As Long
  Dim lngFS As Long
  Dim lngUID As Long
  Dim lngTask As Long
  Dim lngLinkFile As Long
  Dim lngTaskFile As Long
  Dim lngWBS As Long
  Dim lngOBS As Long
  Dim lngCA As Long
  Dim lngCAM As Long
  Dim lngWP As Long
  Dim lngWPM As Long
  Dim lngEVT As Long
  Dim lngEVP As Long
  Dim lngFile As Long
  Dim lngTasks As Long
  Dim lngItem As Long
  'integers
  'doubles
  Dim dblScore As Double
  'booleans
  Dim blnDumpToExcel As Boolean
  'variants
  Dim vHeader As Variant
  Dim vField As Variant
  'dates
  Dim dtStatus As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Pleave provide a Status Date.", vbExclamation + vbOKOnly, "Status Date Required"
    ChangeStatusDate
    If Not IsDate(ActiveProject.StatusDate) Then
      MsgBox "Status Date is required. Exiting.", vbCritical + vbOKOnly, "No Status Date"
      GoTo exit_here
    End If
  End If
  
  dtStatus = ActiveProject.StatusDate 'todo: what dates does GetField return? times?
  
  If Not ValidMap Then GoTo exit_here
  
  cptSpeed True
  
  lngFile = FreeFile
  strDir = Environ("tmp")
  strFile = strDir & "\Schema.ini"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngFile
  Print #lngFile, "[tasks.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=UID integer"
  Print #lngFile, "Col2=WBS text"
  Print #lngFile, "Col3=OBS text"
  Print #lngFile, "Col4=CA text"
  Print #lngFile, "Col5=CAM text"
  Print #lngFile, "Col6=WP text"
  Print #lngFile, "Col7=WPM text"
  Print #lngFile, "Col8=EVT text"
  Print #lngFile, "Col9=EVP integer"
  Print #lngFile, "Col10=FS date"
  Print #lngFile, "Col11=FF date"
  Print #lngFile, "Col12=BLS date"
  Print #lngFile, "Col13=BLF date"
  Print #lngFile, "Col14=AS date"
  Print #lngFile, "Col15=AF date"
  Print #lngFile, "Col16=BDUR integer"
  Print #lngFile, "Col17=DUR integer"
  Print #lngFile, "Col18=SUMMARY text" 'Yes/No
  Print #lngFile, "Col19=CONST text"
  Print #lngFile, "Col20=TS integer"
  Print #lngFile, "[links.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=FROM integer"
  Print #lngFile, "Col2=TO integer"
  Print #lngFile, "Col3=TYPE text"
  Print #lngFile, "Col4=LAG integer"
  Print #lngFile, "Col5=LAG_TYPE integer"
  Print #lngFile, "[assignments.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=TASK_UID integer"
  Print #lngFile, "Col2=RESOURCE_UID integer"
  Print #lngFile, "Col3=BLW Double"
  Print #lngFile, "Col4=BLC Double"
  Print #lngFile, "Col5=RW Double"
  Print #lngFile, "Col6=RC Double"
  Print #lngFile, "Col7=EOC text"
  Close #lngFile
  
  lngTaskFile = FreeFile
  strFile = strDir & "\tasks.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngTaskFile
  
  lngLinkFile = FreeFile
  strFile = strDir & "\links.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngLinkFile
  
  lngAssignmentFile = FreeFile
  strFile = strDir & "\assignments.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngAssignmentFile
  
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  lngTasks = ActiveProject.Tasks.Count
  
  'get settings
  lngUID = FieldNameToFieldConstant("Unique ID")
  lngWBS = CLng(Split(cptGetSetting("Integration", "CWBS"), "|")(0))
  lngOBS = CLng(Split(cptGetSetting("Integration", "OBS"), "|")(0))
  lngCA = CLng(Split(cptGetSetting("Integration", "CA"), "|")(0))
  lngCAM = CLng(Split(cptGetSetting("Integration", "CAM"), "|")(0))
  lngWP = CLng(Split(cptGetSetting("Integration", "WPCN"), "|")(0))
  lngWPM = CLng(Split(cptGetSetting("Integration", "WPM"), "|")(0))
  lngEVT = CLng(Split(cptGetSetting("Integration", "EVT"), "|")(0))
  strLOE = cptGetSetting("Integration", "LOE")
  lngEVP = CLng(Split(cptGetSetting("Integration", "EVP"), "|")(0))
  'todo: clean EVP: remove %; normalize to whole number?
  lngFS = FieldNameToFieldConstant("Start")
  lngFF = FieldNameToFieldConstant("Finish")
  lngBLS = FieldNameToFieldConstant("Baseline Start")
  lngBLF = FieldNameToFieldConstant("Baseline Finish")
  lngAS = FieldNameToFieldConstant("Actual Start")
  lngAF = FieldNameToFieldConstant("Actual Finish")
  lngBDur = FieldNameToFieldConstant("Baseline Duration")
  lngDur = FieldNameToFieldConstant("Duration")
  lngSummary = FieldNameToFieldConstant("Summary")
  lngConst = FieldNameToFieldConstant("Constraint Type")
  lngTS = FieldNameToFieldConstant("Total Slack")
  
  'headers
  Print #lngTaskFile, "UID,WBS,OBS,CA,CAM,WP,WPM,EVT,EVP,FS,FF,BLS,BLF,AS,AF,BDUR,DUR,SUMMARY,CONST,TS,"
  Print #lngLinkFile, "FROM,TO,TYPE,LAG,"
  Print #lngAssignmentFile, "TASK_UID,RESOURCE_UID,BLW,BLC,RW,RC,EOC,"
  
  cptDECM_frm.Caption = "DECM v5.0 (cpt " & cptGetVersion("cptDECM_bas") & ")"
  lngItem = 0
  cptDECM_frm.lboHeader.Clear
  cptDECM_frm.lboHeader.AddItem
  For Each vHeader In Split("METRIC,TITLE,TARGET,X,Y,SCORE,ICON,DESCRIPTION,TBD", ",")
    cptDECM_frm.lboHeader.List(0, lngItem) = vHeader
    lngItem = lngItem + 1
  Next vHeader
  cptDECM_frm.lboMetrics.Clear
  cptDECM_frm.Show False
  
  blnDumpToExcel = False 'todo: if DumpToExcel then single Workbook, do better with tab names
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    'If oTask.Summary Then GoTo next_task
    'todo: external?
'    If blnIncompleteOnly Then If IsDate(oTask.ActualFinish) Then GoTo next_task
'    If blnDiscreteOnly Then If oTask.GetField(lngEVT) = "A" Then GoTo next_task 'todo: what else is non-discrete? apportioned?
    For Each vField In Array(lngUID, lngWBS, lngOBS, lngCA, lngCAM, lngWP, lngWPM, lngEVT, lngEVP, lngFS, lngFF, lngBLS, lngBLF, lngAS, lngAF, lngBDur, lngDur, lngSummary, lngConst, lngTS)
      If vField = FieldNameToFieldConstant("Physical % Complete") Then
        strRecord = strRecord & cptRegEx(oTask.GetField(vField), "[0-9]{1,}") & ","
      ElseIf vField = FieldNameToFieldConstant("% Complete") Then
        strRecord = strRecord & cptRegEx(oTask.GetField(vField), "[0-9]{1,}") & ","
      ElseIf vField = lngBDur Then
        strRecord = strRecord & oTask.BaselineDuration & ","
      ElseIf vField = lngDur Then
        strRecord = strRecord & oTask.Duration & ","
      ElseIf vField = lngConst Then
        strRecord = strRecord & Choose(oTask.ConstraintType + 1, "ASAP", "ALAP", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & ","
      ElseIf vField = lngTS Then
        strRecord = strRecord & oTask.TotalSlack & "," 'todo: convert to days?
      Else
        strRecord = strRecord & oTask.GetField(CLng(vField)) & ","
      End If
    Next vField
    Print #lngTaskFile, strRecord
    For Each oLink In oTask.TaskDependencies
      'todo: convert lag to effective days
      Print #lngLinkFile, oLink.From & "," & oLink.To & "," & Choose(oLink.Type + 1, "FF", "FS", "SF", "SS") & "," & oLink.Lag & ","
    Next oLink
    For Each oAssignment In oTask.Assignments
      Print #lngAssignmentFile, Join(Array(oTask.UniqueID, oAssignment.ResourceUniqueID, oAssignment.BaselineWork, oAssignment.BaselineCost, oAssignment.RemainingWork, oAssignment.RemainingCost, oAssignment.Resource.GetField(Split(cptGetSetting("Integration", "EOC"), "|")(0))), ",")
    Next
next_task:
    strRecord = ""
    lngTask = lngTask + 1
    Application.StatusBar = "Loading Data...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptDECM_frm.lblStatus.Caption = "Loading Data...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptDECM_frm.lblProgress.Width = (lngTask / lngTasks) * cptDECM_frm.lblStatus.Width
    DoEvents
  Next oTask
  
  Close #lngTaskFile
  Close #lngLinkFile
  Close #lngAssignmentFile
  
  cptDECM_frm.lblStatus.Caption = "Loading...done."
  Application.StatusBar = "Loading...done."
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  
  'lboMetrics: METRIC,TITLE,THRESHOLD,X,Y,SCORE,DESCRIPTION,?sql
  strPass = "[+]"
  strFail = "<!>"
  
  '===== EVMS =====
  '05A101a - 1 CA : 1 OBS
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 05A101a..."
  Application.StatusBar = "Getting EVMS: 05A101a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "05A101a"
  'cptDECM_frm.lboMetrics.Value = "05A101a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "1 CA : 1 OBS"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = Count of CAs with more than one OBS element or no OBS elements assigned
  'Y = Total count of CAs
  'X/Y = 0%
  strSQL = "SELECT DISTINCT CA FROM tasks.csv WHERE CA IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    .Close
  End With
  strSQL = "SELECT CA,COUNT(OBS) AS CountOfOBS "
  strSQL = strSQL & "FROM (SELECT DISTINCT CA,OBS FROM [tasks.csv]) "
  strSQL = strSQL & "WHERE CA IS NOT NULL "
  strSQL = strSQL & "GROUP BY CA "
  strSQL = strSQL & "HAVING COUNT(OBS)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("CA") & "," 'todo: UID is not in the query
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 05A101a...done."
  Application.StatusBar = "Getting EVMS: 05A101a...done."
  DoEvents
  
  '05A102a - 1 CA : 1 CAM
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 05A102a..."
  Application.StatusBar = "Getting EVMS: 05A102a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "05A102a"
  'cptDECM_frm.lboMetrics.Value = "05A102a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "1 CA : 1 CAM"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of CAs that have more than one CAM or no CAM assigned
  'Y = Total count of CAs
  'X/Y <= 5%
  'we already have lngY...
  strSQL = "SELECT CA,COUNT(CAM) AS CountOfCAM "
  strSQL = strSQL & "FROM (SELECT DISTINCT CA,CAM FROM [tasks.csv]) "
  strSQL = strSQL & "WHERE CA IS NOT NULL "
  strSQL = strSQL & "GROUP BY CA "
  strSQL = strSQL & "HAVING COUNT(CAM)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("CA") & "," 'todo: fix this - UID is not in the query
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 05A102a...done."
  Application.StatusBar = "Getting EVMS: 05A102a...done."
  DoEvents
  
  '05A103a - 1 CA : 1 WBS
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 05A103a..."
  Application.StatusBar = "Getting EVMS: 05A103a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "05A103a"
  'cptDECM_frm.lboMetrics.Value = "05A103a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "1 CA : 1 WBS"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = Count of CAs with more than one WBS element or no WBS elements assigned
  'Y = Total count of CAs
  'X/Y = 0%
  'we already have lngY...
  strSQL = "SELECT CA,COUNT(WBS) AS CountOfWBS "
  strSQL = strSQL & "FROM (SELECT DISTINCT CA,WBS FROM [tasks.csv]) "
  strSQL = strSQL & "WHERE CA IS NOT NULL "
  strSQL = strSQL & "GROUP BY CA "
  strSQL = strSQL & "HAVING COUNT(WBS)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("CA") & "," 'todo: fix this - UID is not in the query
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 05A103a...done."
  Application.StatusBar = "Getting EVMS: 05A103a...done."
  DoEvents
  
  '10A102a - 1 WP : 1 EVT
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A102a..."
  Application.StatusBar = "Getting EVMS: 05A103a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A102a"
  'cptDECM_frm.lboMetrics.Value = "10A102a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "1 WP : 1 EVT"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = count of incomplete WPs that have more than one EVT or no EVT assigned
  'Y = count of incomplete WPs
  'X/Y <= 5%
  strSQL = "SELECT WP,COUNT(EVT) AS CountOfEVT "
  strSQL = strSQL & "FROM (SELECT DISTINCT WP,EVT FROM [tasks.csv] WHERE WP IS NOT NULL AND AF IS NULL) "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING COUNT(EVT)>1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("WP") & "," 'todo: fix this - UID is not in the query
        .MoveNext
      Loop
    End If
    If blnDumpToExcel Then DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT DISTINCT WP "
  strSQL = strSQL & "FROM [tasks.csv] "
  strSQL = strSQL & "WHERE WP IS NOT NULL AND AF IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore < 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A102a...done."
  Application.StatusBar = "Getting EVMS: 10A102a...done."
  DoEvents
  
  '10A109b - all WPs have budget
  cptDECM_frm.lblStatus.Caption = "Getting EVMS Metric: 10A109b..."
  Application.StatusBar = "Getting EVMS Metric: 10A109b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A109b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "WPs With Budgets"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of WPs/PPs/SLPPs with BAC ? 0
  'Y = Total count of WPs/PPs/SLPPs
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] WHERE WP IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = .RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT t.WP,SUM(a.BLW) AS [BLW],SUM(a.BLC) AS [BLC] FROM [tasks.csv] t "
  strSQL = strSQL & "INNER JOIN [assignments.csv] a on a.TASK_UID=t.UID "
  strSQL = strSQL & "WHERE t.WP IS NOT NULL "
  strSQL = strSQL & "GROUP BY t.WP "
  strSQL = strSQL & "HAVING SUM(a.BLW)=0 AND SUM(a.BLC)=0"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("WP") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore < 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A109b...done."
  Application.StatusBar = "Getting EVMS: 10A109b...done."
  DoEvents
  
  '10A202a - mixed EOC
  cptDECM_frm.lblStatus.Caption = "Getting EVMS Metric: 10A202a..."
  Application.StatusBar = "Getting EVMS Metric: 10A109b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A202a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "WPs w/mixed EOCs"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%?"
  DoEvents
  'todo: there is not defined criteria for this metric...
  'X = WPs with multiple EOCs
  'Y = total count of WPs
  'we already have lngY
  strSQL = "SELECT WP,COUNT(EOC) FROM ("
  strSQL = strSQL & "SELECT DISTINCT t.WP,a.EOC "
  strSQL = strSQL & "FROM [tasks.csv] as t "
  strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TASK_UID,EOC FROM [assignments.csv]) AS a ON a.TASK_UID = t.UID) "
  strSQL = strSQL & "GROUP BY WP "
  strSQL = strSQL & "HAVING Count(EOC) > 1"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("WP") & ","
        .MoveNext
      Loop
    End If
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore < 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A202a...done."
  Application.StatusBar = "Getting EVMS: 10A202a...done."
  DoEvents
  
  '===== SCHEDULE =====
  '06A204b - Dangling Logic
  '06A204b todo: ignore first/last milestone
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A204b..."
  Application.StatusBar = "Getting Schedule Metric: 06A204b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A204b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Dangling Logic"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  'cptDECM_frm.lboMetrics.Value = "06A204b"
  DoEvents
  'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = DECM("06A204b")
  'Y = count incomplete Non-LOE tasks/activities & milestones
  'X = count of tasks with open starts or finishes
  'X/Y = 0%
  strSQL = "SELECT * FROM [tasks.csv] WHERE AF IS NULL AND (EVT<>'" & strLOE & "' OR EVT IS NULL) AND SUMMARY='No'"
  oRecordset.Open strSQL, strCon, adOpenKeyset
  lngY = oRecordset.RecordCount
  'start with this list - guilty until proven innocent
  Set oDict = CreateObject("Scripting.Dictionary")
  With oRecordset
    .MoveFirst
    Do While Not .EOF
      oDict.Add CStr(oRecordset("UID")), CStr(oRecordset("UID"))
      .MoveNext
    Loop
  End With
  oRecordset.Close
  
  strSQL = "SELECT t.UID,t.DUR,p.[TYPE],p.[FROM] FROM [tasks.csv] t "
  strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT * FROM [links.csv]) p ON p.TO=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    .MoveFirst
    Do While Not .EOF
      If oRecordset("UID") <> "" And (oRecordset("TYPE") = "SS" Or oRecordset("TYPE") = "FS") Then  'And oRecordset("DUR") > 0
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      ElseIf oRecordset("UID") <> "" And oRecordset("DUR") = 0 And Not IsNull(oRecordset("TYPE")) Then
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      End If
      .MoveNext
    Loop
  End With
  oRecordset.Close
  'extract the guilty to a string for later consolidation
  For lngItem = 0 To oDict.Count - 1
    strLinks = strLinks & oDict.Items(lngItem) & "," 'keep trailing comma, we're going to build on it
  Next lngItem
    
  'now do successors - guilty until proven innocent
  oDict.RemoveAll
  strSQL = "SELECT * FROM [tasks.csv] WHERE AF IS NULL AND (EVT<>'" & strLOE & "' OR EVT IS NULL) AND SUMMARY='No'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    .MoveFirst
    Do While Not .EOF
      oDict.Add CStr(oRecordset("UID")), CStr(oRecordset("UID"))
      .MoveNext
    Loop
    .Close
  End With
  
  strSQL = "SELECT t.UID,t.DUR,s.[TYPE],s.[TO] FROM [tasks.csv] t "
  strSQL = strSQL & "LEFT JOIN (SELECT DISTINCT * FROM [links.csv]) s ON s.[FROM]=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    .MoveFirst
    Do While Not .EOF
      If oRecordset("UID") <> "" And (oRecordset("TYPE") = "FF" Or oRecordset("TYPE") = "FS") Then  'And oRecordset("DUR") > 0
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      ElseIf oRecordset("UID") <> "" And oRecordset("DUR") = 0 And Not IsNull(oRecordset("TYPE")) Then
        If oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Remove (CStr(oRecordset("UID")))
      End If
      .MoveNext
    Loop
    .Close
  End With
  
  'extract the guilty to a string for later consolidation
  For lngItem = 0 To oDict.Count - 1
    strLinks = strLinks & oDict.Items(lngItem) & ","
  Next lngItem
  strLinks = Left(strLinks, Len(strLinks) - 1)
  oDict.RemoveAll
  strList = ""
  For Each vField In Split(strLinks, ",")
    If Len(vField) > 0 And Not oDict.Exists(vField) Then
      oDict.Add vField, vField
      strList = strList & vField & ","
    End If
  Next vField
  lngX = oDict.Count
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A204b...done."
  Application.StatusBar = "Getting Schedule Metric: 06A204b...done."
  DoEvents
  
  '06A205a - Lags (todo: what about leads?)
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A205a..."
  Application.StatusBar = "Getting Schedule Metric: 06A205a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A205a"
  'cptDECM_frm.lboMetrics.Value = "06A205a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Lags"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
  DoEvents
  'X = count of incomplete tasks/activities & milestones with at least one lag in the pred logic
  'Y = count of incomplete tasks/activities & milestones in the IMS
  'X/Y <=10%
  'we already have lngY...
  strSQL = "SELECT t.UID FROM [tasks.csv] t "
  strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TO FROM [links.csv] WHERE LAG>0) p ON p.TO=t.UID " 'todo
  'todo: ABOVE LINE OR strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TO FROM [links.csv] WHERE LAG<>0) p ON p.TO=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL) "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    If lngX > 0 Then
      strList = ""
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.1 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A205a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A205a...done."
  DoEvents
  
  '06A208a - summary tasks with logic
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A208a..."
  Application.StatusBar = "Getting Schedule Metric: 06A208a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A208a"
  'cptDECM_frm.lboMetrics.Value = "06A208a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Summary Logic"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of summary tasks/activities with logic applied (# predecessors > 0 or # successors > 0)
  'X = 0
  strSQL = "SELECT t.UID FROM [tasks.csv] t "
  strSQL = strSQL & "INNER JOIN [links.csv] l ON L.TO=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='Yes'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A208a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A208a...done."
  DoEvents
  
  '06A209a - hard constraints
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A209a..."
  Application.StatusBar = "Getting Schedule Metric: 06A209a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A209a"
  'cptDECM_frm.lboMetrics.Value = "06A209a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Hard Constraints"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = count of incomplete tasks/activities & milestones with hard constraints
  'Y = count of incomplete tasks/activities & milestones
  'X/Y = 0%
  'we already have lngY...
  strSQL = "SELECT UID FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND AF IS NULL AND (EVT<>'" & strLOE & "' OR EVT IS NULL) "
  strSQL = strSQL & "AND (CONST='SNLT' OR CONST='FNLT' OR CONST='MSO' OR CONST='MFO')"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.1 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A209a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A209a...done."
  DoEvents
  
  '06A210a - LOE Driving Discrete
  'todo: add note: filter shows both LOE pred and Non-LOE successor
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A210a..."
  Application.StatusBar = "Getting Schedule Metric: 06A210a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A210a"
  'cptDECM_frm.lboMetrics.Value = "06A210a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "LOE Driving Discrete"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  'X = count of incomplete LOE tasks/activities in the IMS with at least one Non-LOE successor
  'Y = count of incomplete LOE tasks/activities in the IMS
  'X/Y = 0%
  strSQL = "SELECT UID FROM [tasks.csv] WHERE EVT='" & strLOE & "' AND AF IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT DISTINCT t.UID,t.EVT,s.TO,d.EVT "
  strSQL = strSQL & "FROM ([tasks.csv] t "
  strSQL = strSQL & "INNER JOIN (SELECT [FROM],TO FROM [links.csv]) s ON s.[FROM]=t.UID) "
  strSQL = strSQL & "INNER JOIN [tasks.csv] d ON d.UID=s.TO "
  strSQL = strSQL & "WHERE t.EVT='" & strLOE & "' AND t.AF IS NULL AND d.EVT<>'" & strLOE & "'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    oDict.RemoveAll
    strList = ""
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        If Not oDict.Exists(CStr(oRecordset("UID"))) Then oDict.Add CStr(oRecordset("UID")), CStr(oRecordset("UID"))
        strList = strList & .Fields("UID") & "," & .Fields("TO") & "," 'includes guilty successors
        .MoveNext
      Loop
    End If
    lngX = oDict.Count
    
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList 'todo: need guilty link too
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A210a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A210a...done."
  DoEvents
  
  '06A211a - High Float
  '06A211a - High Float todo: refine TS into effective days (elapsed, etc)
  '06A211a - High Float todo: need rationale; user can mark 'acceptable'
  '06A211a - High Float todo: allow user input for lngX
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A211a..."
  Application.StatusBar = "Getting Schedule Metric: 06A211a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A211a"
  'cptDECM_frm.lboMetrics.Value = "06A211a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "High Float"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 20%"
  DoEvents
'  X = count of high total float Non-LOE tasks/activities & milestones sampled with inadequate rationale
'  Y = count of high total float Non-LOE tasks/activities & milestones sampled
'  X/Y <= 20%
  strSQL = "SELECT UID,ROUND(TS/480,2) AS HTF " 'todo: replace 480 with user settings?
  strSQL = strSQL & "FROM [tasks.csv] "
  strSQL = strSQL & "WHERE EVT<>'" & strLOE & "' "
  strSQL = strSQL & "GROUP BY UID,ROUND(TS/480,2) "
  strSQL = strSQL & "HAVING ROUND(TS/480,2)>44 "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    lngY = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.2 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A211a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A211a...done."
  DoEvents
  
  '06A212a - out of sequence
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A212a..."
  Application.StatusBar = "Getting Schedule Metric: 06A212a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A212a"
  'cptDECM_frm.lboMetrics.Value = "06A501a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Out of Sequence"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of out of sequence conditions
  strList = cptGetOutOfSequence 'function returns lngX|uid vbtab uid vbtab uid
  lngX = CLng(Split(strList, "|")(0))
  strList = Split(strList, "|")(1)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = ""
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  ElseIf lngX > 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description; see workbook"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A212a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A212a...done."
  DoEvents
  
  '6A301a - vertical integration todo: lower level baselines rollup
  
  '6A401a - critical path todo: can our tool satisfy?
  
  '6A501a - baselines
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A501a..."
  Application.StatusBar = "Getting Schedule Metric: 06A501a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A501a"
  'cptDECM_frm.lboMetrics.Value = "06A501a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Baselines"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of tasks/activities & milestones without baseline dates
  'Y = Total count of tasks/activities & milestones
  'X/Y <= 5%
  strSQL = "SELECT UID,BLS,BLF FROM [tasks.csv] WHERE SUMMARY='No'"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,BLS,BLF FROM [tasks.csv] WHERE SUMMARY='No' AND (BLS IS NULL OR BLF IS NULL)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A501a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A501a...done."
  DoEvents
  
  '06A504a - AS changed - too complicated, keep it manual
  
  '06A504b - AF changed - too complicated, keep it manual

  '06A505a - In-Progress Tasks Have AS
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A505a..."
  Application.StatusBar = "Getting Schedule Metric: 06A505a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A505a"
  'cptDECM_frm.lboMetrics.Value = "06A505a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "In-Progress Tasks w/o Actual Start"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = count of in-progress tasks/activities & milestones with no actual start date
  'Y = count of in-progress tasks/activities & milestones
  'X/Y <= 5%
  strSQL = "SELECT UID,EVP,[AS] FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP<100 AND EVP>0 "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,EVP,[AS] FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP<100 AND EVP>0 "
  strSQL = strSQL & "AND [AS] IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A505a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A505a...done."
  DoEvents
  
  '06A505b - Complete Tasks Have AF
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A505b..."
  Application.StatusBar = "Getting Schedule Metric: 06A505a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A505b"
  'cptDECM_frm.lboMetrics.Value = "06A505b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Complete Tasks w/o Actual Finish"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = count of complete tasks/activities & milestones with no actual finish date
  'Y = count of complete tasks/activities & milestones
  'X/Y <= 5%
  strSQL = "SELECT UID,EVP,AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP=100 "
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,EVP,AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE SUMMARY='No' AND EVP=100 AND AF IS NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = oRecordset.RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A505b...done."
  Application.StatusBar = "Getting Schedule Metric: 06A505b...done."
  DoEvents
  
  '06A506a - bogus actuals
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506a..."
  Application.StatusBar = "Getting Schedule Metric: 06A506a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A506a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Bogus Actuals"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  'cptDECM_frm.lboMetrics.Value = "06A506a"
  DoEvents
  'X = count of tasks/activities & milestones with either actual start or actual finish after status date
  'Y = count of tasks/activities & milestones with an actual start date
  'X/Y <= 5%
  strSQL = "SELECT UID,[AS],AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE [AS] IS NOT NULL OR AF IS NOT NULL"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngY = oRecordset.RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  strSQL = "SELECT UID,[AS],AF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE ([AS]>#" & dtStatus & "# OR AF>#" & dtStatus & "#)"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  dblScore = Round(lngX / IIf(lngY = 0, 1, lngY), 2)
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.05 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506a...done."
  Application.StatusBar = "Getting Schedule Metric: 06A506a...done."
  DoEvents
  
  '06A506b - bogus forecast
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506b..."
  Application.StatusBar = "Getting Schedule Metric: 06A506b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A506b"
  'cptDECM_frm.lboMetrics.Value = "06A506b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Bogus Forecast"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of incomplete tasks/activities & milestones with either forecast start or forecast finish before the status date
  'X = 0
  strSQL = "SELECT UID,FS,FF FROM [tasks.csv] "
  strSQL = strSQL & "WHERE ((FS<#" & dtStatus & "# AND [AS] IS NULL) "
  strSQL = strSQL & "OR (FF<#" & dtStatus & "# AND AF IS NULL))"
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset
    lngX = .RecordCount
    strList = ""
    If lngX > 0 Then
      .MoveFirst
      Do While Not .EOF
        strList = strList & .Fields("UID") & ","
        .MoveNext
      Loop
    End If
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506b...done."
  Application.StatusBar = "Getting Schedule Metric: 06A506b...done."
  DoEvents
  
  Application.StatusBar = "DECM Scoring Complete"
  cptDECM_frm.lblStatus.Caption = "DECM Scoring Complete"
  DoEvents
  
exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Set oDict = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  cptSpeed False
  Application.StatusBar = ""
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing
  For lngFile = 0 To FreeFile
    Close #lngFile
  Next lngFile
  Set oLink = Nothing
  Set oTask = Nothing
  
  Exit Sub
err_here:
 On Error Resume Next
 Call cptHandleErr("cptDECM", "cptDECM_GET_DATA", Err, Erl)
 Resume exit_here
End Sub

Function DECM(strDECM As String, Optional blnNotify As Boolean = False) As Double
  Dim oTask As MSProject.Task
  Dim oLinks As Scripting.Dictionary
  Dim oLink As TaskDependency
  Dim lngX As Long
  Dim lngY As Long
  Dim strLinks As String
  
  'If Not ValidMap Then GoTo exit_here
  
  Select Case strDECM
    Case "06A204b" 'dangling logic
      
      ActiveWindow.TopPane.Activate
      FilterClear
      GroupClear
      OptionsViewEx DisplaySummaryTasks:=True
      OutlineShowAllTasks
    
      Set oLinks = CreateObject("Scripting.Dictionary")
      oLinks.Add 0, "FF"
      oLinks.Add 1, "FS"
      oLinks.Add 2, "SF"
      oLinks.Add 3, "SS"
      
      cptSpeed True
      
      lngEVT = CLng(Split(cptGetSetting("Integration", "EVT"), "|")(0))
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task
        'todo: external?
        If Not oTask.Active Then GoTo next_task
        oTask.Marked = False
        If oTask.Summary Then GoTo next_task
        If oTask.GetField(lngEVT) = "A" Then GoTo next_task
        If IsDate(oTask.ActualFinish) Then GoTo next_task
        lngY = lngY + 1
        'whether task or milestone, must have PE and SE
        If oTask.PredecessorTasks.Count = 0 Then
          lngX = lngX + 1
          oTask.Marked = True
          GoTo next_task
        End If
        If oTask.SuccessorTasks.Count = 0 Then
          lngX = lngX + 1
          oTask.Marked = True
          GoTo next_task
        End If
        If oTask.Duration > 0 Then 'examine tasks
          strLinks = ""
          For Each oLink In oTask.TaskDependencies
            If oLink.To = oTask Then 'examine predecessors
              strLinks = strLinks & oLinks(oLink.Type) & ","
            End If
          Next oLink
          If InStr(strLinks, "FS") = 0 And InStr(strLinks, "SS") = 0 Then
            lngX = lngX + 1
            oTask.Marked = True
            GoTo next_task
          End If
          strLinks = ""
          For Each oLink In oTask.TaskDependencies
            If oLink.From = oTask Then 'examine successors
              strLinks = strLinks & oLinks(oLink.Type) & ","
            End If
          Next oLink
          If InStr(strLinks, "FS") = 0 And InStr(strLinks, "FF") = 0 Then
            lngX = lngX + 1
            oTask.Marked = True
            GoTo next_task
          End If
        End If
next_task:
      Next oTask
      
      cptSpeed False
      
      ActiveWindow.TopPane.Activate
      OptionsViewEx DisplaySummaryTasks:=True
      OutlineShowAllTasks
      OptionsViewEx DisplaySummaryTasks:=False
      SetAutoFilter "Marked", pjAutoFilterFlagYes
      
      If blnNotify Then MsgBox "X: " & lngX & vbCrLf & "Y: " & lngY & vbCrLf & "X/Y = " & Format(lngX / lngY, "0%"), vbInformation + vbOKOnly, "06A204b"
      cptDECM_frm.txtTitle = "X: " & lngX & vbCrLf & "Y: " & lngY
      'X = count of incomplete Non-LOE tasks/activities & milestones in the IMS WITH open starts or finishes
      'Y = count of incomplete Non-LOE tasks/activities & milestones in the IMS
      DECM = Round(lngX / lngY, 2)
  End Select

exit_here:
  Set oTask = Nothing
  Set oLink = Nothing
  Set oLinks = Nothing
  Exit Function

End Function

Private Sub DumpRecordsetToExcel(ByRef oRecordset As ADODB.Recordset)
  'objects
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error GoTo err_here
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.APplication")
  End If
  
  Set oExcel = GetObject(, "Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "AD HOC"
  For lngItem = 0 To oRecordset.Fields.Count - 1
    oWorksheet.Cells(1, lngItem + 1) = oRecordset.Fields(lngItem).Name
  Next lngItem
  oWorksheet.[A2].Select
  oWorksheet.[A2].CopyFromRecordset oRecordset
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
  oWorksheet.[A1].AutoFilter
  oWorksheet.Columns.AutoFit

exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("basDECM_bas", "DumpRecordsetToExcel", Err, Erl)
  Resume exit_here

End Sub

Sub opencsv(strFile)
  Shell "C:\Windows\notepad.exe '" & Environ("tmp") & "\" & strFile & "'", vbNormalFocus
End Sub

Sub cptDECM_EXPORT(Optional blnDetail As Boolean = False)
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRange As Excel.Range
  Dim oCell As Excel.Range
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  cptSpeed True
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  
  oExcel.WindowState = xlMaximized
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "DECM Dashboard"
  oWorksheet.[A1:I1] = cptDECM_frm.lboHeader.List
  oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].Offset(cptDECM_frm.lboMetrics.ListCount - 1, cptDECM_frm.lboMetrics.ColumnCount - 1)) = cptDECM_frm.lboMetrics.List
  oWorksheet.[A2].Select
  With oExcel.ActiveWindow
    .Zoom = 85
    .SplitRow = 1
    .SplitColumn = 0
    .FreezePanes = True
  End With
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).HorizontalAlignment = xlLeft
  With oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown))
    .Font.Name = "Calibri"
    .Font.Size = 11
    .HorizontalAlignment = xlCenter
    .Columns.AutoFilter
    .Columns.AutoFit
  End With
  
  oWorksheet.Columns(1).HorizontalAlignment = xlLeft
  oWorksheet.Columns(2).HorizontalAlignment = xlLeft
  'oWorksheet.Columns(8).HorizontalAlignment = xlLeft
  oWorksheet.Columns("H:I").Delete
  
  With oWorksheet.Range(oWorksheet.[G2], oWorksheet.[G2].End(xlDown))
    .Replace what:=strPass, Replacement:="2", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    .Replace what:=strFail, Replacement:="0", lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    .FormatConditions.AddIconSetCondition
    .FormatConditions(.FormatConditions.Count).SetFirstPriority
    With .FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = True
        .IconSet = oWorkbook.IconSets(xl3Symbols)
    End With
    With .FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 1
        .Operator = 7
    End With
    With .FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 2
        .Operator = 7
    End With
  End With
  
  blnDetail = True 'todo: make this an option on the form
  
  If blnDetail Then
    With cptDECM_frm
      For lngItem = 0 To .lboMetrics.ListCount - 1
        .lboMetrics.Value = .lboMetrics.List(lngItem)
        .lboMetrics.Selected(lngItem) = True
        .lboMetrics_AfterUpdate
        Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
        oWorksheet.Activate
        oWorksheet.Name = .lboMetrics.List(lngItem)
        oWorksheet.Tab.Color = 5287936
        oExcel.ActiveWindow.Zoom = 85
        SelectAll
        EditCopy
        On Error Resume Next
        Set oTasks = ActiveSelection.Tasks
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If Not oTasks Is Nothing Then
          oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
          oWorksheet.[A2].Select
          oWorksheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False
          oWorksheet.Cells.Font.Name = "Calibri"
          oWorksheet.Cells.Font.Size = 11
          oWorksheet.Cells.WrapText = False
          oWorksheet.[B3].Select
          oExcel.ActiveWindow.FreezePanes = True
          oWorksheet.Columns.AutoFit
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Tab.Color = 192
          End If
        End If
        Set oTasks = Nothing
      Next lngItem
    End With
    'create hyperlinks
    Set oWorksheet = oWorkbook.Sheets("DECM Dashboard")
    oWorksheet.Activate
    Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].End(xlDown))
    For Each oCell In oRange.Cells
      oWorksheet.Hyperlinks.Add Anchor:=oCell, Address:="", SubAddress:="'" & CStr(oCell.Value) & "'!A1", TextToDisplay:=CStr(oCell.Value), ScreenTip:="Jump to " & CStr(oCell.Value)
    Next oCell
  End If
  
  Application.ActivateMicrosoftApp pjMicrosoftExcel

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  cptSpeed False
  Set oCell = Nothing
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_bas", "cptDECM_EXPORT", Err, Erl)
  Resume exit_here
End Sub

Sub cptDECM_UPDATE_VIEW(strMetric As String, Optional strList As String)

  ScreenUpdating = False
  ActiveWindow.TopPane.Activate
  FilterClear
  GroupClear
  OptionsViewEx DisplaySummaryTasks:=True
  OutlineShowAllTasks
  OptionsViewEx DisplaySummaryTasks:=False

  Select Case strMetric
    Case "05A101a" '1 CA : 1 OBS
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
      'todo: group by CA,OBS
    
    Case "05A102a" '1 CA : 1 CAM
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
      'todo: group by CA,CAM
    
    Case "05A103a" '1 CA : 1 WBS
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
      'todo: group by CA,WBS
    
    Case "06A212a" 'out of sequence
      If Len(strList) > 0 Then
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
    
    Case "10A102a" '1 WP : 1 EVT
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
      'todo: group by WP,EVT
    
    Case "10A109b" 'WP with no budget
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WPCN"), "|")(0)), pjAutoFilterIn, "contains", strList 'todo: "WPCN" > "WP"
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
    
    Case "10A202a" 'WP with Mixed EOCs
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WPCN"), "|")(0)), pjAutoFilterIn, "contains", strList 'todo: "WPCN" > "WP"
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
        
    Case Else
      If Len(strList) > 0 Then
        strList = Left(Replace(strList, ",", vbTab), Len(strList) - 1) 'remove last comma
        SetAutoFilter "Unique ID", pjAutoFilterIn, "contains", strList
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
      
  End Select
  SelectBeginning
  ScreenUpdating = True
End Sub


Function cptGetOutOfSequence() As String
  'todo: rename and convert to function?
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oOOS As Scripting.Dictionary
  Dim oCalendar As MSProject.Calendar
  Dim oSubproject As MSProject.Subproject
  Dim oSubMap As Scripting.Dictionary
  Dim oTask As MSProject.Task
  Dim oLink As MSProject.TaskDependency
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  'strings
  Dim strOOS As String
  Dim strEarliest As String
  Dim strProject As String
  Dim strMacro As String
  Dim strMsg As String
  Dim strProjectNumber As String
  Dim strProjectName As String
  Dim strDir As String
  Dim strFile As String
  'longs
  Dim lngLagType As Long
  Dim lngLag As Long
  Dim lngFactor As Long
  Dim lngToUID As Long
  Dim lngFromUID As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngLastRow As Long
  Dim lngOOS As Long
  'integers
  'doubles
  'booleans
  Dim blnElapsed As Boolean
  Dim blnSubprojects As Boolean
  'variants
  'dates
  Dim dtStatus As Date
  Dim dtDate As Date
  
  If cptErrorTrapping Then
    On Error GoTo err_here
    cptSpeed True
  Else
    On Error GoTo 0
  End If
  
  blnSubprojects = ActiveProject.Subprojects.Count > 0
  If blnSubprojects Then
    'get correct task count
    lngTasks = ActiveProject.Tasks.Count
    'set up mapping
    If oSubMap Is Nothing Then
      Set oSubMap = CreateObject("Scripting.Dictionary")
    Else
      oSubMap.RemoveAll
    End If
    For Each oSubproject In ActiveProject.Subprojects
      If InStr(oSubproject.Path, "\") > 0 Then 'offline
        oSubMap.Add Replace(Dir(oSubproject.Path), ".mpp", ""), 0
      ElseIf InStr(oSubproject.Path, "<>") > 0 Then 'online
        oSubMap.Add oSubproject.Path, 0
      End If
      lngTasks = lngTasks + oSubproject.SourceProject.Tasks.Count
    Next oSubproject
    For Each oTask In ActiveProject.Tasks
      If oTask Is Nothing Then GoTo next_mapping_task
      If oSubMap.Exists(oTask.Project) Then
        If oSubMap(oTask.Project) > 0 Then GoTo next_mapping_task
        If Not oTask.Summary Then
          oSubMap.Item(oTask.Project) = CLng(oTask.UniqueID / 4194304)
        End If
      End If
next_mapping_task:
    Next oTask
    
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  
  Set oExcel = CreateObject("Excel.Application")
  On Error Resume Next
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "06A212a"
  oWorksheet.[A2:K2] = Split("UID,ID,TASK,DATE,TYPE,LAG,UID,ID,TASK,DATE,COMMENT", ",")
  oWorksheet.[A1:D1].Merge
  oWorksheet.[A1].Value = "FROM"
  oWorksheet.[A1].HorizontalAlignment = xlCenter
  oWorksheet.[G1:J1].Merge
  oWorksheet.[G1].Value = "TO"
  oWorksheet.[G1].HorizontalAlignment = xlCenter
  oWorksheet.[A1:K2].Font.Bold = True
  oExcel.EnableEvents = False
    
  lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
  lngTask = 0
  
  Set oOOS = CreateObject("Scripting.Dictionary")
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If oTask.Summary Then GoTo next_task 'skip summary tasks
    If Not oTask.Active Then GoTo next_task 'skip inactive tasks
    If IsDate(oTask.ActualFinish) Then GoTo next_task 'incomplete predecessors only
    
    For Each oLink In oTask.TaskDependencies
      If oLink.From.Guid = oTask.Guid Then   'predecessors only
        If Not oLink.To.Active Then GoTo next_link
        lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
        If blnSubprojects And oLink.From.ExternalTask Then
          'fix the pred UID if master-sub
          lngFromUID = oLink.From.GetField(185073906) Mod 4194304
          strProject = oLink.From.Project
          If InStr(oLink.From.Project, "\") > 0 Then
            strProject = Replace(strProject, ".mpp", "")
            strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
          End If
          lngFactor = oSubMap(strProject)
          lngFromUID = (lngFactor * 4194304) + lngFromUID
        Else
          If blnSubprojects Then
            lngFactor = Round(oTask / 4194304, 0)
            lngFromUID = (lngFactor * 4194304) + oLink.From.UniqueID
          Else
            lngFromUID = oLink.From.UniqueID
          End If
        End If
        lngToUID = oLink.To.UniqueID
        lngLag = 0
        If oLink.Lag <> 0 Then
          'elapsed lagType properties are even
          'see https://learn.microsoft.com/en-us/office/vba/api/project.pjformatunit
          blnElapsed = False
          If (oLink.LagType Mod 2) = 0 Then
            blnElapsed = True
          End If
          lngLag = oLink.Lag
        End If
        If oLink.To.Calendar = "None" Then
          Set oCalendar = ActiveProject.Calendar
        Else
          Set oCalendar = oLink.To.CalendarObject
        End If
        Select Case oLink.Type
          Case pjFinishToFinish
            'get target successor finish date
            'account for lag
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Finish)
            Else
              dtDate = Application.DateAdd(oLink.From.Finish, lngLag, oCalendar)
            End If
            If oLink.To.Finish < dtDate Or IsDate(oLink.To.ActualFinish) Then
              lngOOS = lngOOS + 1
              If Not oOOS.Exists(oLink.From.UniqueID) Then oOOS.Add oLink.From.UniqueID, oLink.From.UniqueID
              If Not oOOS.Exists(oLink.To.UniqueID) Then oOOS.Add oLink.To.UniqueID, oLink.To.UniqueID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnSubprojects, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Finish
              oWorksheet.Cells(lngLastRow, 5) = "FF"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnSubprojects, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Finish
              If IsDate(oLink.To.ActualFinish) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Finish"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish (has deadline)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Finish" 'what else would cause this?
                End If
              End If
            End If
          Case pjFinishToStart
            'get target successor start date
            'account for lag
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Finish)
            Else
              dtDate = Application.DateAdd(oLink.From.Finish, lngLag, oCalendar)
            End If
            'compare and report
            If oLink.To.Start < dtDate Or IsDate(oLink.To.ActualStart) Then
              lngOOS = lngOOS + 1
              If Not oOOS.Exists(oLink.From.UniqueID) Then oOOS.Add oLink.From.UniqueID, oLink.From.UniqueID
              If Not oOOS.Exists(oLink.To.UniqueID) Then oOOS.Add oLink.To.UniqueID, oLink.To.UniqueID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnSubprojects, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Finish
              oWorksheet.Cells(lngLastRow, 5) = "FS"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnSubprojects, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Start
              If IsDate(oLink.To.ActualStart) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Start"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish (has deadline)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Finish" 'what else could cause this?
                End If
              End If
            End If
          Case pjStartToStart
            'get target successor start date
            'account for lag
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Start)
            Else
              dtDate = Application.DateAdd(oLink.From.Start, lngLag, oCalendar)
            End If
            'compare and report
            If IsDate(oLink.To.ActualStart) Or oLink.To.Start < dtDate Then 'should not be an issue if both have actual starts
              lngOOS = lngOOS + 1
              If Not oOOS.Exists(oLink.From.UniqueID) Then oOOS.Add oLink.From.UniqueID, oLink.From.UniqueID
              If Not oOOS.Exists(oLink.To.UniqueID) Then oOOS.Add oLink.To.UniqueID, oLink.To.UniqueID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnSubprojects, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Start
              oWorksheet.Cells(lngLastRow, 5) = "SS"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnSubprojects, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Start
              If IsDate(oLink.To.ActualStart) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Start"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start (has deadline lead)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Start < Predecessor Start" 'what else could cause this?
                End If
              End If
            End If
          Case pjStartToFinish
            'this should never happen
            'get target finish
            If blnElapsed Then
              dtDate = DateAdd("n", lngLag, oLink.From.Start)
            Else
              dtDate = Application.DateAdd(oLink.From.Start, lngLag, oCalendar)
            End If
            'compare and report
            If IsDate(oLink.To.ActualFinish) Or oLink.To.Finish < oLink.From.Start Then
              lngOOS = lngOOS + 1
              If Not oOOS.Exists(oLink.From.UniqueID) Then oOOS.Add oLink.From.UniqueID, oLink.From.UniqueID
              If Not oOOS.Exists(oLink.To.UniqueID) Then oOOS.Add oLink.To.UniqueID, oLink.To.UniqueID
              oWorksheet.Cells(lngLastRow, 1) = lngFromUID
              oWorksheet.Cells(lngLastRow, 2) = IIf(blnSubprojects, "-", oLink.From.ID)
              oWorksheet.Cells(lngLastRow, 3) = oLink.From.Name
              oWorksheet.Cells(lngLastRow, 4) = oLink.From.Start
              oWorksheet.Cells(lngLastRow, 5) = "SF"
              oWorksheet.Cells(lngLastRow, 6) = oLink.Lag / (8 * 60)
              oWorksheet.Cells(lngLastRow, 7) = lngToUID
              oWorksheet.Cells(lngLastRow, 8) = IIf(blnSubprojects, "-", oLink.To.ID)
              oWorksheet.Cells(lngLastRow, 9) = oLink.To.Name
              oWorksheet.Cells(lngLastRow, 10) = oLink.To.Finish
              If IsDate(oLink.To.ActualFinish) Then
                oWorksheet.Cells(lngLastRow, 11) = "Successor has Actual Finish"
              Else
                If IsDate(oLink.To.ConstraintDate) And ActiveProject.HonorConstraints Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start (has " & Choose(oLink.To.ConstraintType, "", "MSO", "MFO", "SNET", "SNLT", "FNET", "FNLT") & " constraint)"
                ElseIf lngLag < 0 Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start (has " & Format(lngLag / (60 * 8), "#0d") & " lead)"
                ElseIf IsDate(oLink.To.Deadline) Then
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start (has deadline)"
                Else
                  oWorksheet.Cells(lngLastRow, 11) = "Successor Finish < Predecessor Start"
                End If
              End If
            End If
        End Select
      End If
next_link:
    Next oLink
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "[06A212a] Analyzing Out of Sequence Status...(" & Format(lngTask / lngTasks, "0%") & ")" & IIf(lngOOS > 0, " | " & lngOOS & " found", "")
    DoEvents
  Next oTask
    
  'get ~count of OOS oTasks
  lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
  
  'only open workbook if OOS oTasks found
  If lngOOS = 0 Then
    oWorkbook.Close False
    GoTo return_val
  Else
    strOOS = Join(oOOS.Keys, vbTab)
  End If
    
  With oExcel.ActiveWindow
    .Zoom = 85
    .SplitRow = 2
    .SplitColumn = 0
    .FreezePanes = True
  End With
  oWorksheet.Columns.AutoFit
  
  'add a macro
  strMacro = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf
  strMacro = strMacro & " Dim MSPROJ As Object, Task As Object" & vbCrLf
  strMacro = strMacro & " Dim strFrom As String, strTo As String" & vbCrLf
  strMacro = strMacro & "" & vbCrLf
  strMacro = strMacro & " On Error GoTo exit_here" & vbCrLf
  strMacro = strMacro & "" & vbCrLf
  strMacro = strMacro & " If Target.Cells.Count <> 1 Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  If Target.Row < 3 Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  If Target.Column > Me.[A2].End(xlToRight).Column Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  If Target.Row > Me.[A1048576].End(xlUp).Row Then Exit Sub" & vbCrLf
  strMacro = strMacro & "  Set MSPROJ = GetObject(, ""MSProject.Application"")" & vbCrLf
  strMacro = strMacro & "  MSPROJ.ActiveWindow.TopPane.Activate" & vbCrLf
  strMacro = strMacro & "  MSPROJ.ScreenUpdating = False" & vbCrLf
  strMacro = strMacro & "  MSPROJ.FilterClear" & vbCrLf
  strMacro = strMacro & "  MSPROJ.OptionsViewEx DisplaySummaryTasks:=True" & vbCrLf
  strMacro = strMacro & "  MSPROJ.OutlineShowAllTasks" & vbCrLf
  strMacro = strMacro & "  strFrom = Me.Cells(Target.Row, 1).Value" & vbCrLf
  strMacro = strMacro & "  strTo = Me.Cells(Target.Row, 7).Value" & vbCrLf
  strMacro = strMacro & "  MSPROJ.SetAutoFilter FieldName:=""Unique ID"", FilterType:=2, Criteria1:=strFrom & Chr$(9) & strTo '1=pjAutoFilterIn" & vbCrLf
  strMacro = strMacro & "  MSPROJ.Find ""Unique ID"", ""equals"", CLng(strFrom)" & vbCrLf
  strMacro = strMacro & "  MSPROJ.EditGoto Date:=MSPROJ.ActiveProject.Tasks.UniqueID(CLng(strFrom)).Start" & vbCrLf & vbCrLf
  strMacro = strMacro & "exit_here:" & vbCrLf
  strMacro = strMacro & "  MSPROJ.ScreenUpdating = True" & vbCrLf
  strMacro = strMacro & "  Set MSPROJ = Nothing" & vbCrLf
  strMacro = strMacro & "End Sub"
  
  oWorkbook.VBProject.VBComponents("Sheet1").CodeModule.AddFromString strMacro
  
  oExcel.Visible = True
  oExcel.WindowState = xlMinimized
  
return_val:
  cptGetOutOfSequence = CStr(lngOOS) & "|" & strOOS
  
exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  oOOS.RemoveAll
  Set oOOS = Nothing
  Set oCalendar = Nothing
  Set oSubproject = Nothing
  Set oSubMap = Nothing
  Application.StatusBar = ""
  oExcel.EnableEvents = True
  cptSpeed False
  Set oTask = Nothing
  Set oLink = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptFindOutOfSequence", Err, Erl)
  Resume exit_here
  
End Function

