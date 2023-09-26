Attribute VB_Name = "cptDECM_bas"
'<cpt_version>v0.0.3</cpt_version>
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
  Dim oComboBox As MSForms.ComboBox
  'strings
  Dim strLOE As String
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
  
  With New cptIntegration_frm
    
    .Caption = "Integration (" & cptGetVersion("cptIntegration_frm") & ")"
    
    For Each vControl In Split("WBS,OBS,CA,CAM,WP,WPM,EVT,LOE,EVP,EOC", ",")
      If vControl = "WBS" Then vControl = "CWBS" 'todo: fix saved setting name
      If vControl = "WP" Then vControl = "WPCN"  'todo: fix saved setting name
      strSetting = cptGetSetting("Integration", CStr(vControl))
      If Len(strSetting) = 0 Then
        If vControl = "EVP" Then
          strSetting = cptGetSetting("Metrics", "cboEVP")
          If Len(strSetting) = 0 Then
            blnValid = False
          Else
            strSetting = strSetting & "|" & FieldConstantToFieldName(strSetting)
            cptSaveSetting "Integration", "EVP", strSetting
          End If
        ElseIf vControl = "EVT" Then
          strSetting = cptGetSetting("Metrics", "cboLOEField")
          If Len(strSetting) = 0 Then
            blnValid = False
          Else
            strSetting = strSetting & "|" & FieldConstantToFieldName(strSetting)
            cptSaveSetting "Integration", "EVT", strSetting
          End If
        ElseIf vControl = "LOE" Then
          strSetting = cptGetSetting("Metrics", "txtLOE")
          If Len(strSetting) = 0 Then
            blnValid = False
          Else
            cptSaveSetting "Integration", "LOE", strSetting
            .cboLOE.Value = strSetting
          End If
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
        On Error Resume Next
        .cboLOE.Value = strLOE
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
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
  Unload cptDECM_frm

  Exit Function
err_here:
  Call cptHandleErr("cptDECM_bas", "ValidMap", Err, Erl)
  Resume exit_here
    
End Function

Sub cptDECM_GET_DATA()
'Optional blnIncompleteOnly As Boolean = True, Optional blnDiscreteOnly As Boolean = True
  'objects
  Dim oException As MSProject.Exception
  Dim oTasks As MSProject.Tasks
  Dim oCell As Excel.Range
  Dim oListObject As Excel.ListObject
  Dim oFile As Scripting.TextStream
  Dim oFSO As Scripting.FileSystemObject
  Dim oAssignment As MSProject.Assignment
  Dim oDict As Scripting.Dictionary
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRecordset As ADODB.Recordset
  Dim oLink As MSProject.TaskDependency
  Dim oTask As MSProject.Task
  'strings
  Dim strProgramAcroymn As String
  Dim strProgramAcronym As String
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
  Dim blnFiscalExists As Boolean
  Dim blnDumpToExcel As Boolean
  'variants
  Dim vFile As Variant
  Dim vHeader As Variant
  Dim vField As Variant
  'dates
  Dim dtPrevious As Date
  Dim dtCurrent As Date
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
  
  strProgramAcronym = cptGetProgramAcronym
  
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
  For Each vFile In Split("wp-ims.csv,wp-ev.csv,wp-not-in-ims.csv,wp-not-in-ev.csv,10A302b-x.csv,10A303a-x.csv", ",")
    Print #lngFile, "[" & vFile & "]"
    Print #lngFile, "Format=CSVDelimited"
    Print #lngFile, "ColNameHeader=False"
    Print #lngFile, "Col1=WP text"
  Next vFile
  Print #1, "[06A506c-x.csv]"
  Print #1, "ColNameHeader=True"
  Print #1, "Format=CSVDelimited"
  Print #1, "Col1=UID integer"
  Print #1, "Col2=P1_TASK_FINISH date"
  Print #1, "Col3=P1_STATUS_DATE date"
  Print #1, "Col4=P1_DELTA Double"
  Print #1, "Col5=P2_TASK_FINISH date"
  Print #1, "Col6=P2_STATUS_DATE date"
  Print #1, "Col7=P2_DELTA Double"
  Print #1, "[fiscal.csv]"
  Print #1, "ColNameHeader=True"
  Print #1, "Format=CSVDelimited"
  Print #1, "Col1=FISCAL_END date"
  Print #1, "Col2=LABEL text"
  
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
  
  '10A103a - 0/100 EVTs in one fiscal period
  cptDECM_frm.lblStatus.Caption = "Getting EVMS Metric: 10A103a..."
  Application.StatusBar = "Getting EVMS Metric: 10A109b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A103a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "0/100 EVTs in >1 period"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of 0-100 EVT incomplete WPs with more than one accounting period of budget
  'Y = Total count of 0-100 EVT incomplete WPs
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] "
  strSQL = strSQL & "WHERE WP IS NOT NULL "
  strSQL = strSQL & "AND AF IS NULL "
  strSQL = strSQL & "AND EVT='F'" 'todo: what about other tools or values?
  With oRecordset
    .Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    lngY = .RecordCount
    'DumpRecordsetToExcel oRecordset
    .Close
  End With
  If lngY > 0 Then
    Set oWorkbook = cptGetEVTAnalysis
    Set oWorksheet = oWorkbook.Sheets(1)
    Set oListObject = oWorksheet.ListObjects(1)
    lngY = oListObject.DataBodyRange.Rows.Count
    oListObject.Range.AutoFilter Field:=6, Criteria1:=">1", Operator:=xlAnd
    lngX = oListObject.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
    strList = ""
    If lngX > 0 Then
      For Each oCell In oListObject.ListColumns("WP").DataBodyRange.SpecialCells(xlCellTypeVisible).Cells
        If InStr(strList, oCell.Value) = 0 Then strList = strList & oCell.Value & vbTab
      Next oCell
    End If
  Else
    lngX = 0
  End If
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
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A103a...done."
  Application.StatusBar = "Getting EVMS: 10A103a...done."
  DoEvents
  
  '10A109b - all WPs have budget
  cptDECM_frm.lblStatus.Caption = "Getting EVMS Metric: 10A109b..."
  Application.StatusBar = "Getting EVMS Metric: 10A109b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A109b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "WPs With Budgets"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 5%"
  DoEvents
  'X = Count of WPs/PPs/SLPPs with BAC = 0
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
  'todo: there is no defined criteria for this metric...
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
  
  '10A302b - PPs with progress
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 10A302b..."
  Application.StatusBar = "Getting Schedule Metric: 10A302b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A302b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "PPs w/EVP > 0"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 2%"
  DoEvents
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] "
  strSQL = strSQL & "WHERE EVT='K' " 'todo: what about other values/tools
  oRecordset.Open strSQL, strCon, adOpenStatic, adLockReadOnly
  If oRecordset.EOF Then
    lngY = 0
    lngX = 0
    dblScore = 0
    strList = ""
  Else
    lngY = oRecordset.RecordCount
    oRecordset.Close
    strSQL = "SELECT WP,IIF(SUM(EVP)>0,TRUE,FALSE) AS HasProgress  FROM [tasks.csv] "
    strSQL = strSQL & "WHERE EVT='K' "
    strSQL = strSQL & "GROUP BY WP "
    strSQL = strSQL & "HAVING SUM(EVP)>0"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    If oRecordset.EOF Then
      lngX = 0
    Else
      lngX = oRecordset.RecordCount
    End If
    dblScore = Round(lngX / lngY, 2)
    strList = ""
    With oRecordset
      .MoveFirst
      Do While Not .EOF
        strList = strList & oRecordset(0) & ","
        .MoveNext
      Loop
    End With
    Set oFile = oFSO.CreateTextFile(strDir & "\10A302b-x.csv", True)
    oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
    oFile.Close
  End If
  oRecordset.Close
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.02 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A302b...done."
  Application.StatusBar = "Getting EVMS: 10A302b...done."
  DoEvents
  
  '10A303a - all PPs have duration?
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 10A303a..."
  Application.StatusBar = "Getting Schedule Metric: 06A101a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "10A303a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "PPs duration = 0"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
  DoEvents
  'we already have lngY
  If lngY = 0 Then
    lngX = 0
    dblScore = 0
    strList = ""
  Else
    strSQL = "SELECT WP,IIF(SUM(DUR)>0,TRUE,FALSE) AS HasProgress  FROM [tasks.csv] "
    strSQL = strSQL & "WHERE EVT='K' "
    strSQL = strSQL & "GROUP BY WP "
    strSQL = strSQL & "HAVING SUM(EVP)>0"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    strList = ""
    If oRecordset.EOF Then
      lngX = 0
    Else
      lngX = oRecordset.RecordCount
      With oRecordset
        .MoveFirst
        Do While Not .EOF
          strList = strList & oRecordset(0) & ","
          .MoveNext
        Loop
      End With
      Set oFile = oFSO.CreateTextFile(strDir & "\10A303a-x.csv", True)
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oFile.Close
    End If
    dblScore = Round(lngX / lngY, 2)
    oRecordset.Close
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = Format(dblScore, "0%")
  If dblScore <= 0.02 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 10A303a...done."
  Application.StatusBar = "Getting EVMS: 10A303a...done."
  DoEvents

  '===== SCHEDULE =====
  '06A101a - WPs Missing between IMS vs EV
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A101a..."
  Application.StatusBar = "Getting Schedule Metric: 06A101a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A101a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "WPs IMS vs EV Tool"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y = 0%"
  DoEvents
  strSQL = "SELECT DISTINCT WP FROM [tasks.csv] "
  strSQL = strSQL & "WHERE AF IS NULL AND EVT<>'" & strLOE & "' AND SUMMARY='No'"
  oRecordset.Open strSQL, strCon, adOpenKeyset
  lngX = oRecordset.RecordCount 'pending upload
  lngY = oRecordset.RecordCount 'pending upload
  Set oFile = oFSO.CreateTextFile(strDir & "\wp-ims.csv", True)
  oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  oRecordset.Close
  oFile.Close
  FileCopy strDir & "\wp-ims.csv", strDir & "\wp-not-in-ev.csv"
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
  'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting EVMS: 06A101a...done."
  Application.StatusBar = "Getting EVMS: 06A101a...done."
  DoEvents
    
  '06A204b - Dangling Logic
  '06A204b - todo: ignore first/last milestone - how?
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
  'todo: to include leads replace above with strSQL = strSQL & "INNER JOIN (SELECT DISTINCT TO FROM [links.csv] WHERE LAG<>0) p ON p.TO=t.UID "
  strSQL = strSQL & "WHERE t.SUMMARY='No' AND t.AF IS NULL AND (t.EVT<>'" & strLOE & "' OR t.EVT IS NULL) "
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
  
  '6A301a - vertical integration todo: lower level baselines rollup...refers to supplemental schedules...too complicated
  
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
  
  'confirm fiscal calendar exists and export it
  'needed for: 06A504a; 06A504b;
  blnFiscalExists = cptCalendarExists("cptFiscalCalendar")
  If blnFiscalExists Then
    'export fiscal.csv
    lngFile = FreeFile
    strFile = strDir & "\fiscal.csv"
    Open strFile For Output As #lngFile
    Print #1, "FISCAL_END,LABEL,"
    For Each oException In ActiveProject.BaseCalendars("cptFiscalCalendar").Exceptions
      Print #1, oException.Finish & " 5:00 PM," & oException.Name & ","
    Next oException
    Close #lngFile
  End If
  
  'confirm cpt-cei.adtg exists and convert it
  'needed for: 06A504a; 06A504b;
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) <> vbNullString Then
    'copy cpt-cei.adtg to tmp dir
    FileCopy strFile, strDir & "\cpt-cei.adtg"
    
    'convert to csv for sql query...
    strFile = strDir & "\cpt-cei.adtg"
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Open strFile
    
    'clean it up
    oRecordset.Filter = "TASK_NAME LIKE '%,%'"
    If Not oRecordset.EOF Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
        oRecordset.Fields("TASK_NAME") = Replace(oRecordset.Fields("TASK_NAME"), ",", "-")
        oRecordset.MoveNext
      Loop
    End If
    oRecordset.Filter = 0
    oRecordset.Save strFile
    
    'capture field names
    strList = ""
    For lngItem = 0 To oRecordset.Fields.Count - 1
      strList = strList & oRecordset.Fields(lngItem).Name & ","
    Next lngItem
    'save as csv
    Set oFile = oFSO.CreateTextFile(strDir & "\cpt-cei.csv", True)
    oFile.Write strList & vbCrLf
    oRecordset.MoveFirst
    oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
    oRecordset.Close
    Kill strDir & "\cpt-cei.adtg"
    oFile.Close
    
  End If
  
  If Dir(strDir & "\cpt-cei.csv") <> vbNullString And blnFiscalExists Then
    Set oRecordset = CreateObject("ADODB.Recordset")
    strSQL = "SELECT "
    strSQL = strSQL & "    MAX(STATUS_DATE) "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            T1.FISCAL_END, "
    strSQL = strSQL & "            T2.STATUS_DATE "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [fiscal.csv] T1 "
    strSQL = strSQL & "            LEFT JOIN ( "
    strSQL = strSQL & "                SELECT "
    strSQL = strSQL & "                    DISTINCT STATUS_DATE "
    strSQL = strSQL & "                FROM "
    strSQL = strSQL & "                    [cpt-cei.csv] "
    strSQL = strSQL & "                WHERE "
    strSQL = strSQL & "                    PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "            ) T2 ON T2.[STATUS_DATE] = T1.FISCAL_END "
    strSQL = strSQL & "    )"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If oRecordset.EOF Then
      oRecordset.Close
      blnFiscalExists = False
      GoTo skip_fiscal
    End If
    dtCurrent = oRecordset(0)
    oRecordset.Close
    'get 2nd most recent fiscal end/status date
    strSQL = strSQL & " WHERE STATUS_DATE<#" & dtCurrent & "#"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If oRecordset.EOF Then
      oRecordset.Close
      blnFiscalExists = False
      GoTo skip_fiscal
    End If
    dtPrevious = oRecordset(0)
    oRecordset.Close
  End If
  
  '06A504a - AS changed - only if task history otherwise notify to 'use capture period'
  strFile = strDir & "\cpt-cei.csv"
  If Dir(strFile) <> vbNullString And blnFiscalExists Then
    cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A504a..."
    Application.StatusBar = "Getting Schedule Metric: 06A504a..."
    cptDECM_frm.lboMetrics.AddItem
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A504a"
    'cptDECM_frm.lboMetrics.Value = "06A505a"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Changed Actual Start"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
    DoEvents
    
    'X = Count of tasks/activities & milestones where actual start date does not equal previously reported actual start date
    'Y = Total count of tasks/activities & milestones with actual start dates
        
    'get Y
    strSQL = "SELECT TASK_UID,TASK_AS FROM [cpt-cei.csv] "
    strSQL = strSQL & "WHERE PROJECT='" & strProgramAcronym & "' AND TASK_AS IS NOT NULL "
    strSQL = strSQL & "AND STATUS_DATE = #" & dtCurrent & "#"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngY = oRecordset.RecordCount Else lngY = 1
    oRecordset.Close
    
    'get X
    strSQL = "SELECT "
    strSQL = strSQL & "    t1.TASK_UID, "
    strSQL = strSQL & "    t1.TASK_AS AS TASK_AS_IS, "
    strSQL = strSQL & "    t2.TASK_AS_WAS "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [cpt-cei.csv] AS t1 "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_AS AS TASK_AS_WAS "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "            AND TASK_AS IS NOT NULL "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtPrevious & "#"
    strSQL = strSQL & "    ) AS T2 ON T2.TASK_UID = T1.TASK_UID "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "    AND TASK_AS IS NOT NULL "
    strSQL = strSQL & "    AND STATUS_DATE = #" & dtCurrent & "# "
    strSQL = strSQL & "    AND T1.TASK_AS <> T2.TASK_AS_WAS; "
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngX = oRecordset.RecordCount Else lngX = 0
    If lngX > 0 Then
      'save results for later - todo: add to export if 06A504a.csv exists
      Set oFile = oFSO.CreateTextFile(strDir & "\06A504a.csv", True)
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oFile.Close
    End If
    oRecordset.Close
    'X/Y <= 10%
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
    'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
    cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A504a...done."
    Application.StatusBar = "Getting Schedule Metric: 06A504a...done."
    DoEvents
  End If
  
  '06A504b - AF changed - only if task history
  If Dir(strDir & "\cpt-cei.csv") <> vbNullString And blnFiscalExists Then
    cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A504b..."
    Application.StatusBar = "Getting Schedule Metric: 06A504b..."
    cptDECM_frm.lboMetrics.AddItem
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A504b"
    'cptDECM_frm.lboMetrics.Value = "06A505a"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Changed Actual Finish"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 10%"
    DoEvents
  
    'X = Count of tasks/activities & milestones where actual finish date does not equal previously reported actual finish date
    'Y = Total count of tasks/activities & milestones with actual finish dates
  
    'get Y
    strSQL = "SELECT TASK_UID,TASK_AF FROM [cpt-cei.csv] "
    strSQL = strSQL & "WHERE PROJECT='" & strProgramAcronym & "' AND TASK_AF IS NOT NULL "
    strSQL = strSQL & "AND STATUS_DATE = #" & dtCurrent & "#"
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngY = oRecordset.RecordCount Else lngY = 1
    oRecordset.Close
    
    'get X
    strSQL = "SELECT "
    strSQL = strSQL & "    t1.TASK_UID, "
    strSQL = strSQL & "    t1.TASK_AF AS TASK_AF_IS, "
    strSQL = strSQL & "    t2.TASK_AF_WAS "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    [cpt-cei.csv] AS t1 "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_AF AS TASK_AF_WAS "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "            AND TASK_AF IS NOT NULL "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtPrevious & "#"
    strSQL = strSQL & "    ) AS T2 ON T2.TASK_UID = T1.TASK_UID "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    PROJECT = '" & strProgramAcronym & "' "
    strSQL = strSQL & "    AND TASK_AF IS NOT NULL "
    strSQL = strSQL & "    AND STATUS_DATE = #" & dtCurrent & "# "
    strSQL = strSQL & "    AND T1.TASK_AF <> T2.TASK_AF_WAS; "
    oRecordset.Open strSQL, strCon, adOpenKeyset
    If Not oRecordset.EOF Then lngX = oRecordset.RecordCount Else lngX = 0
    If lngX > 0 Then
      'save results for later - todo: add to export if 06A504a.csv exists
      Set oFile = oFSO.CreateTextFile(strDir & "\06A504b.csv", True)
      oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
      oFile.Close
    End If
    oRecordset.Close
    
    'X/Y <= 10%
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
    'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
    cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A504b...done."
    Application.StatusBar = "Getting Schedule Metric: 06A504b...done."
    DoEvents
  End If
  
skip_fiscal:
  
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
  
  '06A506b - invalid forecast
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506b..."
  Application.StatusBar = "Getting Schedule Metric: 06A506b..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A506b"
  'cptDECM_frm.lboMetrics.Value = "06A506b"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Invalid Forecast"
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
  'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngX there is no Y
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
  
  'todo: allow user to refresh analysis on a one-by-one basis?
  
  '06A506c - riding status date
  If Dir(strDir & "\cpt-cei.csv") <> vbNullString And blnFiscalExists Then
    cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506c..."
    Application.StatusBar = "Getting Schedule Metric: 06A506c..."
    cptDECM_frm.lboMetrics.AddItem
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06A506c"
    'cptDECM_frm.lboMetrics.Value = "06A506b"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Riding the Status Date"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X/Y <= 1%"
    DoEvents
    
    'X = Count of incomplete tasks/activities & milestones with either forecast start or forecast finish date riding the status date
    'Y = Total count of incomplete tasks/activities & milestones
    strSQL = "SELECT UID FROM [tasks.csv] "
    strSQL = strSQL & "WHERE SUMMARY='No' AND [AF] IS NULL "
    With oRecordset
      .Open strSQL, strCon, adOpenKeyset
      lngY = oRecordset.RecordCount
      'DumpRecordsetToExcel oRecordset
      .Close
    End With
      
    'get list of incomplete UID,finish1,status1,delta1,finish2,status2,delta2 where delta2=delta1
    strSQL = "SELECT "
    strSQL = strSQL & "    P1.*, "
    strSQL = strSQL & "    P2.TASK_FINISH, "
    strSQL = strSQL & "    P2.STATUS_DATE, "
    strSQL = strSQL & "    P2.DELTA "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & "    ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_FINISH, "
    strSQL = strSQL & "            STATUS_DATE, "
    strSQL = strSQL & "            TASK_FINISH - STATUS_DATE AS [DELTA] "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = 'IBL' "
    strSQL = strSQL & "            AND IS_LOE = FALSE "
    strSQL = strSQL & "            AND TASK_AF IS NULL "
    strSQL = strSQL & "            AND TASK_FINISH >= #" & dtCurrent & "# "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtCurrent & "# "
    strSQL = strSQL & "    ) P1 "
    strSQL = strSQL & "    INNER JOIN ( "
    strSQL = strSQL & "        SELECT "
    strSQL = strSQL & "            TASK_UID, "
    strSQL = strSQL & "            TASK_FINISH, "
    strSQL = strSQL & "            STATUS_DATE, "
    strSQL = strSQL & "            TASK_FINISH - STATUS_DATE AS [DELTA] "
    strSQL = strSQL & "        FROM "
    strSQL = strSQL & "            [cpt-cei.csv] "
    strSQL = strSQL & "        WHERE "
    strSQL = strSQL & "            PROJECT = 'IBL' "
    strSQL = strSQL & "            AND IS_LOE = FALSE "
    strSQL = strSQL & "            AND TASK_AF IS NULL "
    strSQL = strSQL & "            AND TASK_FINISH >= #" & dtPrevious & "# "
    strSQL = strSQL & "            AND STATUS_DATE = #" & dtPrevious & "# "
    strSQL = strSQL & "    ) P2 ON P2.TASK_UID = P1.TASK_UID "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "    P1.DELTA = P2.DELTA "
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strSQL, strCon, adOpenKeyset
      If .EOF Then
        lngX = 0
      Else
        lngX = .RecordCount
        Set oFile = oFSO.CreateTextFile(strDir & "\06A506c-x.csv", True)
        oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
        oFile.Close
      End If
      strList = ""
      If lngX > 0 Then
        .MoveFirst
        Do While Not .EOF
          strList = strList & .Fields("TASK_UID") & ","
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
    If dblScore < 0.01 Then
      cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
    Else
      cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
    End If
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
    cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06A506c...done."
    Application.StatusBar = "Getting Schedule Metric: 06A506c...done."
    DoEvents
  End If
  
  '06I201a - SVTs todo: capture task names with "^SVT" ; allow alternative
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06I201a..."
  Application.StatusBar = "Getting Schedule Metric: 06I201a..."
  cptDECM_frm.lboMetrics.AddItem
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 0) = "06I201a"
  'cptDECM_frm.lboMetrics.Value = "06I201a"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 1) = "Schedule Visibility Tasks (SVTs)"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 2) = "X = 0"
  DoEvents
  'X = Count of incomplete tasks/activities and milestones that are not properly identified and controlled as SVTs in the IMS
  'X = 0
        
  Application.OpenUndoTransaction "cpt DECM 06I201a"
  ActiveWindow.TopPane.Activate
  FilterClear
  GroupClear
  OptionsViewEx DisplaySummaryTasks:=True
  OutlineShowAllTasks
  FilterEdit "cpt DECM Filter - 06I201a", True, True, True, , , "Active", , "equals", "Yes"
  FilterEdit "cpt DECM Filter - 06I201a", True, , , , , , "Actual Finish", "equals", "NA"
  FilterEdit "cpt DECM Filter - 06I201a", True, , , , , , "Resource Names", "does not equal", ""
  FilterEdit "cpt DECM Filter - 06I201a", True, , , , , , "Name", "contains", "SVT", , , False
  FilterApply "cpt DECM Filter - 06I201a"
  SelectAll
  Set oTasks = Nothing
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strList = ""
  If Not oTasks Is Nothing Then
    lngX = ActiveSelection.Tasks.Count
    For Each oTask In oTasks
      strList = strList & oTask.UniqueID & ","
    Next oTask
  Else
    lngX = 0
  End If
  Application.CloseUndoTransaction
  If Application.GetUndoListCount > 0 Then Application.Undo 'todo: why is this triggering a fail?
  
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 3) = lngX
  'cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY there is no Y
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 5) = lngX
  If lngX = 0 Then
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strPass
  Else
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 6) = strFail
  End If
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 7) = "todo: description"
  cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 8) = strList
  cptDECM_frm.lblStatus.Caption = "Getting Schedule Metric: 06I201a...done."
  Application.StatusBar = "Getting Schedule Metric: 06I201a...done."
  DoEvents
  
  Application.StatusBar = "DECM Scoring Complete"
  cptDECM_frm.lblStatus.Caption = "DECM Scoring Complete"
  DoEvents
  
exit_here:
  On Error Resume Next
  Set oException = Nothing
  Set oTasks = Nothing
  Set oCell = Nothing
  Set oListObject = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
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
  Dim o10A103a As Excel.Workbook
  Dim oRecordset As ADODB.Recordset
  Dim oTasks As MSProject.Tasks
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oRange As Excel.Range
  Dim oCell As Excel.Range
  'strings
  Dim strDir As String
  Dim strCon As String
  Dim strSQL As String
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
  
  strDir = Environ("tmp")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  If blnDetail Then
    With cptDECM_frm
      For lngItem = 0 To .lboMetrics.ListCount - 1
        .lboMetrics.Value = .lboMetrics.List(lngItem)
        .lboMetrics.Selected(lngItem) = True
        Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
        oWorksheet.Activate
        oWorksheet.Name = .lboMetrics.List(lngItem)
        oWorksheet.Tab.Color = 5287936
        oExcel.ActiveWindow.Zoom = 85
        If .lboMetrics.List(lngItem) = "06A101a" Then
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            'run queries
            oWorksheet.[A2].Value = "NOT IN IMS:"
            If Dir(strDir & "\wp-not-in-ims.csv") <> vbNullString Then
              Set oRecordset = CreateObject("ADODB.Recordset")
              strSQL = "SELECT * FROM [wp-not-in-ims.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
              oWorksheet.[A3].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.[C2].Value = "NOT IN EV TOOL:"
            If Dir(strDir & "\wp-not-in-ev.csv") <> vbNullString Then
              Set oRecordset = CreateObject("ADODB.Recordset")
              strSQL = "SELECT * FROM [wp-not-in-ev.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
              oWorksheet.[C3].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.Cells.Font.Name = "Calibri"
            oWorksheet.Cells.Font.Size = 11
            oWorksheet.Cells.WrapText = False
            oWorksheet.[B3].Select
            oExcel.ActiveWindow.FreezePanes = True
            oWorksheet.Columns.AutoFit
            oWorksheet.Tab.Color = 192
          End If
          GoTo next_item
        ElseIf .lboMetrics.List(lngItem) = "06A504a" Then
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            oWorksheet.[A2].Value = "Changed Actual Starts"
            If Dir(strDir & "\06A504a.csv") <> vbNullString Then
              Set oRecordset = CreateObject("ADODB.Recordset")
              strSQL = "SELECT * FROM [06A504a.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset
              oWorksheet.[A3].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.Cells.Font.Name = "Calibri"
            oWorksheet.Cells.Font.Size = 11
            oWorksheet.Cells.WrapText = False
            oWorksheet.[B4].Select
            oExcel.ActiveWindow.FreezePanes = True
            oWorksheet.Columns.AutoFit
            oWorksheet.Tab.Color = 192
          End If
          GoTo next_item
        ElseIf .lboMetrics.List(lngItem) = "06A504b" Then
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            oWorksheet.[A2].Value = "Changed Actual Finishes"
            If Dir(strDir & "\06A504b.csv") <> vbNullString Then
              Set oRecordset = CreateObject("ADODB.Recordset")
              strSQL = "SELECT * FROM [06A504b.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset
              oWorksheet.[A3].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.Cells.Font.Name = "Calibri"
            oWorksheet.Cells.Font.Size = 11
            oWorksheet.Cells.WrapText = False
            oWorksheet.[B4].Select
            oExcel.ActiveWindow.FreezePanes = True
            oWorksheet.Columns.AutoFit
            oWorksheet.Tab.Color = 192
          End If
          GoTo next_item
        ElseIf .lboMetrics.List(lngItem) = "06A506c" Then
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            oWorksheet.[A2].Value = "Riding the Status Date"
            If Dir(strDir & "\06A506c-x.csv") <> vbNullString Then
              Set oRecordset = CreateObject("ADODB.Recordset")
              strSQL = "SELECT * FROM [06A506c-x.csv]"
              oRecordset.Open strSQL, strCon, adOpenKeyset
              oWorksheet.[A3].CopyFromRecordset oRecordset
              oRecordset.Close
            End If
            oWorksheet.Cells.Font.Name = "Calibri"
            oWorksheet.Cells.Font.Size = 11
            oWorksheet.Cells.WrapText = False
            oWorksheet.[B4].Select
            oExcel.ActiveWindow.FreezePanes = True
            oWorksheet.Columns.AutoFit
            oWorksheet.Tab.Color = 192
          End If
          GoTo next_item
        ElseIf .lboMetrics.List(lngItem) = "10A103a" Then '0/100 in >1 fiscal period
          If .lboMetrics.List(lngItem, 6) = strFail Then
            On Error Resume Next
            Set o10A103a = oExcel.Workbooks(oExcel.Windows("10A103a.xlsx").Index)
            If o10A103a Is Nothing Then
              Set o10A103a = oExcel.Workbooks.Open(strDir & "\10A103a.xlsx")
            End If
            If o10A103a Is Nothing Then
              'todo: what if it doesn't exist?
            Else
              If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
              'replace current worksheet with worksheet from saved workbook
              oExcel.DisplayAlerts = False
              oWorksheet.Delete
              oExcel.DisplayAlerts = True
              o10A103a.Sheets(1).Copy After:=oWorkbook.Sheets(oWorkbook.Sheets.Count)
              o10A103a.Close True
              Set oWorksheet = oWorkbook.Sheets("10A103a")
              oWorksheet.Rows("1:2").Insert
              oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
              oWorksheet.[A2].Value = "0/100 WPs in more than 1 fiscal period"
              oWorksheet.Tab.Color = 192
            End If
          End If
          GoTo next_item
        ElseIf .lboMetrics.List(lngItem) = "10A302b" Then
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            'run query
            oWorksheet.[A2].Value = "PPs with EVP>0"
            Set oRecordset = CreateObject("ADODB.Recordset")
            strSQL = "SELECT * FROM [10A302b-x.csv]"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[C3].CopyFromRecordset oRecordset
            oRecordset.Close
            oWorksheet.Cells.Font.Name = "Calibri"
            oWorksheet.Cells.Font.Size = 11
            oWorksheet.Cells.WrapText = False
            oWorksheet.[B3].Select
            oExcel.ActiveWindow.FreezePanes = True
            oWorksheet.Columns.AutoFit
            oWorksheet.Tab.Color = 192
          End If
          GoTo next_item
        ElseIf .lboMetrics.List(lngItem) = "10A303a" Then
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
            'run query
            oWorksheet.[A2].Value = "PPs with Duration = 0"
            Set oRecordset = CreateObject("ADODB.Recordset")
            strSQL = "SELECT * FROM [10A303a-x.csv]"
            oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
            oWorksheet.[C3].CopyFromRecordset oRecordset
            oRecordset.Close
            oWorksheet.Cells.Font.Name = "Calibri"
            oWorksheet.Cells.Font.Size = 11
            oWorksheet.Cells.WrapText = False
            oWorksheet.[B3].Select
            oExcel.ActiveWindow.FreezePanes = True
            oWorksheet.Columns.AutoFit
            oWorksheet.Tab.Color = 192
          End If
          GoTo next_item
        Else
          .lboMetrics_AfterUpdate
        End If
        SelectAll
        EditCopy
        On Error Resume Next
        Set oTasks = ActiveSelection.Tasks
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If Not oTasks Is Nothing Then
          oWorksheet.Hyperlinks.Add Anchor:=oWorksheet.[A1], Address:="", SubAddress:="'DECM Dashboard'!A2", TextToDisplay:="Dashboard", ScreenTip:="Return to Dashboard"
          oWorksheet.[A2] = .lboMetrics.List(lngItem, 1)
          oWorksheet.[A3].Select
          oWorksheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:=False
          oWorksheet.Cells.Font.Name = "Calibri"
          oWorksheet.Cells.Font.Size = 11
          oWorksheet.Cells.WrapText = False
          oWorksheet.[B4].Select
          oExcel.ActiveWindow.FreezePanes = True
          oWorksheet.Columns.AutoFit
          If .lboMetrics.List(lngItem, 6) = strFail Then
            oWorksheet.Tab.Color = 192
          End If
        End If
next_item:
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
  Set o10A103a = Nothing
  Set oRecordset = Nothing
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
    
    Case "06A101a" 'WP mismatches
      'todo: do what?
    
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
    
    Case "10A103a" '0/100 >1 fiscal periods
      If Len(strList) > 0 Then
        strList = Left(strList, Len(strList) - 1) 'remove last tab
        SetAutoFilter FieldConstantToFieldName(Split(cptGetSetting("Integration", "WPCN"), "|")(0)), pjAutoFilterIn, "contains", strList 'todo: "WPCN" > "WP"
      Else
        SetAutoFilter "Name", pjAutoFilterIn, "contains", "<< zero results >>"
      End If
      
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
      If InStr(oSubproject.Path, "<>") = 0 Then 'offline
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
    If cptDECM_frm.Visible Then
      cptDECM_frm.lblProgress.Width = (lngTask / lngTasks) * cptDECM_frm.lblStatus.Width
    End If
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

Private Function cptGetEVTAnalysis() As Excel.Workbook
  'objects
  Dim oListObject As Excel.ListObject
  Dim oRange As Excel.Range
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim rst As ADODB.Recordset
  Dim oException As MSProject.Exception
  Dim oCalendar As MSProject.Calendar
  Dim oProject As MSProject.Project
  Dim oTask As MSProject.Task
  'strings
  Dim strMissingBaselines As String
  Dim strLOE As String
  Dim strLOEField As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  'longs
  Dim lngWP As Long
  Dim lngFiscalPeriodsCol As Long
  Dim lngFiscalEndCol As Long
  Dim lngLastRow As Long
  Dim lngFile As Long
  Dim lngEVT As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  'variants
  Dim vbResponse As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oProject = ActiveProject
  
  'ensure project is baselined
  If Not IsDate(oProject.BaselineSavedDate(pjBaseline)) Then
    MsgBox "This project is not yet baselined.", vbCritical + vbOKOnly, "No Baseline"
    GoTo exit_here
  End If
  
  'ensure fiscal calendar is still loaded
  If Not cptCalendarExists("cptFiscalCalendar") Then
    MsgBox "The Fiscal Calendar (cptFiscalCalendar) is missing! Please reset it and try again.", vbCritical + vbOKOnly, "What happened?"
    GoTo exit_here
  End If
      
  'export the calendar
  Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
  lngFile = FreeFile
  strFile = Environ("tmp") & "\fiscal.csv"
  Open strFile For Output As #lngFile
  Print #lngFile, "fisc_end,label,"
  For Each oException In oCalendar.Exceptions
    Print #lngFile, oException.Finish & "," & oException.Name
  Next oException
  Close #lngFile
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "10A103a"
  oWorksheet.[A1:F1] = Split("UID,WP,BLS,BLF,EVT,FiscalPeriods", ",")
  
  Set rst = CreateObject("ADODB.Recordset")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT '-' AS [UID],t1.WP,t2.BLS,t2.BLF,t1.EVT "
  strSQL = strSQL & "FROM [tasks.csv] AS t1 "
  strSQL = strSQL & "LEFT JOIN (SELECT WP,MIN(BLS) AS [BLS],MAX(BLF) AS [BLF] FROM [tasks.csv] GROUP BY WP) AS t2 ON T2.WP=T1.WP "
  strSQL = strSQL & "WHERE t1.WP IS NOT NULL "
  strSQL = strSQL & "AND t1.AF IS NULL "
  strSQL = strSQL & "AND t1.EVT='F'" 'todo: what about other ways of identifying 0/100?
  
  rst.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  oWorksheet.[A2].CopyFromRecordset rst
  rst.Close
  
  strSQL = "SELECT * FROM [fiscal.csv]"
  rst.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  oWorksheet.[H1:I1] = Split("fisc_end,label", ",")
  oWorksheet.[H2].CopyFromRecordset rst
  rst.Close
  
  Set oRange = oWorksheet.Range(oWorksheet.[A1].End(xlToRight).Offset(1, 0), oWorksheet.[A1].End(xlDown).Offset(0, 5))
  lngFiscalEndCol = oWorksheet.Rows(1).Find(what:="fisc_end").Column
  lngLastRow = oWorksheet.Cells(2, lngFiscalEndCol).End(xlDown).Row
  'Excel 2016 compatibility
  'oRange.FormulaR1C1 = "=COUNTIFS(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ","">=""&RC[-3],R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ",""<""&RC[-2])+1"
  '=SUMPRODUCT(--($G$2:$G$109>=B15)*--($G$2:$G$109<C15)*1)+1
  oRange.FormulaR1C1 = "=SUMPRODUCT(--(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ">=RC[-3])*--(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & "<RC[-2])*1)+1"
  lngFiscalPeriodsCol = oWorksheet.Rows(1).Find(what:="FiscalPeriods").Column
  oWorksheet.Columns(lngFiscalPeriodsCol).NumberFormat = "#0"
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
  oListObject.TableStyle = ""
  oWorksheet.[A1].AutoFilter
  oWorksheet.Columns.AutoFit
  If Dir(Environ("tmp") & "\10A103a.xlsx") <> vbNullString Then Kill Environ("tmp") & "\10A103a.xlsx"
  oWorkbook.SaveAs Environ("tmp") & "\10A103a.xlsx", 51
  Set cptGetEVTAnalysis = oWorkbook
  
exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  cptSpeed False
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  For lngFile = 1 To FreeFile
    Close #lngFile
  Next lngFile
  If rst.State = 1 Then rst.Close
  Set rst = Nothing
  Set oException = Nothing
  Set oCalendar = Nothing
  Set oTask = Nothing
  Set oProject = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptDECM_bas", "cptGetEVTAnalysis", Err, Erl)
  Resume exit_here
End Function
