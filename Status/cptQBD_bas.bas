Attribute VB_Name = "cptQBD_bas"
'<cpt_version>v0.0.1</cpt_version>
Option Explicit

Sub cptShowQBD_frm()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  Dim strProgramAcronym As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'get/set program acronym
  strProgramAcronym = cptGetProgramAcronym
  
  'ensure status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Please set a Status Date.", vbExclamation + vbOKOnly, "Required"
    ChangeStatusDate
    GoTo exit_here
  Else
    dtStatus = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  End If
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then 'create it
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Fields.Append "PROGRAM", adVarChar, 255
    oRecordset.Fields.Append "PROJECT_UID", adInteger
    oRecordset.Fields.Append "TASK_UID", adInteger
    oRecordset.Fields.Append "TASK_SUB_UID", adInteger
    oRecordset.Fields.Append "STEP_ORDER", adInteger
    oRecordset.Fields.Append "STEP_NAME", adVarChar, 255
    oRecordset.Fields.Append "STEP_WEIGHT", adInteger
    oRecordset.Fields.Append "STEP_AS", adDate
    oRecordset.Fields.Append "STEP_AF", adDate
    oRecordset.Fields.Append "STEP_PERCENT", adInteger 'should we force 50/50?
    oRecordset.Fields.Append "STATUS_DATE", adDate
    oRecordset.Open
    oRecordset.Save strFile, adPersistADTG
    oRecordset.Close
  End If

  With cptQBD_frm
    .lboHeader.Clear
    .lboHeader.AddItem
    .lboHeader.List(0, 0) = "#"
    .lboHeader.List(0, 1) = "NAME"
    .lboHeader.List(0, 2) = "WEIGHT"
    .lboHeader.List(0, 3) = "AS"
    .lboHeader.List(0, 4) = "AF"
    .lboHeader.List(0, 5) = "%"
  End With

  Call cptUpdateQBDForm

  cptQBD_frm.Show False

exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_bas", "cptShowQBD_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateQBDForm()
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  Dim strProgramAcronym As String
  'longs
  Dim lngSubUID As Long
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'get program acronym
  strProgramAcronym = cptGetProgramAcronym
  
  'clear the form
  With cptQBD_frm
    .Caption = "QBD - " & cptGetVersion("cptQBD_frm")
    .lboSteps.Clear
    .txtName = ""
    .txtWeight = ""
    .txtAS = ""
    .txtAF = ""
    .txtPercent = ""
    .txtWeights = 0
    .txtPerformed = 0
    .txtEV = 0
  End With
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then GoTo exit_here
  If oTasks.Count <> 1 Then GoTo exit_here
  
  cptQBD_frm.lblUID.Caption = oTasks(1).UniqueID
  
  'determine if master/sub situation
  blnMaster = False
  If ActiveProject.Subprojects.Count > 0 Then
    blnMaster = True
  End If
  
  'derive UniqueID
  If Not blnMaster Then
    lngUID = oTasks(1).UniqueID
  Else
    lngUID = cptGetSubprojectUID(oTasks(1).UniqueID)
    lngSubUID = oTasks(1).UniqueID
  End If
  
  'get file
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    If Not blnMaster Then
      .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID
    Else
      .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID & " AND TASK_SUB_UID=" & lngSubUID
    End If
    If .EOF Then
      .Filter = 0
      .Close
      GoTo exit_here
    Else
      .MoveFirst
      Do While Not .EOF
        cptQBD_frm.lboSteps.AddItem
        cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 0) = .Fields("STEP_ORDER")
        cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 1) = .Fields("STEP_NAME")
        cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 2) = .Fields("STEP_WEIGHT")
        If .Fields("STEP_AS") > 0 Then
          cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 3) = FormatDateTime(.Fields("STEP_AS"), vbShortDate)
        Else
          cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 3) = "NA"
        End If
        If .Fields("STEP_AF") > 0 Then
          cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 4) = FormatDateTime(.Fields("STEP_AF"), vbShortDate)
        Else
          cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 4) = "NA"
        End If
        cptQBD_frm.lboSteps.List(cptQBD_frm.lboSteps.ListCount - 1, 5) = .Fields("STEP_PERCENT")
        .MoveNext
      Loop
      .Filter = 0
      .Close
    End If
  End With
  
  Call cptRefreshQBDCalc
    
exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_bas", "cptUpdateQBDForm", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshQBDCalc()
  'objects
  'strings
  'longs
  Dim lngItem As Long
  Dim lngPercent As Long
  Dim lngWeights As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With cptQBD_frm
    If .lboSteps.ListCount > 0 Then
      For lngItem = 0 To .lboSteps.ListCount - 1
        lngWeights = lngWeights + .lboSteps.List(lngItem, 2)
        lngPercent = lngPercent + (.lboSteps.List(lngItem, 2) * (.lboSteps.List(lngItem, 5) / 100))
      Next lngItem
      .txtWeights.Value = lngWeights
      .txtPerformed.Value = lngPercent
      .txtEV.Value = Format(lngPercent / lngWeights, "0%")
    End If
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_bas", "cptRefreshQBDCalc", Err, Erl)
  Resume exit_here
End Sub
