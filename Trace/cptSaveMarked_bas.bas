Attribute VB_Name = "cptSaveMarked_bas"
'<cpt_version>v1.0.6</cpt_version>
Option Explicit

Sub cptShowSaveMarked_frm()
  'objects
  'strings
  Dim strApplyFilter As String
  Dim strProgram As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Call cptUpdateMarked
  cptSaveMarked_frm.Caption = "Import Marked (" & cptGetVersion("cptSaveMarked_frm") & ")"
  strApplyFilter = cptGetSetting("SaveMarked", "chkApplyFilter")
  If Len(strApplyFilter) > 0 Then
    cptSaveMarked_frm.chkApplyFilter = CBool(strApplyFilter)
  Else
    cptSaveMarked_frm.chkApplyFilter = False
  End If
  strProgram = cptGetProgramAcronym
  If Len(strProgram) > 0 Then
    cptSaveMarked_frm.cboProjects.AddItem strProgram
    cptSaveMarked_frm.cboProjects.Value = strProgram
    cptSaveMarked_frm.cboProjects.Locked = True
    cptSaveMarked_frm.cboProjects.Enabled = False
  End If
  cptSaveMarked_frm.Show False

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_bas", "cptShowSaveMarked_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateMarked(Optional strFilter As String)
  'objects
  Dim rstMarked As Object 'ADODB.Recordset 'Object
  'strings
  Dim strProject As String
  Dim strMarked As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strProject = cptGetProgramAcronym
  If Len(strProject) = 0 Then
    MsgBox "Program Acronym is required for this feature.", vbExclamation + vbOKOnly, "Program Acronym Needed"
    GoTo exit_here
  End If
  
  'clear listboxes and reset headers
  With cptSaveMarked_frm
    .cboProjects.Clear
    .lboMarked.Clear
    .lboMarked.AddItem
    .lboMarked.List(.lboMarked.ListCount - 1, 0) = "TIMESTAMP"
    .lboMarked.List(.lboMarked.ListCount - 1, 1) = "PROJECT"
    .lboMarked.List(.lboMarked.ListCount - 1, 2) = "DESCRIPTION"
    .lboMarked.List(.lboMarked.ListCount - 1, 3) = "COUNT"
    .lboDetails.Clear
    .lboDetails.AddItem
    .lboDetails.List(.lboDetails.ListCount - 1, 0) = "UID"
    .lboDetails.List(.lboDetails.ListCount - 1, 1) = "TASK"
  End With
  
  'get list of marked sets
  'todo: filter for where PROJECT=cptGetProgramAcronym?
  strMarked = cptDir & "\cpt-marked.adtg"
  If Dir(strMarked) = vbNullString Then
    MsgBox "No marked tasks saved.", vbCritical + vbOKOnly, "Nada"
    GoTo exit_here
  End If
  Set rstMarked = CreateObject("ADODB.Recordset")
  With rstMarked
    .Open strMarked
    .Sort = "TSTAMP DESC"
    If Len(strFilter) > 0 Then
      .Filter = "DESCRIPTION Like '*" & strFilter & "*' AND PROJECT_ID='" & strProject & "'"
    Else
      .Filter = "PROJECT_ID='" & strProject & "'"
    End If
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        With cptSaveMarked_frm
          .lboMarked.AddItem
          .lboMarked.List(.lboMarked.ListCount - 1, 0) = rstMarked(1)
          .lboMarked.List(.lboMarked.ListCount - 1, 1) = rstMarked(2)
          .lboMarked.List(.lboMarked.ListCount - 1, 2) = rstMarked(3)
        End With
        .MoveNext
      Loop
    End If
    .Filter = 0
    .Close
    
    'get marked task count
    strMarked = cptDir & "\cpt-marked-details.adtg"
    rstMarked.Open strMarked
    With cptSaveMarked_frm
      For lngItem = 1 To .lboMarked.ListCount - 1
        rstMarked.Filter = "TSTAMP=#" & CDate(.lboMarked.List(lngItem, 0)) & "#"
        .lboMarked.List(lngItem, 3) = rstMarked.RecordCount
        rstMarked.Filter = 0
      Next lngItem
    End With
  End With

exit_here:
  On Error Resume Next
  If rstMarked.State = 1 Then rstMarked.Close
  Set rstMarked = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveMarked_bas", "cptUpdateMarked", Err, Erl)
  Resume exit_here
End Sub

Sub cptSaveMarked()
  'objects
  Dim oTask As Task 'Object
  Dim rstMarked As Object 'ADODB.Recordset 'Object
  'strings
  Dim strProject As String
  Dim strGUID As String
  Dim strDescription As String
  Dim strMarked As String
  'longs
  Dim lngSelected As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtTimestamp As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set rstMarked = CreateObject("ADODB.Recordset")
  strMarked = cptDir & "\cpt-marked.adtg"
  If Dir(strMarked) = vbNullString Then
    rstMarked.Fields.Append "GUID", adGUID
    rstMarked.Fields.Append "TSTAMP", adDBTimeStamp
    rstMarked.Fields.Append "PROJECT_ID", adVarChar, 255
    rstMarked.Fields.Append "Description", adVarChar, 255
    rstMarked.Open
    rstMarked.Save strMarked, adPersistADTG
    rstMarked.Close
  End If
  If rstMarked.State <> 1 Then rstMarked.Open strMarked
  
  strProject = cptGetProgramAcronym
  If Len(strProject) = 0 Then
    MsgBox "You must set a program acronym to use this feature.", vbCritical + vbOKOnly, "Program Acronym Needed"
    GoTo exit_here
  End If

  strDescription = InputBox("Describe this capture:", "Save Marked")
  If Len(strDescription) = 0 Then
    MsgBox "No description; nothing saved.", vbExclamation + vbOKOnly
    GoTo exit_here
  End If
  dtTimestamp = Now()
  rstMarked.AddNew Array(1, 2, 3), Array(dtTimestamp, strProject, strDescription)
  rstMarked.Update
  rstMarked.Save
  rstMarked.Close
  
  Set rstMarked = CreateObject("ADODB.Recordset")
  strMarked = cptDir & "\cpt-marked-details.adtg"
  If Dir(strMarked) = vbNullString Then
    rstMarked.Fields.Append "TSTAMP", adDBTimeStamp
    rstMarked.Fields.Append "UID", adInteger
    rstMarked.Open
    rstMarked.Save strMarked, adPersistADTG
  End If
  If rstMarked.State <> 1 Then rstMarked.Open strMarked
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Marked Then
      rstMarked.AddNew Array(0, 1), Array(dtTimestamp, oTask.UniqueID)
      rstMarked.Update
    End If
next_task:
  Next oTask
  rstMarked.Save
  rstMarked.Close
  
  dtTimestamp = 0
  If cptSaveMarked_frm.Visible Then
    If Not IsNull(cptSaveMarked_frm.lboMarked.Value) Then dtTimestamp = cptSaveMarked_frm.lboMarked.Value
    cptUpdateMarked
    If dtTimestamp > 0 Then cptSaveMarked_frm.lboMarked.Value = dtTimestamp
  End If

exit_here:
  On Error Resume Next
  Set oTask = Nothing
  If rstMarked.State = 1 Then rstMarked.Close
  Set rstMarked = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptSaveMarked", Err, Erl)
  Resume exit_here
End Sub
