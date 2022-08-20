VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptQBD_frm 
   Caption         =   "QBD"
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975.001
   OleObjectBlob   =   "cptQBD_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptQBD_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.1</cpt_version>
Option Explicit

Private Sub cmdAddStep_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  Dim strProgramAcronym As String
  'longs
  Dim lngStepNumber As Long
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strProgramAcronym = cptGetProgramAcronym
  lngUID = CLng(Me.lblUID.Caption)
  lngStepNumber = Me.lboSteps.ListCount + 1
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please reopen the form and try again.", vbCritical + vbOKOnly, "Error"
    Me.Hide
    GoTo exit_here
  End If
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID & " AND STEP_NAME='{step name}'"
    If .EOF Then
      .Filter = 0
      .AddNew Array("PROGRAM", "TASK_UID", "STEP_ORDER", "STEP_NAME", "STEP_WEIGHT"), Array(strProgramAcronym, lngUID, lngStepNumber, "{step name}", 10)
      .Save strFile
      Call cptUpdateQBDForm
    Else
      lngStepNumber = .Fields("STEP_ORDER")
      .Filter = 0
      .Close
    End If
  End With
  
  Me.lboSteps.Value = lngStepNumber
  Me.lboSteps_AfterUpdate
  Me.txtName.SetFocus
  Me.txtName.SelStart = 0
  Me.txtName.SelLength = Me.txtName.TextLength

  
exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "cmdAddStep_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCapture_Click()
  'objects
  'strings
  'longs
  Dim lngEV As Long
  Dim lngUID As Long
  Dim lngEVPField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Me.lblUID.Caption = "" Then GoTo exit_here
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required"
      GoTo exit_here
    End If
  End If
  If CLng(Me.lblUID.Caption) > 0 Then
    lngUID = CLng(Me.lblUID.Caption)
    lngEV = CLng(Replace(Me.txtEV.Value, "%", ""))
    lngEVPField = cptGetSetting("Metrics", "cboEVP")
    OpenUndoTransaction "Update EV% on UID " & lngUID & " to " & lngEV & "%"
    ActiveProject.Tasks.UniqueID(lngUID).SetField lngEVPField, lngEV
    CloseUndoTransaction
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "cmdCapture_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdDeleteStep_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strProgramAcronym As String
  Dim strFile As String
  'longs
  Dim lngItem As Long
  Dim lngStepNumber As Long
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboSteps.Value) Then
    strProgramAcronym = cptGetProgramAcronym
    lngUID = CLng(Me.lblUID.Caption)
    lngStepNumber = CLng(Me.lboSteps.Value)
    strFile = cptDir & "\settings\cpt-qbd.adtg"
    If Dir(strFile) = vbNullString Then
      MsgBox "Please reopen the form and try again.", vbCritical + vbOKOnly, "Error"
      Me.Hide
      GoTo exit_here
    End If
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strFile, , adOpenKeyset
      'find and delete
      .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID & " AND STEP_ORDER=" & lngStepNumber
      If Not .EOF Then .Delete adAffectCurrent
      'now renumber
      .Filter = 0
      .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID
      If Not .EOF Then
        .MoveFirst
        Do While Not .EOF
          If .Fields("STEP_ORDER") > lngStepNumber Then
            .Fields("STEP_ORDER") = .Fields("STEP_ORDER") - 1
          End If
          .MoveNext
        Loop
        .Filter = 0
        .Save strFile
        Call cptUpdateQBDForm
      End If
    End With
  End If
  
  If Me.lboSteps.ListCount > 0 Then
    Me.lboSteps.Value = 1
    Me.lboSteps_AfterUpdate
  End If
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "cmdDeleteStep_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdDown_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  Dim lngStepNumber As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboSteps.Value) Then
    lngStepNumber = CLng(Me.lboSteps.Value)
    If lngStepNumber = 1 Then GoTo exit_here
    strFile = cptDir & "\settings\cpt-qbd.adtg"
    If Dir(strFile) = vbNullString Then
      MsgBox "Please reopen the form and try again.", vbCritical + vbOKOnly, "Error"
      Me.Hide
      GoTo exit_here
    End If
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      oRecordset.Open strFile, , adOpenKeyset
      oRecordset.Filter = "PROGRAM='" & cptGetProgramAcronym & "' AND TASK_UID = " & CLng(Me.lblUID.Caption)
      If Not .EOF Then
        .MoveFirst
        Do While Not .EOF
          If .Fields("STEP_ORDER") = lngStepNumber - 1 Then
            .Fields("STEP_ORDER") = .Fields("STEP_ORDER") + 1
            .Update
          ElseIf .Fields("STEP_ORDER") = lngStepNumber Then
            .Fields("STEP_ORDER") = .Fields("STEP_ORDER") - 1
            .Update
          End If
          .MoveNext
        Loop
        .Filter = 0
        .Sort = "PROGRAM,TASK_UID,STEP_ORDER"
        .Save strFile, adPersistADTG
      End If
    End With
    
    Call cptUpdateQBDForm
  
    Me.lboSteps.Value = lngStepNumber - 1
    Call cptQBD_frm.lboSteps_AfterUpdate
    
  End If

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "cmdDown_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdExport_Click()
  'objects
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  'longs
  Dim lngLastRow As Long
  Dim lngItem As Long
  Dim lngResponse As Long
  'integers
  'doubles
  'booleans
  Dim blnLimit As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngResponse = MsgBox("Limit to this task? If you click 'No' then all project QBDs will be exported.", vbQuestion + vbYesNoCancel)
  Select Case lngResponse
    Case vbYes
      blnLimit = True
    Case vbNo
      blnLimit = False
    Case vbCancel
      GoTo exit_here
  End Select
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  With oExcel.ActiveWindow
    .Zoom = 85
    .SplitRow = 1
    .SplitColumn = 0
    .FreezePanes = True
  End With
  Set oWorksheet = oWorkbook.Sheets(1)
  
  If blnLimit Then
    oWorksheet.[A1] = "UID"
    oWorksheet.Range(oWorksheet.[B1], oWorksheet.[B1].Offset(0, Me.lboHeader.ColumnCount)) = Me.lboHeader.List
    oWorksheet.Range(oWorksheet.Cells(2, 2), oWorksheet.Cells(1 + Me.lboSteps.ListCount, Me.lboSteps.ColumnCount + 1)) = Me.lboSteps.List
    lngLastRow = oWorksheet.[B1048576].End(xlUp).Row
    oWorksheet.Range(oWorksheet.Cells(2, 1), oWorksheet.Cells(lngLastRow, 1)) = CLng(Me.lblUID.Caption)
    oWorksheet.Columns.AutoFit
  Else
    'todo: how would you want to see it?
    'todo: Workbook per stakeholder, worksheet per WPCN
  End If
  
  
  
exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "cmdExport_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdUp_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  Dim lngStepNumber As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboSteps.Value) Then
    lngStepNumber = CLng(Me.lboSteps.Value)
    If lngStepNumber = Me.lboSteps.ListCount Then GoTo exit_here
    strFile = cptDir & "\settings\cpt-qbd.adtg"
    If Dir(strFile) = vbNullString Then
      MsgBox "Please reopen the form and try again.", vbCritical + vbOKOnly, "Error"
      Me.Hide
      GoTo exit_here
    End If
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      oRecordset.Open strFile, , adOpenKeyset
      oRecordset.Filter = "PROGRAM='" & cptGetProgramAcronym & "' AND TASK_UID = " & CLng(Me.lblUID.Caption)
      If Not .EOF Then
        .MoveFirst
        Do While Not .EOF
          If .Fields("STEP_ORDER") = lngStepNumber Then
            .Fields("STEP_ORDER") = .Fields("STEP_ORDER") + 1
            .Update
          ElseIf .Fields("STEP_ORDER") = lngStepNumber + 1 Then
            .Fields("STEP_ORDER") = .Fields("STEP_ORDER") - 1
            .Update
          End If
          .MoveNext
        Loop
        .Filter = 0
        .Sort = "PROGRAM,TASK_UID,STEP_ORDER"
        .Save strFile, adPersistADTG
      End If
    End With
    
    Call cptUpdateQBDForm
  
    Me.lboSteps.Value = lngStepNumber + 1
    Call cptQBD_frm.lboSteps_AfterUpdate

  End If

exit_here:
  On Error Resume Next
  If oRecordset.State Then oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "cmdDown_Click", Err, Erl)
  Resume exit_here
End Sub

Sub lboSteps_AfterUpdate()
  If Not IsNull(Me.lboSteps.Value) Then
    Me.txtName = Me.lboSteps.List(Me.lboSteps.ListIndex, 1)
    Me.txtWeight = Me.lboSteps.List(Me.lboSteps.ListIndex, 2)
    Me.txtAS.Enabled = True
    Me.txtAS = Me.lboSteps.List(Me.lboSteps.ListIndex, 3)
    Me.txtAF.Enabled = True
    Me.txtAF = Me.lboSteps.List(Me.lboSteps.ListIndex, 4)
    Me.txtPercent.Enabled = True
    Me.txtPercent = Me.lboSteps.List(Me.lboSteps.ListIndex, 5)
    cptRefreshQBDCalc
  End If
End Sub

Private Sub txtAF_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strProgramAcronym As String
  Dim strFile As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.txtAF.BorderColor = 8421504
  Me.Repaint

  If Me.txtName = "{step name}" Then GoTo exit_here

  If IsNull(Me.lboSteps.Value) Then GoTo exit_here
  
  If Me.txtAF.Text = Me.lboSteps.List(Me.lboSteps.ListIndex, 4) Then GoTo exit_here
  
  If Not IsDate(Me.txtAF.Text) Then
    Me.txtAF.BorderColor = 192
    Me.Repaint
    GoTo exit_here
  End If
  
  If CDate(Me.txtAF.Text) > Now Then
    Me.txtAF.BorderColor = 192
    Me.Repaint
    GoTo exit_here
  End If
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please reopen the form and try again.", vbCritical + vbOKOnly, "Error"
    Me.Hide
    GoTo exit_here
  End If
  
  strProgramAcronym = cptGetProgramAcronym
  lngUID = CLng(Me.lblUID.Caption)
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID & " AND STEP_ORDER=" & Me.lboSteps.Value
    If Not .EOF Then
      .Update "STEP_AF", FormatDateTime(Me.txtAF.Text, vbShortDate)
      Me.lboSteps.List(Me.lboSteps.ListIndex, 4) = FormatDateTime(Me.txtAF.Text, vbShortDate)
    End If
    .Filter = 0
    .Save strFile, adPersistADTG
    .Close
  End With

exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "txtAF_Change", Err, Erl)
  Resume exit_here
End Sub


Private Sub txtAS_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strProgramAcronym As String
  Dim strFile As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.txtAS.BorderColor = 8421504
  Me.Repaint

  If Me.txtName = "{step name}" Then GoTo exit_here

  If IsNull(Me.lboSteps.Value) Then GoTo exit_here
  
  If Me.txtAS.Text = Me.lboSteps.List(Me.lboSteps.ListIndex, 3) Then GoTo exit_here
  
  If Not IsDate(Me.txtAS.Text) Then
    Me.txtAS.BorderColor = 192
    Me.Repaint
    GoTo exit_here
  End If
  
  If CDate(Me.txtAS.Text) > Now Then
    Me.txtAS.BorderColor = 192
    Me.Repaint
    GoTo exit_here
  End If
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please reopend the form and try again.", vbCritical + vbOKOnly, "Error"
    Me.Hide
    GoTo exit_here
  End If
  
  strProgramAcronym = cptGetProgramAcronym
  lngUID = CLng(Me.lblUID.Caption)
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    .Filter = "PROGRAM='" & strProgramAcronym & "' AND TASK_UID=" & lngUID & " AND STEP_ORDER=" & Me.lboSteps.Value
    If Not .EOF Then
      .Update "STEP_AS", FormatDateTime(Me.txtAS.Text, vbShortDate)
      Me.lboSteps.List(Me.lboSteps.ListIndex, 3) = FormatDateTime(Me.txtAS.Text, vbShortDate)
    End If
    .Filter = 0
    .Save strFile, adPersistADTG
    .Close
  End With

exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "txtAS_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtName_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Me.txtName.BorderColor = 8421504
  Me.Repaint

  If IsNull(Me.lboSteps.Value) Then GoTo exit_here
  
  If Me.txtName.Text = "{step name}" Then GoTo exit_here
  
  If Me.txtName.Text = Me.lboSteps.List(Me.lboSteps.ListIndex, 1) Then GoTo exit_here
  
  If Len(Me.txtName) = 0 Then
    Me.txtName.BorderColor = 192
    Me.Repaint
    GoTo exit_here
  End If
  
  If InStr(Me.txtName.Text, Chr(39)) > 0 Then
    Me.txtName.Text = Replace(Me.txtName.Text, Chr(39), "")
  End If
  If InStr(Me.txtName.Text, Chr(34)) > 0 Then
    Me.txtName.Text = Replace(Me.txtName.Text, Chr(34), "")
  End If
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please reopend the form and try again.", vbCritical + vbOKOnly, "Error"
    Me.Hide
    GoTo err_here
  End If
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    .Filter = "PROGRAM='" & cptGetProgramAcronym & "' AND TASK_UID=" & CLng(Me.lblUID.Caption) & " AND STEP_ORDER=" & Me.lboSteps.Value
    If Not .EOF Then
      .Update "STEP_NAME", Me.txtName.Text
      Me.lboSteps.List(Me.lboSteps.ListIndex, 1) = Me.txtName.Text
    End If
    .Filter = 0
    .Save strFile, adPersistADTG
  End With
    
exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "txtName_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtPercent_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.txtPercent.BorderColor = 8421504
  Me.Repaint
  
  If IsNull(Me.lboSteps.Value) Then GoTo exit_here
  
  If Me.txtName.Text = "{step name}" Then GoTo exit_here
  
  If Me.txtPercent.Text = Me.lboSteps.List(Me.lboSteps.ListIndex, 5) Then GoTo exit_here
  
  Me.txtPercent.Text = cptRegEx(Me.txtPercent.Text, "[0-9]{1,}")
  If Len(Me.txtPercent.Text) = 0 Then
    Me.txtPercent.BorderColor = 192
    Me.Repaint
    Exit Sub
  End If
  If CLng(Me.txtPercent.Text) > 100 Then
    Me.txtPercent.BorderColor = 192
    Me.Repaint
    Exit Sub
  End If
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please reopend the form and try again.", vbCritical + vbOKOnly, "Error"
    Me.Hide
    GoTo err_here
  End If

  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    .Filter = "PROGRAM='" & cptGetProgramAcronym & "' AND TASK_UID=" & CLng(Me.lblUID.Caption) & " AND STEP_ORDER=" & Me.lboSteps.Value
    If Not .EOF Then
      .Update "STEP_PERCENT", CLng(Me.txtPercent.Text)
      Me.lboSteps.List(Me.lboSteps.ListIndex, 5) = Me.txtPercent.Text
    End If
    .Filter = 0
    .Save strFile, adPersistADTG
  End With

  Call cptRefreshQBDCalc

exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "txtPercent_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtWeight_Change()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile  As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.txtWeight.BorderColor = 8421504
  Me.Repaint
  
  If IsNull(Me.lboSteps.Value) Then GoTo exit_here
  
  If Me.txtName.Text = "{step name}" Then GoTo exit_here
  
  If Me.txtWeight.Text = Me.lboSteps.List(Me.lboSteps.ListIndex, 2) Then GoTo exit_here
  
  Me.txtWeight.Text = cptRegEx(Me.txtWeight.Text, "[0-9]{1,}")
  If Len(Me.txtWeight.Text) = 0 Then
    Me.txtWeight.BorderColor = 192
    Me.Repaint
    Exit Sub
  End If
  If CLng(Me.txtWeight.Text) = 0 Then
    Me.txtWeight.BorderColor = 192
    Me.Repaint
    Exit Sub
  End If
  
  strFile = cptDir & "\settings\cpt-qbd.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please reopend the form and try again.", vbCritical + vbOKOnly, "Error"
    Me.Hide
    GoTo err_here
  End If
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile, , adOpenKeyset
    .Filter = "PROGRAM='" & cptGetProgramAcronym & "' AND TASK_UID=" & CLng(Me.lblUID.Caption) & " AND STEP_ORDER=" & Me.lboSteps.Value
    If Not .EOF Then
      .Update "STEP_WEIGHT", CLng(Me.txtWeight.Text)
      Me.lboSteps.List(Me.lboSteps.ListIndex, 2) = Me.txtWeight.Text
    End If
    .Filter = 0
    .Save strFile, adPersistADTG
  End With
  
  Call cptRefreshQBDCalc
  
exit_here:
  On Error Resume Next
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQBD_frm", "txtWeight_Change", Err, Erl)
  Resume exit_here
  
End Sub
