VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptStatusSheet_frm 
   Caption         =   "Create Status Sheets"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "cptStatusSheet_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptStatusSheet_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.5.3</cpt_version>
Option Explicit

Private Sub cboCostTool_Change()
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'setup EVT dictionary
  Set oEVTs = CreateObject("Scripting.Dictionary")
  If Not oEVTs Is Nothing Then
    If oEVTs.Count > 0 Then oEVTs.RemoveAll
  End If
  If Not IsNull(cptStatusSheet_frm.cboCostTool.Value) Then
    If cptStatusSheet_frm.cboCostTool.Value = "COBRA" Then
      oEVTs.Add "A", "Level of Effort"
      oEVTs.Add "B", "Milestones"
      oEVTs.Add "C", "% Complete"
      oEVTs.Add "D", "Units Complete"
      oEVTs.Add "E", "50-50"
      oEVTs.Add "F", "0-100"
      oEVTs.Add "G", "100-0"
      oEVTs.Add "H", "User Defined"
      oEVTs.Add "J", "Apportioned"
      oEVTs.Add "K", "Planning Package"
      oEVTs.Add "L", "Assignment % Complete"
      oEVTs.Add "M", "Calculated Apportionment"
      oEVTs.Add "N", "Steps"
      oEVTs.Add "O", "Earned As Spent"
      oEVTs.Add "P", "% Complete Manual Entry"
    ElseIf cptStatusSheet_frm.cboCostTool.Value = "MPM" Then
      oEVTs.Add "0", "No EVM required"
      oEVTs.Add "1", "0/100"
      oEVTs.Add "2", "25/75"
      oEVTs.Add "3", "40/60"
      oEVTs.Add "4", "50/50"
      oEVTs.Add "5", "% Complete"
      oEVTs.Add "6", "LOE"
      oEVTs.Add "7", "Earned Standards"
      oEVTs.Add "8", "Milestone Weights"
      oEVTs.Add "9", "BCWP Entry"
      oEVTs.Add "A", "Apportioned"
      oEVTs.Add "P", "Milestone Weights with % Complete"
      oEVTs.Add "K", "Key Event"
    End If
  End If


exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboCostTool_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboCreate_Change()
  'objects
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Select Case Me.cboCreate
    Case 0 'A single workbook
      Me.lboItems.ForeColor = -2147483630
      Me.chkSendEmails.Caption = "Create Email"
      Me.chkLocked.Caption = "Protect Workbook"
      'Me.lblForEach.Visible = False
      Me.cboEach.Enabled = False
      Me.lboItems.Enabled = False
      Me.chkAllItems = False
      Me.chkAllItems.Enabled = False
      FilterClear
      If Not cptFilterExists("cptStatusSheet Filter") Then
        Call cptRefreshStatusTable
      End If
      
    Case 1 'A worksheet for each
      Me.lboItems.ForeColor = -2147483630
      Me.chkSendEmails.Caption = "Create Email"
      Me.chkLocked.Caption = "Protect Workheets"
      'Me.lblForEach.Visible = True
      Me.cboEach.Enabled = True
      Me.lboItems.Enabled = True
      Me.chkAllItems.Enabled = True
      If Me.Visible Then Me.cboEach.DropDown

    Case 2 'A workbook for each
      Me.lboItems.ForeColor = -2147483630
      Me.chkSendEmails.Caption = "Create Emails"
      Me.chkLocked.Caption = "Protect Workbooks"
      'Me.lblForEach.Visible = True
      Me.cboEach.Enabled = True
      Me.lboItems.Enabled = True
      Me.chkAllItems.Enabled = True
      If Me.Visible Then Me.cboEach.DropDown

    End Select
        
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboCreate_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEach_Change()
  'objects
  Dim rstItems As Object 'ADODB.Recordset
  Dim oTask As MSProject.Task
  'strings
  Dim strFieldName As String
  'longs
  Dim lngItem As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboItems.Clear
  Me.lboItems.ForeColor = -2147483630
  If Me.Visible Then
    ActiveWindow.TopPane.Activate
    FilterApply "cptStatusSheet Filter"
  End If
  
  On Error Resume Next
  lngField = FieldNameToFieldConstant(Me.cboEach)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If lngField > 0 Then
    Set rstItems = CreateObject("ADODB.Recordset")
    With rstItems
      .Fields.Append Me.cboEach.Value, 200, 255 '200=adVarChar
      .Open
      For Each oTask In ActiveProject.Tasks
        If oTask Is Nothing Then GoTo next_task
        If Not oTask.Active Then GoTo next_task
        If oTask.ExternalTask Then GoTo next_task
        If oTask.Summary Then GoTo next_task
        If IsDate(oTask.ActualFinish) Then GoTo next_task
        If Len(oTask.GetField(lngField)) > 0 Then
          If .RecordCount > 0 Then .MoveFirst
          .Find "[" & Me.cboEach.Value & "]='" & oTask.GetField(lngField) & "'"
          If .EOF Then
            .AddNew Array(0), Array(oTask.GetField(lngField))
          End If
        End If
next_task:
      Next oTask
      'validate field has items
      If .RecordCount = 0 Then
        If Len(CustomFieldGetName(lngField)) > 0 Then
          strFieldName = CustomFieldGetName(lngField)
        Else
          strFieldName = FieldConstantToFieldName(lngField)
        End If
        MsgBox "The field '" & strFieldName & "' contains no values.", vbExclamation + vbOKOnly, "Invalid Selection"
      Else
        .Sort = "[" & Me.cboEach.Value & "]"
        .MoveFirst
        Do While Not .EOF
          Me.lboItems.AddItem .Fields(0)
          .MoveNext
        Loop
      End If
      Me.txtFileName_Change
    End With
  End If 'lngField > 0
  
exit_here:
  On Error Resume Next
  If rstItems.State = 1 Then rstItems.Close
  Set rstItems = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cboEach_Change", "cboEach_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVP_AfterUpdate()
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Len(Me.cboEVP.Value) > 0 Then
    Me.lblEVP.ForeColor = -2147483630 '"Black"
  Else
    Me.lblEVP.ForeColor = 192
  End If

exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVP_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVP_Change()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptStatusSheet_frm.Visible Then Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVP_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVT_AfterUpdate()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.cboEVT.Value) > 0 Then
    Me.lblEVT.ForeColor = -2147483630 '"Black"
  Else
    Me.lblEVT.ForeColor = 192
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVT_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub cboEVT_Change()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptStatusSheet_frm.Visible Then Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cboEVT_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub chkAllItems_Click()
  Dim lngItem As Long
  Dim strCriteria As String
  Dim strFieldName As String
  
  If IsNull(Me.cboEach) Or Me.lboItems.ListCount = 0 Then Exit Sub
  strFieldName = Me.cboEach.Value
  If Me.chkAllItems Then
    For lngItem = 0 To Me.lboItems.ListCount - 1
      Me.lboItems.Selected(lngItem) = True
      strCriteria = strCriteria & Me.lboItems.List(lngItem) & Chr$(9)
    Next lngItem
    If Len(strCriteria) = 0 Then Exit Sub
    strCriteria = Left(strCriteria, Len(strCriteria) - 1)
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
  Else
    If Me.ActiveControl.Name = Me.chkAllItems.Name Then
      For lngItem = 0 To Me.lboItems.ListCount - 1
        Me.lboItems.Selected(lngItem) = False
      Next lngItem
      SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterClear
    End If
  End If
  
End Sub

Private Sub chkAppendStatusDate_Click()
  If Me.chkAppendStatusDate Then
    Me.lblDirSample.Caption = Me.txtDir & Format(CDate(Me.txtStatusDate), "yyyy-mm-dd") & "\"
  Else
    Me.lblDirSample.Caption = Me.txtDir
  End If
End Sub

Private Sub chkContour_Click()
  cptRefreshStatusTable
End Sub

Private Sub chkHide_Click()

  If Not Me.Visible Then GoTo exit_here
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.txtHideCompleteBefore.Enabled = Me.chkHide
  If Me.Visible Then Call cptRefreshStatusTable
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("chkHide_Click", "chkHide_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub chkKeepOpen_Click()
  If Me.chkKeepOpen Then
    Me.chkSendEmails = False
    Me.chkSendEmails.Enabled = False
  Else
    Me.chkSendEmails.Enabled = True
  End If
End Sub

Private Sub chkSendEmails_Click()
  Dim strQuickPart As String
  Dim blnExists As Boolean
  Dim lngItem As Long

  Me.txtSubject.Enabled = Me.chkSendEmails
  Me.txtCC.Enabled = Me.chkSendEmails
  Me.cboQuickParts.Enabled = Me.chkSendEmails
  If Me.chkSendEmails Then
    Me.chkKeepOpen = False
    Me.chkKeepOpen.Enabled = False
    Me.lblEmailHints.Visible = True
    Call cptListQuickParts(True)
    strQuickPart = cptGetSetting("StatusSheet", "cboQuickPart")
    If Len(strQuickPart) > 0 Then
      blnExists = False
      For lngItem = 0 To Me.cboQuickParts.ListCount - 1
        If Me.cboQuickParts.List(lngItem, 0) = strQuickPart Then
          Me.cboQuickParts.Value = strQuickPart
          blnExists = True
          Exit For
        End If
      Next lngItem
      If Not blnExists Then
        MsgBox "QuickPart '" & strQuickPart & "' not found.", vbExclamation + vbOKOnly, "Stored Setting Invalid"
      End If
    End If
  Else
    Me.chkKeepOpen.Enabled = True
    Me.lblEmailHints.Visible = False
  End If

End Sub

Sub cmdAdd_Click()
  Dim lngField As Long, lngExport As Long, lngExists As Long
  Dim blnExists As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngField = 0 To Me.lboFields.ListCount - 1
    If Me.lboFields.Selected(lngField) Then
      'ensure doesn't already exist
      blnExists = False
      For lngExists = 0 To Me.lboExport.ListCount - 1
        If Me.lboExport.List(lngExists, 0) = Me.lboFields.List(lngField) Then
          GoTo next_item
        End If
      Next lngExists
      Me.lboExport.AddItem
      lngExport = Me.lboExport.ListCount - 1
      Me.lboExport.List(lngExport, 0) = Me.lboFields.List(lngField, 0)
      Me.lboExport.List(lngExport, 1) = Me.lboFields.List(lngField, 1)
      Me.lboExport.List(lngExport, 2) = Me.lboFields.List(lngField, 2)
    End If
next_item:
  Next lngField

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdAdd_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdAddAll_Click()
  Dim lngField As Long, lngExport As Long, lngExists As Long
  Dim blnExists As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngField = 0 To Me.lboFields.ListCount - 1
    'ensure doesn't already exist
    blnExists = False
    For lngExists = 0 To Me.lboExport.ListCount - 1
      If Me.lboExport.List(lngExists, 0) = Me.lboFields.List(lngField) Then
        GoTo next_item
      End If
    Next lngExists
    Me.lboExport.AddItem
    lngExport = Me.lboExport.ListCount - 1
    Me.lboExport.List(lngExport, 0) = Me.lboFields.List(lngField, 0)
    Me.lboExport.List(lngExport, 1) = Me.lboFields.List(lngField, 1)
    Me.lboExport.List(lngExport, 2) = Me.lboFields.List(lngField, 2)
next_item:
  Next lngField

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdAddAll_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdCancel_Click()
  Dim strFileName As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  If Not oEVTs Is Nothing Then oEVTs.RemoveAll
  Set oEVTs = Nothing
  Unload Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdCancel_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdDir_Click()
  'objects
  Dim oShell As Object
  Dim oFileDialog As Object 'FileDialog
  Dim oExcel As Excel.Application
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oExcel = CreateObject("Excel.Application")
  Set oFileDialog = oExcel.FileDialog(msoFileDialogFolderPicker)
  With oFileDialog
    .AllowMultiSelect = False
    If Left(ActiveProject.Path, 2) = "<>" Or Left(ActiveProject.Path, 4) = "http" Then 'server project: default to Desktop
      Set oShell = CreateObject("WScript.Shell")
      .InitialFileName = oShell.SpecialFolders("Desktop")
    Else 'not a server project
      .InitialFileName = ActiveProject.Path
    End If
    If .Show Then
      Me.txtDir = .SelectedItems(1) & "\" & IIf(Me.chkAppendStatusDate, Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "\", "")
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set oShell = Nothing
  Set oFileDialog = Nothing
  Set oExcel = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdDir_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdDown_Click()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If lngExport < Me.lboExport.ListCount - 1 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport + 1, 0)
        strField = Me.lboExport.List(lngExport + 1, 1)
        strField2 = Me.lboExport.List(lngExport + 1, 2)
        'move selected values
        Me.lboExport.List(lngExport + 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport + 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport + 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport + 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport

  If blnSelected Then Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("frmStatusSeet", "cmdDown_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdRemove_Click()
  Dim lngExport As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If Me.lboExport.Selected(lngExport) Then
      Me.lboExport.RemoveItem lngExport
    End If
  Next lngExport

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRemove_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdRemoveAll_Click()
  Dim lngExport As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  For lngExport = Me.lboExport.ListCount - 1 To 0 Step -1
    Me.lboExport.RemoveItem lngExport
  Next lngExport

  Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRemoveAll_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdRun_Click()
  'objects
  'strings
  'longs
  Dim lngDateFormat As Long
  Dim lngSelectedItems As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnError As Boolean
  Dim blnIncluded As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  lngDateFormat = Application.DefaultDateFormat
  blnError = False

  '-2147483630 = Black
  Me.lblStatusDate.ForeColor = -2147483630
  Me.lblEVT.ForeColor = -2147483630
  Me.lblEVP.ForeColor = -2147483630
  Me.chkHide.ForeColor = -2147483630
  Me.lblStatus.ForeColor = -2147483630
  Me.cboCostTool.ForeColor = -2147483630
  Me.cboCreate.ForeColor = -2147483630
  Me.cboEach.BorderColor = -2147483642
  Me.lblDirectory.ForeColor = -2147483630
  Me.txtDir.BorderColor = -2147483642
  Me.lblNamingConvention.ForeColor = -2147483630

  'validation
  If Not IsDate(Me.txtStatusDate.Value) Then
    Me.lblStatusDate.ForeColor = 192  'Red
    blnError = True
  ElseIf IsDate(Me.txtStatusDate.Value) Then
    If CDate(Me.txtStatusDate.Value) < #1/1/1984# Then
      Me.lblStatusDate.ForeColor = 192  'Red
      blnError = True
    End If
  End If
  If Me.chkHide.Value = True Then
    If Not IsDate(Me.txtHideCompleteBefore.Value) Then
      Me.chkHide.ForeColor = 192  'Red
      blnError = True
    ElseIf IsDate(Me.txtHideCompleteBefore.Value) Then
      If CDate(Me.txtHideCompleteBefore.Value) < #1/1/1984# Then
        Me.chkHide.ForeColor = 192
        blnError = True
      End If
    End If
  End If
  If Len(Me.cboCostTool.Value) = 0 Then
    Me.lblCostTool.ForeColor = 192 'Red
    blnError = True
  End If
  'hide complete before must be prior to or equal to status date
  If IsDate(Me.txtStatusDate.Value) And IsDate(Me.txtHideCompleteBefore.Value) Then
    If CDate(Me.txtHideCompleteBefore.Value) > CDate(Me.txtStatusDate.Value) Then
      MsgBox "'Hide Complete Before' date must be prior to, or equal to, status date.", vbExclamation + vbOKOnly, "Invalid Hide Complete Before Date"
      Me.chkHide.ForeColor = 192
      blnError = True
    End If
  End If
  If Len(Me.cboEVT.Value) = 0 Then
    Me.lblEVT.ForeColor = 192 'Red
    blnError = True
  End If
  If Len(Me.cboEVP.Value) = 0 Then
    Me.lblEVP.ForeColor = 192 'Red
    blnError = True
  End If
  'ensure unique filenames
  If Me.cboCreate.Value = "0" Then 'one workbook
    If InStr(Me.txtFileName, "[item]") > 0 Then
      Me.lblNamingConvention.ForeColor = 192 'red
      MsgBox "Cannot use '[item]' in naming convention when creating a single workbook.", vbExclamation + vbOKOnly, "Invalid Naming Convention"
      blnError = True
    End If
  ElseIf Me.cboCreate.Value = "1" Then 'worksheet for each
    If InStr(Me.txtFileName, "[item]") > 0 Then
      Me.lblNamingConvention.ForeColor = 192 'red
      MsgBox "Cannot use '[item]' in naming convention when creating worksheet for each.", vbExclamation + vbOKOnly, "Invalid Naming Convention"
      blnError = True
    End If
  ElseIf Me.cboCreate.Value = "2" Then 'workbook for each
    If InStr(Me.txtFileName, "[item]") = 0 Then
      Me.lblNamingConvention.ForeColor = 192 'red
      MsgBox "Must include '[item]' in naming convention when creating workbook for each.", vbExclamation + vbOKOnly, "Invalid Naming Convention"
      blnError = True
    End If
  End If
  If Me.cboCreate.Value <> "0" Then
    'a limiting field must be selected
    If Me.cboEach.Value = 0 Then
      Me.cboEach.BorderColor = 192
      blnError = True
    End If
    'at least one item selected
    For lngItem = 0 To Me.lboItems.ListCount - 1
      If Me.lboItems.Selected(lngItem) Then lngSelectedItems = lngSelectedItems + 1
    Next lngItem
    If lngSelectedItems = 0 Then
      Me.lboItems.Selected(0) = True
      'Me.lboItems.ForeColor = 92
      blnError = True
    End If
    'the limiting field should be included in the export list
    blnIncluded = False
    For lngItem = 0 To Me.lboExport.ListCount - 1
      If Me.lboExport.List(lngItem, 1) = Me.cboEach Then blnIncluded = True
    Next lngItem
    If Not blnIncluded Then
      If MsgBox("The For Each field '" & Me.cboEach & "' is not included in the export list." & vbCrLf & vbCrLf & "Include it?", vbYesNo + vbQuestion, "Include For Each Field?") = vbYes Then
        For lngItem = 0 To Me.lboFields.ListCount - 1
          Me.lboFields.Selected(lngItem) = Me.lboFields.List(lngItem, 1) = Me.cboEach
        Next lngItem
        Me.cmdAdd_Click
      End If
    End If
  End If
  If Dir(Me.txtDir, vbDirectory) = vbNullString Then
    Me.lblDirectory.ForeColor = 192
    Me.txtDir.BorderColor = 192
    blnError = True
  End If
  'prevent duplication of EVT and EV%
  If Me.lboExport.ListCount > 0 Then
    For lngItem = Me.lboExport.ListCount - 1 To 0 Step -1
      If Me.lboExport.List(lngItem, 1) = Me.cboEVP.Value Or Me.lboExport.List(lngItem, 1) = Me.cboEVT.Value Then
        MsgBox "'" & Me.lboExport.List(lngItem, 1) & "' is included by default; removing from export list.", vbInformation + vbOKOnly, "Duplicate"
        Me.lboExport.RemoveItem lngItem
      End If
    Next lngItem
  End If
  'todo: ensure notes column title is unique in the columns
  If blnError Then
    Me.lblStatus.ForeColor = 192 'red
    Me.lblStatus.Caption = " Please complete all required fields."
  Else
    'save settings
    Call cptSaveStatusSheetSettings
    'create the sheet
    Application.DefaultDateFormat = pjDate_mm_dd_yyyy
    Call cptCreateStatusSheet
  End If

exit_here:
  On Error Resume Next
  Application.DefaultDateFormat = lngDateFormat
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdRun_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdUp_Click()
  Dim lngExport As Long
  Dim lngField As Long, strField As String, strField2 As String
  Dim blnSelected As Boolean
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  blnSelected = False
  For lngExport = 0 To Me.lboExport.ListCount - 1
    If lngExport > 0 Then
      If Me.lboExport.Selected(lngExport) Then
        blnSelected = True
        'capture values
        lngField = Me.lboExport.List(lngExport - 1, 0)
        strField = Me.lboExport.List(lngExport - 1, 1)
        strField2 = Me.lboExport.List(lngExport - 1, 2)
        'move selected values
        Me.lboExport.List(lngExport - 1, 0) = Me.lboExport.List(lngExport, 0)
        Me.lboExport.List(lngExport - 1, 1) = Me.lboExport.List(lngExport, 1)
        Me.lboExport.List(lngExport - 1, 2) = Me.lboExport.List(lngExport, 2)
        Me.lboExport.Selected(lngExport - 1) = True
        Me.lboExport.List(lngExport, 0) = lngField
        Me.lboExport.List(lngExport, 1) = strField
        Me.lboExport.List(lngExport, 2) = strField2
        Me.lboExport.Selected(lngExport) = False
      End If
    End If
  Next lngExport
  
  If blnSelected Then Call cptRefreshStatusTable

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "cmdUp_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub lblEmailHints_Click()
  Dim strHints As String

  strHints = "The following fields are available for auto replacement in the subject line and your Email Template (a.k.a., 'Quick Part'):" & vbCrLf & vbCrLf
  strHints = strHints & "[STATUS_DATE] > Status Date in MM/DD/YYYY format" & vbCrLf
  'strHints = strHints & "[YYYYMM] > Status Date in MM/DD/YYYY format" & vbCrLf
  strHints = strHints & "[PROGRAM] > Program Acronym" & vbCrLf
  strHints = strHints & vbCrLf & "Send other suggestions to cpt@ClearPlanConsulting.com"
  MsgBox strHints, vbInformation + vbOKOnly, "Email Hints"

End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboItems_Change()
  'strings
  Dim strCriteria As String
  Dim strFieldName As String
  'longs
  Dim lngSelectedItems As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not Me.Visible Then Exit Sub
  If Me.ActiveControl.Name <> Me.lboItems.Name Then Exit Sub
  
  If Application.Calculation = pjAutomatic Then cptSpeed True
  
  strFieldName = Me.cboEach.Value

  For lngItem = 0 To Me.lboItems.ListCount - 1
    If Me.lboItems.Selected(lngItem) Then
      strCriteria = strCriteria & Me.lboItems.List(lngItem) & Chr$(9)
      lngSelectedItems = lngSelectedItems + 1
    End If
  Next lngItem
  
  If Len(strCriteria) = 0 Then
    FilterClear
    FilterApply "cptStatusSheet Filter"
  Else
    FilterClear
    FilterApply "cptStatusSheet Filter"
    strCriteria = Left(strCriteria, Len(strCriteria) - 1)
    SetAutoFilter FieldName:=strFieldName, FilterType:=pjAutoFilterIn, Criteria1:=strCriteria
  End If
  
  If Me.Visible Then
    If Me.ActiveControl.Name = Me.lboItems.Name Then
      Me.chkAllItems.Value = lngSelectedItems = Me.lboItems.ListCount
    End If
  End If
  
  ActiveWindow.TopPane.Activate
  SelectBeginning
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lboItems_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub stxtSearch_Change()
  Dim lngItem As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFields.Clear
  
  With CreateObject("ADODB.REcordset")
    .Open cptDir & "\settings\cpt-status-sheet-search.adtg"
    If Len(Me.stxtSearch.Text) > 0 Then
      .Filter = "[Custom Field Name] LIKE '*" & cptRemoveIllegalCharacters(Me.stxtSearch.Text) & "*'"
    Else
      .Filter = 0
    End If
    If .RecordCount > 0 Then .MoveFirst
    lngItem = 0
    Do While Not .EOF
      Me.lboFields.AddItem
      Me.lboFields.List(lngItem, 0) = .Fields(0)
      Me.lboFields.List(lngItem, 1) = .Fields(1)
      Me.lboFields.List(lngItem, 2) = .Fields(2)
      .MoveNext
      lngItem = lngItem + 1
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "stxtSearch_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub stxtSearch_Enter()
  Dim lngField As Long, strFileName As String
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = cptDir & "\settings\cpt-status-sheet-search.adtg"
  If Dir(strFileName) <> vbNullString Then Exit Sub
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", 200, 100 '200=adVarChar
    .Fields.Append "Custom Field Name", 200, 100 '200=adVarChar
    .Fields.Append "Local Field Name", 200, 100 '200=adVarChar
    .Open
    For lngField = 0 To cptStatusSheet_frm.lboFields.ListCount - 1
      .AddNew Array(0, 1, 2), Array(Me.lboFields.List(lngField, 0), cptStatusSheet_frm.lboFields.List(lngField, 1), cptStatusSheet_frm.lboFields.List(lngField, 2))
    Next lngField
    .Update
    .Save strFileName
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "stxtSearch_Enter", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtDir_Change()
  Dim strDir As String
  
  strDir = Me.txtDir.Text
  If InStr(strDir, "[yyyy-mm-dd]") > 0 Then
    strDir = Replace(strDir, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
  End If
  If Right(strDir, 1) <> "\" Then
    strDir = strDir & "\"
  End If
  Me.lblDirSample.Caption = strDir

  If Dir(strDir, vbDirectory) = vbNullString Then
    Me.lblDirectory.ForeColor = 192
  Else
    Me.lblDirectory.ForeColor = -2147483630
  End If

End Sub

Sub txtFileName_Change()
  Dim strFileName As String
  strFileName = Me.txtFileName.Text
  If InStr(strFileName, "[yyyy-mm-dd]") > 0 Then
    strFileName = Replace(strFileName, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
  End If
  If InStr(strFileName, "[Program]") > 0 Then
    strFileName = Replace(strFileName, "[Program]", cptGetProgramAcronym)
  End If
  If Me.cboCreate.Value > 0 Then 'for each
    If InStr(strFileName, "[item]") > 0 Then
      If Me.lboItems.ListCount > 0 Then
        Me.lblFileNameSample.Caption = Replace(strFileName, "[item]", Me.lboItems.List(0, 0)) & ".xlsx"
      Else
        Me.lblFileNameSample.Caption = "< no item found >"
      End If
    Else
      Me.lblFileNameSample.Caption = strFileName & ".xlsx"
    End If
  Else
    Me.lblFileNameSample.Caption = strFileName & ".xlsx"
  End If
End Sub

Private Sub txtHideCompleteBefore_Change()
  Dim stxt As String
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not Me.Visible Then GoTo exit_here
  stxt = cptRegEx(Me.txtHideCompleteBefore.Text, "[0-9\/]*")
  Me.txtHideCompleteBefore.Text = stxt
  If Len(Me.txtHideCompleteBefore.Text) > 0 Then
    If IsDate(Me.txtHideCompleteBefore.Text) Then
      If CDate(Me.txtHideCompleteBefore.Text) > #1/1/1984# Then
        Me.chkHide.ForeColor = -2147483630 '"Black"
      Else
        Me.chkHide.ForeColor = 192 'red
      End If
    Else
      Me.chkHide.ForeColor = 192 'red
    End If
  End If

exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtHideCompleteBefore", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtStatusDate_Change()
  Dim stxt As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not Me.Visible Then GoTo exit_here
  stxt = cptRegEx(Me.txtStatusDate.Text, "[0-9\/]*")
  Me.txtStatusDate.Text = stxt
  If Len(Me.txtStatusDate.Text) > 0 Then
    If IsDate(Me.txtStatusDate.Text) Then
      If CDate(Me.txtStatusDate.Text) > #1/1/1984# Then
        Me.lblStatusDate.ForeColor = -2147483630 '"Black"
      Else
        Me.lblStatusDate.ForeColor = 192 'red
      End If
    Else
      Me.lblStatusDate.ForeColor = 192 'red
    End If
  End If
  
exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "txtStatusDate_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtSubject_Change()
  Dim strSubject As String
  strSubject = Me.txtSubject.Text
  strSubject = Replace(strSubject, "[yyyy-mm-dd]", Format(ActiveProject.StatusDate, "yyyy-mm-dd"))
  strSubject = Replace(strSubject, "[Program]", cptGetProgramAcronym)
  If Me.cboCreate > 0 And Me.lboItems.ListCount > 0 Then
    strSubject = Replace(strSubject, "[item]", Me.lboItems.List(0, 0))
  End If
  Me.lblSubjectPreview.Caption = strSubject
  
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Dim dtStatus As Date, lngDiff As Long
  If Me.ActiveControl.Name = "txtHideCompleteBefore" Then Exit Sub
  If IsDate(ActiveProject.StatusDate) Then
    dtStatus = FormatDateTime(CDate(Me.txtStatusDate), vbShortDate)
    If dtStatus <> CDate(ActiveProject.StatusDate) Then
      lngDiff = VBA.DateDiff("d", CDate(Me.txtHideCompleteBefore), CDate(Me.txtStatusDate))
      Me.txtStatusDate = FormatDateTime(ActiveProject.StatusDate, vbShortDate)
      Me.txtHideCompleteBefore = DateAdd("d", -lngDiff, ActiveProject.StatusDate)
    End If
  Else
    cptStatusSheet_frm.txtStatusDate.Value = FormatDateTime(DateAdd("d", 6 - Weekday(Now), Now), vbShortDate)
  End If
End Sub
