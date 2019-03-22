VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptText_frm 
   Caption         =   "Text Tools"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "cptText_frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptText_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdApply_Click()
'objects
Dim Task As Object
'strings
'longs
Dim lngItem As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Are you sure?", vbYesNo + vbExclamation, "Please confirm bulk action.") = vbNo Then GoTo exit_here
  
  Application.OpenUndoTransaction "Advanced Text Action"
  For lngItem = 0 To Me.lboOutput.ListCount - 1
    On Error Resume Next
    Set Task = ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0))
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Task Is Nothing Then
      If MsgBox("UID " & cptText_frm.lboOutput.List(lngItem, 0) & " not found in Project: '" & ActiveProject.Name & "'! Proceed?", vbCritical + vbYesNo, "Task Not Found") = vbNo Then
        err.Clear
        GoTo exit_here
      Else
        GoTo next_item
      End If
    End If
    Task.Name = Me.lboOutput.List(lngItem, 1)
next_item:
  Next lngItem

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set Task = Nothing
  Call StartEvents
  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "cmdApply_Click()", err)
  Resume exit_here

End Sub

Private Sub cmdClear_Click()
Dim lngItem As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.txtPrepend.Value = ""
  Me.txtAppend.Value = ""
  Me.txtPrefix.Value = ""
  Me.txtCharacters.Value = ""
  Me.txtStartAt.Value = ""
  Me.txtCountBy.Value = ""
  Me.txtSuffix.Value = ""
  Me.txtReplaceWhat.Value = ""
  Me.txtReplaceWith.Value = ""
  Me.chkIsDirty = False
  For lngItem = 0 To Me.lboOutput.ListCount - 1
    Me.lboOutput.List(lngItem, 1) = ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0)).Name
  Next
  Call UpdatePreview

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "cmdClear_Click", err)
  Resume exit_here
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If InternetIsConnected Then Application.OpenBrowser "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "lblURL_Click", err)
  Resume exit_here
End Sub

Private Sub lblWakeUp_Click()
  Call StartEvents
End Sub

Private Sub txtAppend_Change()
Dim lngItem As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtAppend.Text) > 0 Then
    Call UpdatePreview(strAppend:=Me.txtAppend.Text)
  Else
    Call UpdatePreview
  End If
  Exit Sub
  
  If Len(Me.txtAppend.Text) > 0 Then
    For lngItem = 0 To Me.lboOutput.ListCount - 1
      Me.lboOutput.List(lngItem, 1) = ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0)).Name & " " & Trim(Me.txtAppend.Text)
    Next lngItem
  Else
    For lngItem = 0 To Me.lboOutput.ListCount - 1
      Me.lboOutput.List(lngItem, 1) = ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0)).Name
    Next lngItem
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtAppend_Change", err)
  Resume exit_here
End Sub

Private Sub txtCharacters_Change()
Dim strCharacters As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ensure clng
  If Len(Me.txtCharacters.Text) > 0 Then
    strCharacters = RegEx(Me.txtCharacters.Text, "[0-9]*")
    Me.txtCharacters.Text = strCharacters
    Me.chkIsDirty = True
    If Len(strCharacters) > 0 Then
      Call UpdatePreview(lgCharacters:=CLng(strCharacters))
    Else
      Call UpdatePreview
    End If
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtCharacters_Change", err)
  Resume exit_here

End Sub

Private Sub txtCountBy_Change()
Dim strCountBy As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtCountBy.Text) > 0 Then
    strCountBy = RegEx(Me.txtCountBy.Text, "[0-9]*")
    Me.txtCountBy.Text = strCountBy
    Me.chkIsDirty = True
    If Len(strCountBy) > 0 Then
      Call UpdatePreview(lgCountBy:=CLng(strCountBy))
    Else
      Call UpdatePreview
    End If
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("frmTextToolls", "txtCountBy_Change", err)
  Resume exit_here
End Sub

Private Sub txtPrefix_Change()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtPrefix.Text) > 0 Then
    Call UpdatePreview(strPrefix:=Me.txtPrefix.Text)
    Me.chkIsDirty = True
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtPrefix_Change", err)
  Resume exit_here
  
End Sub

Private Sub txtPrepend_Change()
Dim lngItem As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Call UpdatePreview(strPrepend:=Me.txtPrepend.Text)
  Exit Sub

  If Len(Me.txtPrepend.Text) > 0 Then
    For lngItem = 0 To Me.lboOutput.ListCount - 1
      Me.lboOutput.List(lngItem, 1) = Trim(Me.txtPrepend.Text) & " " & ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0)).Name
    Next lngItem
  Else
    For lngItem = 0 To Me.lboOutput.ListCount - 1
      Me.lboOutput.List(lngItem, 1) = ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0)).Name
    Next lngItem
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtPrepend_Change", err)
  Resume exit_here

End Sub

Private Sub txtReplaceWhat_Change()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtReplaceWhat.Text) > 0 Then
    Call UpdatePreview(strReplaceWhat:=Me.txtReplaceWhat.Text, strReplaceWith:=Me.txtReplaceWith)
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtReplaceWhat_Change", err)
  Resume exit_here
End Sub

Private Sub txtReplaceWith_Change()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtReplaceWith.Text) > 0 Then
    Call UpdatePreview(strReplaceWhat:=Me.txtReplaceWhat, strReplaceWith:=Me.txtReplaceWith.Text)
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtReplaceWith_Change", err)
  Resume exit_here
  
End Sub

Private Sub txtStartAt_Change()
Dim strStartAt As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtStartAt.Text) > 0 Then
    strStartAt = RegEx(Me.txtStartAt.Text, "[0-9]*")
    Me.txtStartAt.Text = strStartAt
    If Len(strStartAt) > 0 Then
      Call UpdatePreview(lgStartAt:=CLng(strStartAt))
    Else
      Call UpdatePreview
    End If
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtStartAt_Change", err)
  Resume exit_here
  
End Sub

Private Sub txtSuffix_Change()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtSuffix.Text) > 0 Then
    Call UpdatePreview(strSuffix:=Me.txtSuffix.Text)
  Else
    Call UpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptText_frm", "txtSuffix_Change", err)
  Resume exit_here
  
End Sub

Public Function CheckDirty() As Boolean
Dim blnDirty As Boolean, ctl As control

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnDirty = False
  For Each ctl In Me.Frame2.Controls
    If ctl.Tag = "getsDirty" Then
      If Len(ctl.Text) > 0 Or Len(ctl.Value) > 0 Then
        blnDirty = True
        Exit For
      End If
    End If
  Next ctl
  CheckDirty = blnDirty

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call HandleErr("cptText_frm", "CheckDirty", err)
  Resume exit_here
  
End Function
