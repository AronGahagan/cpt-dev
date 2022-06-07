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
'<cpt_version>v1.4.0</cpt_version>
Option Explicit

Private Sub cmdApply_Click()
'objects
Dim oTask As Task
'strings
'longs
Dim lngItem As Long
'integers
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Are you sure?", vbYesNo + vbExclamation, "Please confirm bulk action.") = vbNo Then GoTo exit_here
  
  Application.OpenUndoTransaction "Advanced Text Action"
  For lngItem = 0 To Me.lboOutput.ListCount - 1
    If IsNull(cptText_frm.lboOutput.List(lngItem, 0)) Then GoTo exit_here
    On Error Resume Next
    Set oTask = ActiveProject.Tasks.UniqueID(Me.lboOutput.List(lngItem, 0))
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oTask Is Nothing Then
      If MsgBox("UID " & cptText_frm.lboOutput.List(lngItem, 0) & " not found in Project: '" & ActiveProject.Name & "'! Proceed?", vbCritical + vbYesNo, "Task Not Found") = vbNo Then
        Err.Clear
        GoTo exit_here
      Else
        GoTo next_item
      End If
    End If
    oTask.Name = Me.lboOutput.List(lngItem, 1)
next_item:
  Next lngItem

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Call cptStartEvents
  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "cmdApply_Click()", Err, Erl)
  Resume exit_here

End Sub

Private Sub cmdClear_Click()
Dim lngItem As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  Call cptUpdatePreview

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "cmdClear_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "lblURL_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtAppend_Change()
Dim lngItem As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtAppend.Text) > 0 Then
    Call cptUpdatePreview(strAppend:=Me.txtAppend.Text)
  Else
    Call cptUpdatePreview
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
  Call cptHandleErr("cptText_frm", "txtAppend_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtCharacters_Change()
Dim strCharacters As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure clng
  If Len(Me.txtCharacters.Text) > 0 Then
    strCharacters = cptRegEx(Me.txtCharacters.Text, "[0-9]*")
    Me.txtCharacters.Text = strCharacters
    Me.chkIsDirty = True
    If Len(strCharacters) > 0 Then
      Call cptUpdatePreview(lngCharacters:=CLng(strCharacters))
    Else
      Call cptUpdatePreview
    End If
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtCharacters_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtCountBy_Change()
Dim strCountBy As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtCountBy.Text) > 0 Then
    strCountBy = cptRegEx(Me.txtCountBy.Text, "[0-9]*")
    Me.txtCountBy.Text = strCountBy
    Me.chkIsDirty = True
    If Len(strCountBy) > 0 Then
      Call cptUpdatePreview(lngCountBy:=CLng(strCountBy))
    Else
      Call cptUpdatePreview
    End If
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtCountBy_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtPrefix_Change()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtPrefix.Text) > 0 Then
    Call cptUpdatePreview(strPrefix:=Me.txtPrefix.Text)
    Me.chkIsDirty = True
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtPrefix_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtPrepend_Change()
Dim lngItem As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Call cptUpdatePreview(strPrepend:=Me.txtPrepend.Text)
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
  Call cptHandleErr("cptText_frm", "txtPrepend_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtReplaceWhat_Change()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtReplaceWhat.Text) > 0 Then
    Call cptUpdatePreview(strReplaceWhat:=Me.txtReplaceWhat.Text, strReplaceWith:=Me.txtReplaceWith)
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtReplaceWhat_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtReplaceWith_Change()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtReplaceWith.Text) > 0 Then
    Call cptUpdatePreview(strReplaceWhat:=Me.txtReplaceWhat, strReplaceWith:=Me.txtReplaceWith.Text)
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtReplaceWith_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtStartAt_Change()
Dim strStartAt As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtStartAt.Text) > 0 Then
    strStartAt = cptRegEx(Me.txtStartAt.Text, "[0-9]*")
    Me.txtStartAt.Text = strStartAt
    If Len(strStartAt) > 0 Then
      Call cptUpdatePreview(lngStartAt:=CLng(strStartAt))
    Else
      Call cptUpdatePreview
    End If
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtStartAt_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub txtSuffix_Change()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Len(Me.txtSuffix.Text) > 0 Then
    Call cptUpdatePreview(strSuffix:=Me.txtSuffix.Text)
  Else
    Call cptUpdatePreview
  End If
  Me.chkIsDirty = CheckDirty

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_frm", "txtSuffix_Change", Err, Erl)
  Resume exit_here
  
End Sub

Public Function CheckDirty() As Boolean
Dim blnDirty As Boolean, ctl As control

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  Call cptHandleErr("cptText_frm", "CheckDirty", Err, Erl)
  Resume exit_here
  
End Function

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call cptCore_bas.cptStartEvents
End Sub
