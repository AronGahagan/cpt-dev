VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFilterByClipboard_frm 
   Caption         =   "Filter By Clipboard"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "cptFilterByClipboard_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFilterByClipboard_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub chkFilter_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
End Sub

Private Sub cmdClear_Click()
  Me.txtFilter = ""
  Me.txtFilter.Visible = True
  Me.lboFilter.Visible = False
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_frm", "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboFilter_Click()
  'objects
  Dim oTask As Task
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtGoTo As Date

  On Error Resume Next
  If Me.optUID Then
    Set oTask = ActiveProject.Tasks.UniqueID(Me.lboFilter.Value)
  ElseIf Me.optID Then
    Set oTask = ActiveProject.Tasks.Item(CLng(Me.lboFilter.Value))
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTask Is Nothing Then
    If IsDate(oTask.Stop) Then
      dtGoTo = oTask.Stop
    Else
      dtGoTo = oTask.Start
    End If
    If ActiveWindow.ActivePane <> ActiveWindow.TopPane Then ActiveWindow.TopPane.Activate
    On Error Resume Next
    If Not EditGoTo(oTask.ID, dtGoTo) Then
      OptionsViewEx displaysummarytasks:=True
      OutlineShowAllTasks
      EditGoTo oTask.ID, dtGoTo
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "lboFilter_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub optID_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
End Sub

Private Sub optUID_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
End Sub

Private Sub tglEdit_Click()
  If Me.tglEdit Then
    Me.txtFilter.Visible = True
    Me.lboFilter.Visible = False
  Else
    Me.txtFilter.Visible = False
    Me.lboFilter.Visible = True
  End If
  Me.txtFilter.Height = Me.lboFilter.Height
End Sub

Private Sub txtFilter_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
  'objects
  'strings
  Dim strFilter As String
  Dim strItem As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim strNewList As Variant
  Dim vList As Variant
  'dates

  'scrub the incoming data

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  strFilter = Data.GetText
  If InStr(strFilter, vbTab) > 0 Then
    vList = Split(strFilter, vbCrLf)
  ElseIf InStr(strFilter, ",") > 0 Then
    vList = Split(strFilter, ",")
  ElseIf InStr(strFilter, ";") > 0 Then
    vList = Split(strFilter, ";")
  Else
    vList = Array(strFilter)
  End If
  
  If IsEmpty(vList) Then GoTo exit_here
  
  For lngItem = 0 To UBound(vList)
    strItem = cptRegEx(CStr(vList(lngItem)), "[0-9]*")
    If Len(strItem) > 0 Then
      strNewList = strNewList & CLng(strItem) & ","
    End If
  Next lngItem
  Cancel = True
  Effect = fmDropEffectNone
  If Len(strNewList) > 0 Then Me.txtFilter = strNewList
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "txtFilter_BeforeDropOrPaste", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtFilter_Change()
  Call cptUpdateClipboard
End Sub

Private Sub txtFilter_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  Call cptClipboardJump
End Sub

Private Sub txtFilter_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call cptClipboardJump
End Sub
