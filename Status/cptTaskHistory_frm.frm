VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptTaskHistory_frm 
   Caption         =   "Task History"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550.001
   OleObjectBlob   =   "cptTaskHistory_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptTaskHistory_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdExport_Click()
  If IsNumeric(CLng(Me.lblUID.Caption)) Then
    Call cptExportTaskHistory(CLng(Me.lblUID.Caption))
  End If
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboTaskHistory_Click()
  If IsNull(Me.lboTaskHistory.Value) Then Exit Sub
  If Me.ActiveControl.Name <> "lboTaskHistory" Then Exit Sub
  Call cptGetTaskHistoryNote(CDate(Me.lboTaskHistory.Value), CLng(Me.lblUID.Caption))
End Sub

Private Sub optAllHistory_Click()
  If Me.optAllHistory Then
    Me.optAllHistory.Value = False
    Me.tglExport = False
    Call cptExportTaskHistory
  End If
End Sub

Private Sub optCurrentNotes_Click()
  If Me.optCurrentNotes Then
    If IsDate(ActiveProject.StatusDate) Then
      Me.lblWarning.Visible = False
      Me.tglExport.Value = False
      Call cptExportTaskHistory(blnNotesOnly:=True)
    Else
      Me.optCurrentNotes.Value = False
      Me.lblWarning.Caption = "No Status Date."
      Me.lblWarning.Visible = True
    End If
  End If
End Sub

Private Sub optTaskHistory_Click()
  If Me.optTaskHistory Then
    If IsNumeric(Me.lblUID) Then
      Me.lblWarning.Visible = False
      Me.tglExport.Value = False
      Call cptExportTaskHistory(lngUID:=CLng(Me.lblUID))
    Else
      Me.optTaskHistory.Value = False
      Me.lblWarning.Caption = "No task selected."
      Me.lblWarning.Visible = True
    End If
  End If
End Sub

Sub tglExport_Click()
  Me.lblWarning.Visible = False
  If tglExport Then
    Me.txtVariance.Width = 252
    Me.optAllHistory.Value = False
    Me.optCurrentNotes.Value = False
    Me.optTaskHistory.Value = False
    Me.OptionButton4.Value = False
    Me.optAllHistory.Visible = True
    Me.optCurrentNotes.Visible = True
    Me.optTaskHistory.Visible = True
    Me.OptionButton4.Visible = True
  Else
    Me.optAllHistory.Visible = False
    Me.optCurrentNotes.Visible = False
    Me.optTaskHistory.Visible = False
    Me.OptionButton4.Visible = False
    Me.txtVariance.Width = 414
  End If
End Sub

Private Sub txtVariance_Change()
  If Me.ActiveControl.Name <> "txtVariance" Then Exit Sub
  If IsNull(Me.lboTaskHistory.Value) Then
    Me.lblWarning.Caption = "Please select a Status Date."
    Me.lblWarning.Visible = True
    Me.txtVariance.Text = ""
    Exit Sub
  Else
    Me.lblWarning.Visible = False
    Call cptUpdateTaskHistoryNote(CLng(Me.lblUID.Caption), Me.lboTaskHistory.Value, Me.txtVariance.Text)
  End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call cptCore_bas.cptStartEvents
End Sub
