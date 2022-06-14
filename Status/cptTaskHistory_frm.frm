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
  Call cptHandleErr("cptTaskHistory_frm", "lblURL", err, Erl)
  Resume exit_here
End Sub

Private Sub lboTaskHistory_Click()
  If IsNull(Me.lboTaskHistory.Value) Then Exit Sub
  Call cptGetTaskHistoryNote(CDate(Me.lboTaskHistory.Value), CLng(Me.lblUID.Caption))
End Sub

Private Sub txtVariance_Change()
  If Me.ActiveControl.Name <> "txtVariance" Then Exit Sub
  If IsNull(Me.lboTaskHistory.Value) Then
    MsgBox "Please select a Status Date.", vbExclamation + vbOKOnly, "Hold on"
    Exit Sub
  Else
    Call cptUpdateTaskHistoryNote(CLng(Me.lblUID.Caption), Me.lboTaskHistory.Value, Me.txtVariance.Text)
  End If
End Sub
