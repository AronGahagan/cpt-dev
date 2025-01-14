VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSmartDuration_frm 
   Caption         =   "Smart Duration"
   ClientHeight    =   1590
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3750
   OleObjectBlob   =   "cptSmartDuration_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSmartDuration_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v2.0.2</cpt_version>
Public dateError As Boolean
Public finDate As Date
Public StartDate As Date
Public lngUID As Long

Private Sub cmdApply_Click()
  Dim oTask As MSProject.Task
  Dim dtStart As Date
  Dim lngDelta As Long
  
  If finDate = 0 Then Exit Sub
  Set oTask = ActiveProject.Tasks.UniqueID(Me.lngUID)
  
  If oTask.Milestone Or oTask.Duration = 0 Then
    If MsgBox("Proceed with editing a zero-duration milestone?", vbQuestion + vbYesNo, "Please confirm") = vbNo Then
      GoTo exit_here
    End If
  End If
  'todo: should we assume 5 PM finish for elapsed durations?
  'todo: ...maybe yes, to make TS calcs a little cleaner?
  If Len(cptRegEx(CStr(finDate), "(AM|PM)")) = 0 Then
    finDate = CDate(finDate & " 5:00 PM")
  End If
  
  If IsDate(oTask.Resume) Then dtStart = oTask.Resume Else dtStart = oTask.Start
  OpenUndoTransaction "Smart Duration"
  If cptRegEx(oTask.DurationText, "[A-z]") = "e" Then
    oTask.RemainingDuration = oTask.RemainingDuration + VBA.DateDiff("n", oTask.Finish, Me.finDate)
  Else
    If oTask.Calendar = "None" Or oTask.Calendar = ActiveProject.Calendar Then
      If oTask.Finish > Me.finDate Then
        oTask.RemainingDuration = oTask.RemainingDuration - Application.DateDifference(Me.finDate, oTask.Finish)
      ElseIf oTask.Finish < Me.finDate Then
        oTask.RemainingDuration = oTask.RemainingDuration + Application.DateDifference(oTask.Finish, Me.finDate)
      End If
    Else
      If oTask.Finish > Me.finDate Then
        oTask.RemainingDuration = oTask.RemainingDuration - Application.DateDifference(Me.finDate, oTask.Finish, oTask.Calendar)
      ElseIf oTask.Finish < Me.finDate Then
        oTask.RemainingDuration = oTask.RemainingDuration + Application.DateDifference(oTask.Finish, Me.finDate, oTask.Calendar)
      End If
    End If
  End If
  CloseUndoTransaction
  cptSaveSetting "SmartDuration", "chkKeepOpen", IIf(Me.chkKeepOpen, "1", "0")
  If Not Me.chkKeepOpen Then Me.Hide
exit_here:
  Set oTask = Nothing
End Sub

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub txtTargetFinish_Change()
  'limit entry to numbers and /
  Me.txtTargetFinish.Text = cptRegEx(Me.txtTargetFinish.Text, "[0-9\/]{1,}")
  'limit to a dates only
  If Not IsDate(Me.txtTargetFinish.Text) Then
    Me.txtTargetFinish.BorderColor = 192
    Me.lblWeekday.Caption = "-"
    Me.cmdApply.Enabled = False
    Me.Repaint
  Else
    'limit to dates after the start date
    If CDate(FormatDateTime(Me.txtTargetFinish.Text, vbShortDate) & " 5:00 PM") < ActiveProject.Tasks.UniqueID(Me.lngUID).Start Then
      Me.txtTargetFinish.BorderColor = 192
      Me.lblWeekday.Caption = "-"
      Me.cmdApply.Enabled = False
      Me.Repaint
    Else
      Me.finDate = CDate(Me.txtTargetFinish.Text)
      Me.txtTargetFinish.BorderColor = -2147483642
      Me.lblWeekday.Caption = Format(CDate(Me.txtTargetFinish.Text), "dddd")
      Me.cmdApply.Enabled = True
      Me.Repaint
    End If
  End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.Hide
    Cancel = True
  End If
End Sub

Private Sub UserForm_Terminate()
  Unload Me
End Sub
