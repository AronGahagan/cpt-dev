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
'<cpt_version>v2.0.0</cpt_version>
Public dateError As Boolean
Public finDate As Date
Public StartDate As Date
Public lngUID As Long

Private Sub cmdApply_Click()
  If finDate = 0 Then Exit Sub
  If Len(cptRegEx(CStr(finDate), "(AM|PM)")) = 0 Then
    finDate = CDate(finDate & " 5:00 PM")
  End If
  Dim oTask As MSProject.Task
  Set oTask = ActiveProject.Tasks.UniqueID(Me.lngUID)
  OpenUndoTransaction "Smart Duration"
  If Left(cptRegEx(oTask.DurationText, "[A-z]{1,}"), 1) = "e" Then
    oTask.Duration = VBA.DateDiff("n", oTask.Start, Me.finDate)
  Else
    If oTask.Calendar = "None" Or oTask.Calendar = ActiveProject.Calendar Then
      oTask.Duration = Application.DateDifference(oTask.Start, Me.finDate)
    Else
      oTask.Duration = Application.DateDifference(oTask.Start, Me.finDate, oTask.Calendar)
    End If
  End If
  CloseUndoTransaction
  cptSaveSetting "SmartDuration", "chkKeepOpen", IIf(Me.chkKeepOpen, "1", "0")
  If Not Me.chkKeepOpen Then Me.Hide
  Set oTask = Nothing
End Sub

Private Sub cmdClose_Click()
  Unload Me
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
    If CDate(Me.txtTargetFinish.Text) <= ActiveProject.Tasks.UniqueID(Me.lngUID).Start Then
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

