VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAgeDates_frm 
   Caption         =   "Age Dates"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "cptAgeDates_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAgeDates_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboWeeks_Change()
  'objects
  'strings
  Dim strControlName As String
  'longs
  Dim lngControl As Long
  Dim lngWeeks As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not Me.Visible Or Me.cboWeeks.Value = "" Then GoTo exit_here
  lngWeeks = CLng(Replace(Replace(Me.cboWeeks.Value, "weeks", ""), "week", ""))
  For lngControl = 1 To 10
    strControlName = Me.Controls("cboWeek" & lngControl).Name
    If CLng(Replace(strControlName, "cboWeek", "")) <= lngWeeks Then
      Me.Controls("cboWeek" & lngControl).Enabled = True
      Me.Controls("cboWeek" & lngControl).Locked = False
    Else
      Me.Controls("cboWeek" & lngControl).Value = Null
      Me.Controls("cboWeek" & lngControl).Enabled = False
      Me.Controls("cboWeek" & lngControl).Locked = True
    End If
  Next lngControl

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAgeDates_frm", "cboWeeks_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdRun_Click()
  Call cptAgeDates
End Sub

Private Sub lblStatusDate_Click()
  Application.ChangeStatusDate
End Sub
