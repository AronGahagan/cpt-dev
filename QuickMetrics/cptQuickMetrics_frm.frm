VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptQuickMetrics_frm 
   Caption         =   "QuickMetrics - Schedule"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   OleObjectBlob   =   "cptQuickMetrics_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptQuickMetrics_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub ComboBox1_Change()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not FilterApply(Me.ComboBox1) Then
    MsgBox "Filter NAME not found."
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMetrics", "ComboBox1_Change", err)
  Resume exit_here
End Sub
