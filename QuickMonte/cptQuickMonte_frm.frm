VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptQuickMonte_frm 
   Caption         =   "QuickMonte"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11070
   OleObjectBlob   =   "cptQuickMonte_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptQuickMonte_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdQuickPERT_Click()
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Call cptQuickPERT(ActiveSelection.Tasks(1).UniqueID)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMonte_frm", "cptQuickPERT_Click", Err, Erl)
  Resume exit_here
  
End Sub
