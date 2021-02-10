VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFieldBuilder_frm 
   Caption         =   "Field Builder"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   OleObjectBlob   =   "cptFieldBuilder_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFieldBuilder_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboTool_Change()
  With Me.cboLookup
    .Clear
    Select Case Me.cboTool
      Case "COBRA"
        .AddItem "EVT"
      Case "Empower"
        .AddItem "Plan Level Code"
        .AddItem "Task Type"
      Case "MPM"
        .AddItem "Curve Code"
        .AddItem "EVM"
      Case "IPMDAR"
        .AddItem "TaskTypeEnum"
        .AddItem "TaskSubTypeEnum"
      Case "Guru"
        .AddItem "FTE"
        .AddItem "Task Status"
        .AddItem "Task Status Indicator"
    End Select
  End With
End Sub

Private Sub cmdBuildField_Click()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  Call cptBuildField(Me.lboFields.Value, Me.cboTool & "|" & Me.cboLookup)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFieldBuilder_frm", "cmdBuildField_Click", Err, Erl)
  Resume exit_here
  
End Sub
