VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAdjustment_frm 
   Caption         =   "ETC Adjustment"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "cptAdjustment_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAdjustment_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboResources_Change()
  Dim strResource As String
  If Not IsNull(Me.cboResources.Value) Then strResource = Me.cboResources.Value
  Call cptRefreshAdjustment(strResource:=CStr(Me.cboResources.Value))
  
End Sub

Private Sub cmdApply_Click()

  'require an amount
  If Len(Me.txtAmount) = 0 Then
    'todo: turn the frame red
    Exit Sub
  End If
  
  Call cptApplyAdjustment(Me.cboResources.Value, "target", CDbl(Me.txtAmount.Value))
    
End Sub

Private Sub cmdUndo_Click()
  Application.Undo
  cptRefreshAdjustment 'todo: pipe in the settings
End Sub

Private Sub optDelta_Click()
  Me.txtAmount.ControlTipText = ""
  cptRefreshAdjustment 'todo: pipe in the settings
End Sub

Private Sub optPercent_Click()
  Me.txtAmount.ControlTipText = "Please use decimal format"
  cptRefreshAdjustment 'todo: pipe in the settings
End Sub

Private Sub optTarget_Click()
  Me.txtAmount.ControlTipText = ""
  If Me.Visible Then cptRefreshAdjustment 'todo: pipe in the settings
End Sub

Private Sub txtAmount_Change()
  'objects
  'strings
  Dim strAmount As String
  'longs
  'integers
  'doubles
  Dim dblAmount As Double
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  strAmount = Me.txtAmount.Text
  If strAmount = "-" Then GoTo exit_here 'be patient
  If strAmount = "." Or Right(strAmount, 1) = "." Then GoTo exit_here 'be patient
  If strAmount = "-." Then GoTo exit_here 'be patient
  If Len(strAmount) > 0 Then
    If Me.optTarget Then
      dblAmount = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
      If dblAmount < 0 Then
        dblAmount = Abs(dblAmount)
        Me.txtAmount = dblAmount
      End If
    ElseIf Me.optPercent Then
      If Right(strAmount, 1) = "%" Then
        strAmount = Replace(strAmount, "%", "")
        dblAmount = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
        Me.txtAmount.Text = dblAmount / 100
      Else
        Me.txtAmount.Text = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
      End If
    Else
      Me.txtAmount.Text = cptRegEx(strAmount, "(-)?([0-9]{1,})?(\.[0-9]{1,})?")
    End If
  End If
  cptRefreshAdjustment 'todo: pipe in the settings

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_frm", "txtAmount_Change", Err, Erl)
  Resume exit_here
End Sub
