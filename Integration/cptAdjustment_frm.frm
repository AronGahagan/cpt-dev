VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAdjustment_frm 
   Caption         =   "ETC Adjustment"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7725
   OleObjectBlob   =   "cptAdjustment_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAdjustment_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'<cpt_version>v0.0.3</cpt_version>
Option Explicit

Private Sub cboResources_Change()
  
  If Me.Visible Then Call cptRefreshAdjustment
  
End Sub

Private Sub cmdApply_Click()

  'require an amount
  Me.txtAmount.BorderColor = -2147483642
  If IsNull(Me.txtAmount) Or Len(Me.txtAmount) = 0 Then
    Me.txtAmount.BorderColor = 192
    Exit Sub
  End If
  
  If Not Me.tglScope Then 'hours
    Call cptApplyAdjustment
  Else 'dollars
    Call cptTargetToCost
  End If
  Call cptRefreshAdjustment
    
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdUndo_Click()
  If Application.GetUndoListCount > 1 Then
    If Application.GetUndoListItem(1) = "Target to Cost" Or Application.GetUndoListItem(1) = "cptAdjustment" Then
      Application.Undo
      cptRefreshAdjustment
    End If
  End If
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_frm", "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub optDelta_Click()
  Me.txtAmount.ControlTipText = "Add/Reduce by set number of " & IIf(Me.tglScope, "hours", "dollars")
  'stick to apportioning by remaining work
  cptRefreshAdjustment
End Sub

Private Sub optPercent_Click()
  Me.txtAmount.ControlTipText = "Please use decimal format (50% = '.5')"
  'stick to apportioning by remaining work
  cptRefreshAdjustment
End Sub

Private Sub optTarget_Click()
  Me.txtAmount.ControlTipText = "Apportion to hit Target"
  'stick to apportioning by remaining work
  If Me.Visible Then cptRefreshAdjustment
End Sub

Private Sub tglScope_Click()
  If Me.tglScope Then
    Me.tglScope.Caption = "DOLLARS"
    Me.tglScope.ControlTipText = "Adjust Remaining Cost"
    Me.tglScope.BackColor = &H8000&
    Me.cboResources.Value = 0
    Me.cboResources.Enabled = False
    Me.chkIgnoreTaskType.Value = True
    Me.chkIgnoreTaskType.Enabled = False
  Else
    Me.tglScope.Caption = "HOURS"
    Me.tglScope.ControlTipText = "Adjust Remaining Work"
    Me.tglScope.BackColor = &H8000000F
    Me.cboResources.Enabled = True
    Me.chkIgnoreTaskType.Enabled = True
  End If
  If Me.Visible Then cptRefreshAdjustment
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
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
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
  cptRefreshAdjustment

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_frm", "txtAmount_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Application.GetUndoListCount > 0 Then
    Me.cmdUndo.Enabled = Application.GetUndoListItem(1) = "cptAdjustment" Or Application.GetUndoListItem(1) = "Target to Cost"
  End If
End Sub
