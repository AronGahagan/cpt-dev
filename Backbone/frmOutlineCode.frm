VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutlineCode 
   Caption         =   "Create Outline Code"
   ClientHeight    =   1680
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4212
   OleObjectBlob   =   "frmOutlineCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOutlineCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdGo_Click()
'strings
Dim strOutlineCode As String
'longs
Dim lngOutlineCode As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'only fields with custom names have a left parenthesis
  If InStr(Me.cboOutlineCodes.Value, "(") > 0 Then
    strOutlineCode = Left(Me.cboOutlineCodes.Value, InStr(Me.cboOutlineCodes.Value, " (") - 1)
  Else
    strOutlineCode = Me.cboOutlineCodes.Value
  End If
  lngOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
  If Len(Me.txtNameIt.Value) = 0 Then
    MsgBox "Please provide a name.", vbExclamation + vbOKOnly, "No Name"
    GoTo exit_here
  End If
  Call CreateCode(lngOutlineCode, Me.txtNameIt.Value)
  
exit_here:
  On Error Resume Next
  
  Exit Sub

err_here:
  Call HandleErr("basOutlineCodes", "cmdGo_Click", err)
  Resume exit_here
  
End Sub

Private Sub txtNameIt_Change()
'longs
Dim lngField As Long

  'reset to default
  Me.txtNameIt.BorderColor = -2147483642
  Me.txtNameIt.ForeColor = -2147483640
  Me.lblStatus.Caption = "Ready..."
  
  'if name already exists then flag it
  lngField = 0
  On Error Resume Next
  lngField = FieldNameToFieldConstant(Me.txtNameIt.Text)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If lngField <> 0 Then 'exists
    Me.txtNameIt.BorderColor = 255
    Me.txtNameIt.ForeColor = 255
    Me.lblStatus.Caption = FieldConstantToFieldName(FieldNameToFieldConstant(Me.txtNameIt.Text)) & " is already named '" & Me.txtNameIt.Text & "!"
  End If
  
exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call HandleErr("frmOutlineCode", "txtNameIt_Change", err)
  Resume exit_here
  
End Sub
