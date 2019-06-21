VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutlineCode 
   Caption         =   "Create Outline Code"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   OleObjectBlob   =   "frmOutlineCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOutlineCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cboOutlineCodes_Enter()
  Me.cboOutlineCodes.MatchRequired = True
End Sub

Private Sub cmdCancel_Click()
  Me.hide
End Sub

Private Sub cmdGo_Click()
Dim strOutlineCode As String, lgOutlineCode As Long

  If InStr(Me.cboOutlineCodes.Value, "(") > 0 Then
    strOutlineCode = Left(Me.cboOutlineCodes.Value, InStr(Me.cboOutlineCodes.Value, " (") - 1)
  Else
    strOutlineCode = Me.cboOutlineCodes.Value
  End If
  lgOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
  Call CreateCode(lgOutlineCode, Me.txtNameIt)
End Sub
