VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptMetricsSettings_frm 
   Caption         =   "Metrics Settings"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "cptMetricsSettings_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptMetricsSettings_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.1.1</cpt_version>
Option Explicit

Private Sub cmdSave_Click()
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  Dim blnValid As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'reset borders
  blnValid = True
  Me.cboEVP.BorderColor = -2147483642
  Me.cboLOEField.BorderColor = -2147483642
  Me.txtLOE.BorderColor = -2147483642
  
  'validate inputs
  If IsNull(Me.cboEVP.Value) Then
    Me.cboEVP.BorderColor = 192
    blnValid = False
  End If
  If IsNull(Me.cboLOEField.Value) Then
    Me.cboLOEField.BorderColor = 192
    blnValid = False
  End If
  If IsNull(Me.txtLOE.Value) Or Len(Me.txtLOE.Value) = 0 Then
    Me.txtLOE.BorderColor = 192
    blnValid = False
  End If
  If IsNull(Me.txtMyHeaders.Value) Or Len(Me.txtMyHeaders.Value) = 0 Then
    Me.txtMyHeaders.BorderColor = 192
    blnValid = False
  End If
  If blnValid Then
    cptSaveSetting "Metrics", "cboEVP", Me.cboEVP.Value
    cptSaveSetting "Metrics", "cboLOEField", Me.cboLOEField.Value
    cptSaveSetting "Metrics", "txtLOE", Me.txtLOE.Value
    cptSaveSetting "Metrics", "txtMyHeaders", Me.txtMyHeaders.Value
    Unload Me
  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMetricsSettings_frm", "cmdSave_Click", Err, Erl)
  Resume exit_here
End Sub
