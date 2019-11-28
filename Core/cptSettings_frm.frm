VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSettings_frm 
   Caption         =   "ClearPlan Toolbar Settings"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "cptSettings_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSettings_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdSave_Click()
'objects
Dim rst As ADODB.Recordset
'strings
Dim strSettingsFile As String
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set rst = CreateObject("ADODB.Recordset")
  strSettingsFile = cptDir & "\settings\cpt-settings.adtg"
  rst.Open strSettingsFile, , adOpenKeyset
  rst.Find "OPTION='Updates'"
  rst(1) = Not Me.chkDisableUpgrades
  rst.Save strSettingsFile
  rst.Close
  
  Unload Me
  
exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSettings_frm", "cmdSave_Click", Err, Erl)
  Resume exit_here
End Sub
