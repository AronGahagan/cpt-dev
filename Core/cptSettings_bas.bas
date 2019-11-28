Attribute VB_Name = "cptSettings_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowSettingsFrm()
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
  
  strSettingsFile = cptDir & "\settings\cpt-settings.adtg"
  Set rst = CreateObject("ADODB.Recordset")
  If Dir(strSettingsFile) = vbNullString Then 'create one with defaults
    With rst
      .Fields.Append "OPTION", adVarChar, 120
      .Fields.Append "VALUE", adBoolean
      .Open
      .AddNew Array(0, 1), Array("Updates", True)
      .Save strSettingsFile, adPersistADTG
    End With
  Else
    With rst
      .Open strSettingsFile, , adOpenKeyset
      .MoveFirst
      .Find "OPTION='Updates'"
      cptSettings_frm.chkDisableUpgrades = Not rst(1)
    End With
  End If

  cptSettings_frm.Show False

exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSettings_bas", "cptShowSettingsFrm", Err, Erl)
  Resume exit_here
End Sub
