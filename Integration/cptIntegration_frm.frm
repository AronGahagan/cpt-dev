VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptIntegration_frm 
   Caption         =   "Integration"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "cptIntegration_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptIntegration_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.2</cpt_version>
Option Explicit
Public blnValidIntegrationMap As Boolean

Private Sub cboCA_Change()
  If Not Me.Visible Then Exit Sub
  'do not sync WBS, OBS, CA with COBRA Export Tool
  UpdateIntegrationSettings
End Sub

Private Sub cboCAM_Change()
  'objects
  Dim oCDP As DocumentProperty
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not Me.Visible Then Exit Sub
  If Me.chkSyncSettings Then
    On Error Resume Next
    Set oCDP = ActiveProject.CustomDocumentProperties("fCAM")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oCDP Is Nothing Then
      lngField = FieldNameToFieldConstant(ActiveProject.CustomDocumentProperties("fCAM"))
      If Me.cboCAM.Value <> lngField Then
        If MsgBox("COBRA Export Tool setting is '" & CustomFieldGetName(lngField) & "' - use this instead?", vbQuestion + vbYesNo, "Synchronize?") = vbYes Then
          Me.cboCAM.Value = lngField
        End If
      End If
    Else
      Set oCDP = ActiveProject.CustomDocumentProperties.Add("fCAM", False, msoPropertyTypeString, FieldConstantToFieldName(Me.cboCAM.Value))
    End If
  End If
  UpdateIntegrationSettings
  
exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "cboCAM_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub cboEVT_MS_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboEVP_Change()
  'objects
  Dim oCDP As DocumentProperty
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not Me.Visible Then Exit Sub
  If Me.chkSyncSettings Then
    
    On Error Resume Next
    Set oCDP = ActiveProject.CustomDocumentProperties("fPCNT")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oCDP Is Nothing Then
      lngField = FieldNameToFieldConstant(ActiveProject.CustomDocumentProperties("fPCNT"))
      If Me.cboEVP.Value <> lngField Then
        If MsgBox("COBRA Export Tool setting is '" & CustomFieldGetName(lngField) & "' - use this instead?", vbQuestion + vbYesNo, "Synchronize?") = vbYes Then
          Me.cboEVP.Value = lngField
        End If
      End If
    Else
      Set oCDP = ActiveProject.CustomDocumentProperties.Add("fPCNT", False, msoPropertyTypeString, FieldConstantToFieldName(Me.cboEVP.Value))
    End If
  End If
  UpdateIntegrationSettings

exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "cboEVT_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cboEVT_Change()
  'objects
  Dim oCDP As DocumentProperty
  Dim oDict As Scripting.Dictionary
  Dim oTask As MSProject.Task
  'strings
  Dim strValue As String
  'longs
  Dim lngField As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings 'todo: can this be moved to only the end?
  Me.cboLOE.Value = ""
  Me.cboLOE.Clear
  Set oDict = CreateObject("Scripting.Dictionary")
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    strValue = oTask.GetField(Me.cboEVT.Value)
    If Len(strValue) > 0 Then
      If Not oDict.Exists(strValue) Then oDict.Add strValue, strValue
    End If
next_task:
  Next oTask
  For lngItem = 0 To oDict.Count - 1
    Me.cboLOE.AddItem oDict.Items(lngItem)
  Next lngItem
  If Me.chkSyncSettings Then
    
    On Error Resume Next
    Set oCDP = ActiveProject.CustomDocumentProperties("fEVT")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oCDP Is Nothing Then
      lngField = FieldNameToFieldConstant(ActiveProject.CustomDocumentProperties("fEVT"))
      If Me.cboEVT.Value <> lngField Then
        If MsgBox("COBRA Export Tool setting is '" & CustomFieldGetName(lngField) & "' - use this instead?", vbQuestion + vbYesNo, "Synchronize?") = vbYes Then
          Me.cboEVT.Value = lngField
        End If
      End If
    Else
      Set oCDP = ActiveProject.CustomDocumentProperties.Add("fEVT", False, msoPropertyTypeString, FieldConstantToFieldName(Me.cboEVT.Value))
    End If
  End If
  UpdateIntegrationSettings
  
exit_here:
  On Error Resume Next
  Set oDict = Nothing
  Set oCDP = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "cboEVT_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cboLOE_Change()
  If Not Me.Visible Then Exit Sub
  If Me.cboLOE.Value = "" Then
    Me.cboLOE.BorderColor = 192
  Else
    Me.cboLOE.BorderColor = -2147483642
    cptSaveSetting "Metrics", "txtLOE", Me.cboLOE.Value
    cptSaveSetting "Integration", "LOE", Me.cboLOE.Value
  End If
End Sub

Private Sub cboOBS_Change()
  If Not Me.Visible Then Exit Sub
  'do not sync WBS, OBS, CA with COBRA Export Tool
  UpdateIntegrationSettings
End Sub

Private Sub cboWBS_Change()
  If Not Me.Visible Then Exit Sub
  'do not sync WBS, OBS, CA with COBRA Export Tool
  UpdateIntegrationSettings
End Sub

Private Sub cboWP_Change()
  'objects
  Dim oCDP As DocumentProperty
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not Me.Visible Then Exit Sub
  If Me.chkSyncSettings Then
    
    On Error Resume Next
    Set oCDP = ActiveProject.CustomDocumentProperties("fWP")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If Not oCDP Is Nothing Then
      lngField = FieldNameToFieldConstant(oCDP.Value)
      If Me.cboWP.Value <> lngField Then
        If MsgBox("COBRA Export Tool setting is '" & CustomFieldGetName(lngField) & "' - use this instead?", vbQuestion + vbYesNo, "Synchronize?") = vbYes Then
          Me.cboWP.Value = lngField
        End If
      End If
    Else
      Set oCDP = ActiveProject.CustomDocumentProperties.Add("fWP", False, msoPropertyTypeString, CustomFieldGetName(Me.cboWP.Value))
    End If
  End If
  UpdateIntegrationSettings
  
exit_here:
  On Error Resume Next
  Set oCDP = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "cboWP_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub cboWPM_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub chkSyncSettings_Click()
  'objects
  Dim oCDP As DocumentProperty
  Dim oComboBox As MSForms.ComboBox
  Dim oDict As Scripting.Dictionary
  'strings
  Dim strFields As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim vControl As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSaveSetting "Integration", "chkSyncSettings", IIf(Me.chkSyncSettings, "1", "0")
  
  strFields = "CAM,WP,EVT,EVP"
  'map integration settings with COBRA Export Tool settings
  Set oDict = CreateObject("Scripting.Dictionary")
  'do not sync WBS
  'do not sync OBS
  'do not sync CA
  oDict.Add "CAM", "fCAM"
  oDict.Add "WP", "fWP"
  oDict.Add "EVT", "fEVT"
  oDict.Add "EVP", "fPCNT"
  
  If Me.chkSyncSettings Then
    For Each vControl In Split(strFields, ",")
      Set oComboBox = Me.Controls("cbo" & vControl)
      If Not oComboBox.Enabled Then GoTo next_control
      On Error Resume Next
      Set oCDP = ActiveProject.CustomDocumentProperties(oDict(vControl))
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If Not oCDP Is Nothing Then
        'does it still exist?
        On Error Resume Next
        Dim lngField As Long
        lngField = 0
        lngField = FieldNameToFieldConstant(oCDP.Value)
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If lngField = 0 Then
          If MsgBox("The field '" & oCDP.Value & "' in the COBRA settings no longer exists in this file. Update with '" & CustomFieldGetName(oComboBox.Value) & "'?", vbCritical + vbYesNo, "Mapping Invalid") = vbYes Then
            oCDP = CustomFieldGetName(oComboBox.Value)
            oComboBox.BorderColor = -2147483642
          Else
            oComboBox.BorderColor = 192
            oCDP.Delete
          End If
        Else
          If IsNull(oComboBox) Then 'import from COBRA Export Tool setting
            oComboBox.Value = FieldNameToFieldConstant(oCDP.Value)
            cptSaveSetting "Integration", CStr(vControl), oComboBox.Value & "|" & CustomFieldGetName(oComboBox.Value)
            oComboBox.BorderColor = -2147483642
          Else 'notify discrepancy
            If FieldNameToFieldConstant(oCDP.Value) <> oComboBox.Value Then
              oComboBox.BorderColor = 192
            Else
              oComboBox.BorderColor = -2147483642
            End If
          End If
        End If
      End If
next_control:
    Next vControl
  End If
  
exit_here:
  Set oCDP = Nothing
  Set oComboBox = Nothing
  Set oDict = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "chkSyncSettings_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCancel_Click()
  Me.blnValidIntegrationMap = False
  Me.Hide
End Sub

Private Sub cmdConfirm_Click()
  Dim blnValid As Boolean
  Dim oControl As MSForms.Control
  
  blnValid = True
  For Each oControl In Me.Controls
    If Left(oControl.Name, 3) = "cmd" Then GoTo next_control
    If Left(oControl.Name, 3) = "chk" Then GoTo next_control
    If oControl.BorderColor = 192 Then
      blnValid = False
      Exit For
    End If
next_control:
  Next oControl
  
  Me.blnValidIntegrationMap = blnValid
  Me.Hide
End Sub

Private Sub UpdateIntegrationSettings()
  'objects
  'strings
  Dim strControl As String
  Dim strField As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vControl As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not Me.Visible Then Exit Sub
  strControl = Me.ActiveControl.Name
  If Left(strControl, 3) <> "cbo" Then Exit Sub
  lngField = Me.Controls(strControl).Value
  Me.Controls(strControl).BorderColor = -2147483642
  strControl = Replace(strControl, "cbo", "")
  strField = CustomFieldGetName(lngField)
  If Len(strField) = 0 Then strField = FieldConstantToFieldName(lngField)
  cptSaveSetting "Integration", strControl, lngField & "|" & strField
  'sync metrics settings
  If strControl = "EVT" Then
    cptSaveSetting "Metrics", "cboLOEField", lngField
  End If
  If strControl = "LOE" Then
    cptSaveSetting "Metrics", "txtLOE", Me.cboLOE.Value
  End If
  If strControl = "EVP" Then
    cptSaveSetting "Metrics", "cboEVP", lngField
  End If
  'validate against COBRA Export Tool
  'todo: does sync COBRA Export Tool conflict with strRequiredFields?
  If Me.chkSyncSettings Then
    For Each vControl In Split("CAM,WP,EVT,EVP", ",")
      strControl = CStr(vControl)
      If strControl = "EVP" Then strControl = "PCNT"
      If CustomFieldGetName(Me.Controls("cbo" & vControl).Value) <> ActiveProject.CustomDocumentProperties("f" & strControl) Then
        If FieldConstantToFieldName(Me.Controls("cbo" & vControl).Value) <> ActiveProject.CustomDocumentProperties("f" & strControl) Then 'catch if not a custom field (e.g., Physical % Complete)
          If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).BorderColor = 192
        Else
          If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).BorderColor = -2147483642
        End If
      Else
        If Me.Controls("cbo" & vControl).Enabled Then Me.Controls("cbo" & vControl).BorderColor = -2147483642
      End If
    Next vControl
  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "UpdateIntegrationSettings", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtRollingWave_AfterUpdate()
  Dim dtRollingWave As Date
  If Len(Me.txtRollingWave) > 0 Then
    If IsDate(Me.txtRollingWave) Then
      dtRollingWave = CDate(Me.txtRollingWave.Value)
      Me.txtRollingWave = FormatDateTime(dtRollingWave, vbShortDate)
      cptSaveSetting "Integration", "RollingWaveDate", dtRollingWave
    Else
      Me.txtRollingWave = ""
    End If
  Else
    cptDeleteSetting "Integration", "RollingWaveDate"
  End If
  Me.lblWeekday.Visible = False
End Sub

Private Sub txtRollingWave_Change()
  Dim dtDate As Date
  
  If Not Me.Visible Then Exit Sub
  If Not Me.ActiveControl.Name = Me.txtRollingWave.Name Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lblWeekday.Visible = False
  If Len(Me.txtRollingWave.Text) > 0 Then
    Me.txtRollingWave.Text = cptRegEx(Me.txtRollingWave.Text, "[0-9\/]{1,}")
    'limit to a date only
    If Not IsDate(Me.txtRollingWave.Text) Then
      Me.txtRollingWave.BorderColor = 192
      Me.Repaint
    Else
      'limit to dates after the start date
      dtDate = CDate(Format(Me.txtRollingWave.Text, "mm/dd/yyyy") & " 5:00 PM")
      If dtDate < ActiveProject.StatusDate Then
        Me.txtRollingWave.BorderColor = 192
        Me.Repaint
      Else
        Me.txtRollingWave.BorderColor = -2147483642
        Me.lblWeekday.Caption = Format(dtDate, "dddd")
        Me.lblWeekday.Visible = True
        Me.Repaint
      End If
    End If
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIntegration_frm", "txtRollingWave_Change", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtRollingWave_Enter()
  If IsDate(Me.txtRollingWave) Then
    Me.lblWeekday.Caption = Format(CDate(Me.txtRollingWave), "dddd")
    Me.lblWeekday.Visible = True
  End If
End Sub

Private Sub txtRollingWave_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblWeekday.Visible = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Me.blnValidIntegrationMap = False
    Me.Hide
    Cancel = True
  End If
End Sub
