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
  UpdateIntegrationSettings
End Sub

Private Sub cboCAM_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboEVT_MS_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboEVP_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboEVT_Change()
  Dim oDict As Scripting.Dictionary
  Dim strValue As String
  Dim oTask As MSProject.Task
  Dim lngItem As Long
  'If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
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
  Set oDict = Nothing
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
  UpdateIntegrationSettings
End Sub

Private Sub cboWBS_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboWP_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
End Sub

Private Sub cboWPM_Change()
  If Not Me.Visible Then Exit Sub
  UpdateIntegrationSettings
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
  Dim lngField As Long
  Dim strField As String
  Dim strControl As String
  If Not Me.Visible Then Exit Sub
  strControl = Me.ActiveControl.Name
  lngField = Me.Controls(strControl).Value
  Me.Controls(strControl).BorderColor = -2147483642
  strControl = Replace(strControl, "cbo", "")
  strField = CustomFieldGetName(lngField)
  If Len(strField) = 0 Then strField = FieldConstantToFieldName(lngField)
  cptSaveSetting "Integration", strControl, lngField & "|" & strField
  If strControl = "EVT" Then
    cptSaveSetting "Metrics", "cboLOEField", lngField
  End If
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
  If CloseMode = 1 Then Me.blnValidIntegrationMap = False
  If Cancel = 1 Then Me.blnValidIntegrationMap = False
End Sub
