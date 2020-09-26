VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSaveLocal_frm 
   Caption         =   "Save ECF to LCF"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "cptSaveLocal_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSaveLocal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboFieldTypes_Change()
'objects
'strings
Dim strFieldName As String
'longs
Dim lngFieldID As Long
Dim lngFields As Long
Dim lngField As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboLocalFields.Clear
  lngFields = Me.cboFieldTypes.Column(1)
  For lngField = 1 To lngFields
    strFieldName = Me.cboFieldTypes.Column(0) & lngField
    lngFieldID = FieldNameToFieldConstant(strFieldName)
    If Len(CustomFieldGetName(FieldNameToFieldConstant(Me.cboFieldTypes.Column(0) & lngField))) > 0 Then
      Me.lboLocalFields.AddItem
      Me.lboLocalFields.List(Me.lboLocalFields.ListCount - 1, 0) = lngFieldID
      Me.lboLocalFields.List(Me.lboLocalFields.ListCount - 1, 1) = strFieldName & " (" & CustomFieldGetName(lngFieldID) & ")" 'Me.lboLocalFields.List(Me.lboLocalFields.ListCount - 1, 0) = CustomFieldGetName(FieldNameToFieldConstant(Me.cboFieldTypes.Column(0) & lngField))
    Else
      Me.lboLocalFields.AddItem
      Me.lboLocalFields.List(Me.lboLocalFields.ListCount - 1, 0) = lngFieldID
      Me.lboLocalFields.List(Me.lboLocalFields.ListCount - 1, 1) = strFieldName
    End If
  Next

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "cboFieldTypes_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdCustomFields_Click()
'long
Dim lngSelected As Long
'string
Dim strDescription As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.CustomizeField

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "cmdCustomFields_Click()", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdMap_Click()
'objects
Dim oOutlineCode As OutlineCode
'strings
Dim strECF As String
Dim strLocal As String
'longs
Dim lngItem As Long
Dim lngDown As Long
Dim lngCodeNumber As Long
Dim lngLocal As Long
Dim lngECF As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboMap) And Not IsNull(Me.lboLocalFields) Then
    lngLocal = Me.lboLocalFields.List(Me.lboLocalFields.ListIndex)
    lngECF = Me.lboMap.List(Me.lboMap.ListIndex)
    'if already mapped then prompt with ECF name and ask to remap
    For lngItem = 0 To Me.lboMap.ListCount - 1
      If Me.lboMap.List(lngItem, 3) = lngLocal Then
        If MsgBox(FieldConstantToFieldName(lngLocal) & " is already mapped to " & Me.lboMap.List(lngItem, 1) & " - reassign it?", vbExclamation + vbYesNo, "Already Mapped") = vbYes Then
          Me.lboMap.List(lngItem, 3) = ""
          Me.lboMap.List(lngItem, 4) = ""
        Else
          GoTo exit_here
        End If
      End If
    Next lngItem
    'capture outline code
    'todo: copy codemask; default value; rollup; only leaves; etc.
    If InStr(FieldConstantToFieldName(lngLocal), "Outline") > 0 Then
      MsgBox "If copying down an Outline Code, please use the 'Import Field' function of the Custom Fields dialog before clicking Save Local.", vbInformation + vbOKOnly, "Nota Bene"
      VBA.SendKeys "%r", True
      VBA.SendKeys "f", True
      VBA.SendKeys "%y", True
      VBA.SendKeys "o", True
      VBA.SendKeys "{TAB}"
      'repeat the down key based on which code
      lngCodeNumber = CLng(Replace(FieldConstantToFieldName(lngLocal), "Outline Code", ""))
      If lngCodeNumber > 1 Then
        For lngDown = 1 To lngCodeNumber - 1
          VBA.SendKeys "{DOWN}", True
        Next lngDown
      End If
      VBA.SendKeys "%i", True
      VBA.SendKeys "%f", True
      VBA.SendKeys "%{DOWN}", True
      VBA.SendKeys Left(FieldConstantToFieldName(Me.lboMap.List(Me.lboMap.ListIndex, 0)), 1), True
    End If
    'capture rename
    If Len(CustomFieldGetName(Me.lboLocalFields)) > 0 Then
      If MsgBox("Rename " & FieldConstantToFieldName(Me.lboLocalFields) & " to " & FieldConstantToFieldName(Me.lboMap) & "?", vbQuestion + vbYesNo, "Please confirm") = vbYes Then
        'rename it
        CustomFieldRename CLng(Me.lboLocalFields), Me.lboMap.List(Me.lboMap.ListIndex, 1) & " (" & FieldConstantToFieldName(Me.lboLocalFields) & ")"
        'rename in lboLocalFields
        Me.lboLocalFields.List(Me.lboLocalFields.ListIndex, 1) = FieldConstantToFieldName(Me.lboLocalFields) & " (" & CustomFieldGetName(Me.lboLocalFields) & ")"
      Else
        GoTo exit_here
      End If
    Else
      CustomFieldRename CLng(Me.lboLocalFields), Me.lboMap.List(Me.lboMap.ListIndex, 1) & " (" & FieldConstantToFieldName(Me.lboLocalFields) & ")"
      Me.lboLocalFields.List(Me.lboLocalFields.ListIndex, 1) = FieldConstantToFieldName(Me.lboLocalFields) & " (" & CustomFieldGetName(Me.lboLocalFields) & ")"
    End If
    'get formula
    If Len(CustomFieldGetFormula(lngECF)) > 0 Then
      CustomFieldSetFormula lngLocal, CustomFieldGetFormula(lngECF)
    End If
    'get indicators
    'todo: warn user these are not exposed/available
    'get pick list
    strECF = Me.lboMap.List(Me.lboMap.ListIndex, 1)
    On Error Resume Next
    Set oOutlineCode = GlobalOutlineCodes(strECF)
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Not oOutlineCode Is Nothing Then
      If GlobalOutlineCodes(strECF).LookupTable.Count > 0 Then
        'make it a picklist
        CustomFieldPropertiesEx lngLocal, pjFieldAttributeValueList
        For lngItem = 1 To GlobalOutlineCodes(strECF).LookupTable.Count
          CustomFieldValueListAdd lngLocal, GlobalOutlineCodes(strECF).LookupTable(lngItem).Name, GlobalOutlineCodes(strECF).LookupTable(lngItem).Description
        Next lngItem
      End If
    End If
    Me.lboMap.List(Me.lboMap.ListIndex, 3) = Me.lboLocalFields
    Me.lboMap.List(Me.lboMap.ListIndex, 4) = CustomFieldGetName(Me.lboLocalFields)
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "cmdMap_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdSaveLocal_Click()
  Call cptSaveLocal
End Sub

Private Sub cmdUnmap_Click()
  'objects
  Dim rstSavedMap As ADODB.Recordset
  'strings
  Dim strGUID As String
  Dim strSavedMap As String
  'longs
  Dim lngECF As Long
  Dim lngItem As Long
  Dim lngLCF As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Me.lboMap.ListIndex < 0 Then GoTo exit_here
  If IsNull(Me.lboMap.List(Me.lboMap.ListIndex, 3)) Then GoTo exit_here

  If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Please Confirm") = vbNo Then GoTo exit_here
  
  'get the ECF and LCF
  lngECF = Me.lboMap.List(Me.lboMap.ListIndex, 0)
  lngLCF = Me.lboMap.List(Me.lboMap.ListIndex, 3)
  
  'delete it from LCF
  CustomFieldDelete lngLCF
  
  'delete it from saved map
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    Set rstSavedMap = CreateObject("ADODB.Recordset")
    rstSavedMap.Open strSavedMap
    rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF & " AND LCF=" & lngLCF
    If Not rstSavedMap.EOF Then
      rstSavedMap.Delete adAffectCurrent
      rstSavedMap.Save
    End If
  End If
  
  'remove from lboMap
  Me.lboMap.List(Me.lboMap.ListIndex, 3) = ""
  Me.lboMap.List(Me.lboMap.ListIndex, 4) = ""
  
  'rename in lboLocal
  For lngItem = 0 To Me.lboLocalFields.ListCount - 1
    If Me.lboLocalFields.List(lngItem, 0) = lngLCF Then
      Me.lboLocalFields.List(lngItem, 1) = FieldConstantToFieldName(lngLCF)
    End If
  Next lngItem

exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "cmdUnmap", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblShowFormula_Click()
  If Me.lboMap.ListIndex >= 0 Then
    MsgBox CustomFieldGetFormula(Me.lboMap.List(Me.lboMap.ListIndex, 0)), vbInformation + vbOKOnly, "Formula:"
  End If
End Sub

Private Sub lblURL_Click()
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "lblURL_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboMap_Click()
  'objects
  Dim oLookupTable  As LookupTable
  'strings
  Dim strSwitch As String
  Dim strECF As String
  'longs
  Dim lngItem As Long
  Dim lngMax As Long
  Dim lngECF As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.lblShowFormula.Visible = False
  Me.lblStatus.Caption = "Analyzing..."

  lngECF = Me.lboMap.List(Me.lboMap.ListIndex, 0)
  strECF = Me.lboMap.List(Me.lboMap.ListIndex, 1)
  
  Select Case Me.lboMap.List(Me.lboMap.ListIndex, 2)
    Case "Cost"
      Me.lblStatus.Caption = "This is likely a Cost field."
      strSwitch = "Cost"
    Case "Date"
      Me.lblStatus.Caption = "This is likely a Date field."
      strSwitch = "Date"
    Case "Duration"
      Me.lblStatus.Caption = "This is likely a Duration field."
      strSwitch = "Duration"
    Case "Flag"
      Me.lblStatus.Caption = "This is likely a Flag field."
      strSwitch = "Flag"
    Case "MaybeFlag"
      Me.lblStatus.Caption = "This is likely a Flag field."
      strSwitch = "Flag"
    Case "Number"
      Me.lblStatus.Caption = "This is likely a Number field."
      strSwitch = "Number"
    Case "Outline Code"
      Me.lblStatus.Caption = "This field requires an Outline Code."
      strSwitch = "Outline Code"
    Case "MaybeText"
      Me.lblStatus.Caption = "This is likely a Text field."
      strSwitch = "Text"
    Case "Text"
      Me.lblStatus.Caption = "This is likely a Text field."
      strSwitch = "Text"
    Case Else
      Me.lblStatus.Caption = "Undetermined: confirm manually."
  End Select
  
  If Me.chkAutoSwitch And Me.cboFieldTypes.Value <> strSwitch Then
    Me.cboFieldTypes.Value = strSwitch
  End If
  
exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "lboMap_Click", Err, Erl)
  Resume exit_here
End Sub
