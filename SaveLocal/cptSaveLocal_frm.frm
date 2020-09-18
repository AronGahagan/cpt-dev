VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSaveLocal_frm 
   Caption         =   "Save ECF to LCF"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205.001
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

Private Sub cmdMap_Click()
'objects
'strings
'longs
Dim lngItem As Long
Dim lngDown As Long
Dim lngCodeNumber As Long
Dim lngField As Long
Dim lngMap As Long
'integers
'doubles
'booleans
'variants
'dates

  'todo: use lngECF and lngLocal; strECR and strLocal

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboMap) And Not IsNull(Me.lboLocalFields) Then
    lngField = Me.lboLocalFields.List(Me.lboLocalFields.ListIndex)
    'if already mapped then prompt with ECF name and ask to remap
    For lngMap = 0 To Me.lboMap.ListCount - 1
      If Me.lboMap.List(lngMap, 2) = lngField Then
        If MsgBox(FieldConstantToFieldName(lngField) & " is already mapped to " & Me.lboMap.List(lngMap, 1) & " - reassign it?", vbExclamation + vbYesNo, "Already Mapped") = vbYes Then
          Me.lboMap.List(lngMap, 2) = ""
          Me.lboMap.List(lngMap, 3) = ""
        Else
          GoTo exit_here
        End If
      End If
    Next lngMap
    'capture outline code
    If InStr(FieldConstantToFieldName(lngField), "Outline") > 0 Then
      MsgBox "If copying down an Outline Code, please use the 'Import Field' function of the Custom Fields dialog before clicking Save Local.", vbInformation + vbOKOnly, "Nota Bene"
      VBA.SendKeys "%r", True
      VBA.SendKeys "f", True
      VBA.SendKeys "%y", True
      VBA.SendKeys "o", True
      VBA.SendKeys "{TAB}"
      'repeat the down key based on which code
      lngCodeNumber = CLng(Replace(FieldConstantToFieldName(lngField), "Outline Code", ""))
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
    'get pick list
    Dim strECF As String
    'strLocal
    strECF = Me.lboMap.List(Me.lboMap.ListIndex, 1)
    'lngECF
    If GlobalOutlineCodes(strECF).LookupTable.Count > 0 Then
      'make it a picklist
      CustomFieldPropertiesEx lngField, pjFieldAttributeValueList
      For lngItem = 1 To GlobalOutlineCodes(strECF).LookupTable.Count
        Dim lngLocal
        CustomFieldValueListAdd lngField, GlobalOutlineCodes(strECF).LookupTable(lngItem).Name
      Next lngItem
    End If
    Me.lboMap.List(Me.lboMap.ListIndex, 2) = Me.lboLocalFields
    Me.lboMap.List(Me.lboMap.ListIndex, 3) = CustomFieldGetName(Me.lboLocalFields)
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptSaveLocal_frm", "cmdMap_Click", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub cmdSaveLocal_Click()
  Call cptSaveLocal
End Sub

Private Sub lblShowFormula_Click()
  If Me.lboMap.ListIndex >= 0 Then
    MsgBox CustomFieldGetFormula(Me.lboMap.List(Me.lboMap.ListIndex, 0)), vbInformation + vbOKOnly, "Formula:"
  End If
End Sub

Private Sub lboMap_Click()
  'objects
  Dim oLookupTable As LookupTable
  'strings
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

  strECF = Me.lboMap.List(Me.lboMap.ListIndex, 1)
  lngECF = Me.lboMap.List(Me.lboMap.ListIndex, 0)
  
  On Error Resume Next
  Set oLookupTable = GlobalOutlineCodes(strECF).LookupTable
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oLookupTable Is Nothing Then
    'does it need an Outline Code?
    If oLookupTable.Count = 0 Then GoTo skip_outline_code_check
    For lngItem = 1 To oLookupTable.Count
      If oLookupTable(lngItem).Level > lngMax Then lngMax = oLookupTable(lngItem).Level
    Next lngItem
    If lngMax = 1 Then 'No
      Me.lblStatus.Caption = "This field requires a Lookup."
      'todo: text or numeric?
    Else 'Yes
      Me.lblStatus.Caption = "This field requires an Outline Code."
      Me.cboFieldTypes.Value = "Outline Code"
      GoTo exit_here
    End If
  End If
  
skip_outline_code_check:
  
  'does it have a formulae?
  If Len(CustomFieldGetFormula(lngECF)) > 0 Then
    Me.lblStatus.Caption = "This field has a formula."
    Me.lblShowFormula.Visible = True
    'todo: text or numeric or duration or date or flag? analyze current field data to determine
  End If
  
  'if all else fails, analyze the data
  
  
exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_frm", "lboMap_Click", Err, Erl)
  Resume exit_here
End Sub
