VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSaveLocal_frm 
   Caption         =   "Save Local"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   OleObjectBlob   =   "cptSaveLocal_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSaveLocal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboFieldTypes_Change()
'objects
'strings
'longs
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
    Me.lboLocalFields.AddItem Me.cboFieldTypes.Column(0) & lngField
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
Dim lngDown As Long
Dim lngCodeNumber As Long
Dim lngField As Long
Dim lngMap As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboMap) And Not IsNull(Me.lboLocalFields) Then
    lngField = FieldNameToFieldConstant(Me.lboLocalFields.List(Me.lboLocalFields.ListIndex))
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
    Me.lboMap.List(Me.lboMap.ListIndex, 3) = Me.lboLocalFields.List(Me.lboLocalFields.ListIndex)
    Me.lboMap.List(Me.lboMap.ListIndex, 2) = FieldNameToFieldConstant(Me.lboLocalFields.List(Me.lboLocalFields.ListIndex))
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
