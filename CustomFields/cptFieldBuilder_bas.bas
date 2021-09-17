Attribute VB_Name = "cptFieldBuilder_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowFieldBuilder_frm()
'objects
Dim aFields As Object
'strings
Dim strEnterprise As String
Dim strCustomFieldName As String
'longs
Dim lngField As Long
Dim lngType As Long
Dim lngCount As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'set up array of field types and counts
  Set aFields = CreateObject("System.Collections.ArrayList")
  'aFields.Add Array("Cost", 10)
  'aFields.Add Array("Date", 10)
  'aFields.Add Array("Duration", 10)
  'aFields.Add Array("Finish", 10)
  aFields.Add Array("Flag", 20)
  aFields.Add Array("Number", 20)
  'aFields.Add Array("Start", 10)
  aFields.Add Array("Text", 30)
  'aFields.Add Array("Outline Code", 10)

  lngType = pjTask
  With cptFieldBuilder_frm
  
    .optResource.Enabled = False
  
    'add field types
    With .cboFieldType
      '.AddItem "Cost"
      '.AddItem "Date"
      '.AddItem "Duration"
      '.AddItem "Finish"
      .AddItem "Flag"
      .AddItem "Number"
      '.AddItem "Start"
      .AddItem "Text"
      '.AddItem "Outline Code"
      .Value = "Text" 'triggers cboFieldType_Change
    End With
    
    'add supported tools
    With .cboTool
      .AddItem "COBRA"
      .AddItem "Empower"
      .AddItem "IPMDAR"
      .AddItem "MPM"
      .AddItem "Guru"
    End With '--> change triggers reload of cboLookups
    
    With .lboFields
      .Clear
      For lngItem = 0 To aFields.Count - 1
        For lngCount = 1 To aFields(lngItem)(1)
          .AddItem
          lngField = FieldNameToFieldConstant(aFields(lngItem)(0) & lngCount, lngType)
          .List(.ListCount - 1, 0) = lngField
          strCustomFieldName = CustomFieldGetName(lngField)
          If Len(strCustomFieldName) > 0 Then
            .List(.ListCount - 1, 1) = strCustomFieldName & " (" & FieldConstantToFieldName(lngField) & ")"
          Else
            .List(.ListCount - 1, 1) = FieldConstantToFieldName(lngField)
          End If
        Next lngCount
      Next lngItem
      
    End With
  
    .Show
    
  End With
exit_here:
  On Error Resume Next
  Set aFields = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFieldBuilder_bas", "cptShowFrmFieldBuilder", Err, Erl)
  Resume exit_here
  cptFieldBuilder_frm.Show False
End Sub

Sub cptBuildField(lngField As Long, strAction As String)
'objects
'strings
Dim strFieldName As String
Dim strField As String
Dim strTool As String
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'todo: trap an accidental overwrite of formulae, lookups, etc.
  strTool = Left(strAction, InStr(strAction, "|") - 1)
  strField = Mid(strAction, InStr(strAction, "|") + 1)
  
  If strTool = "COBRA" Then
    
    Select Case strField
      Case "EVT"
        
        If CustomFieldGetName(lngField) = FieldConstantToFieldName(lngField) Then
          CustomFieldRename lngField, "COBRA EVT"
        Else 'name exists
          strFieldName = CustomFieldGetName(lngField)
          If MsgBox("Keep custom field name '" & strFieldName & "'?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
            strFieldName = "COBRA EVT"
          End If
        End If
        
        CustomFieldDelete lngField
        CustomFieldRename lngField, strFieldName
        CustomFieldPropertiesEx FieldID:=lngField, Attribute:=pjFieldAttributeValueList, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
        CustomFieldValueListAdd lngField, "A", "Level of Effort"
        CustomFieldValueListAdd lngField, "B", "Milestones"
        CustomFieldValueListAdd lngField, "C", "% Complete"
        CustomFieldValueListAdd lngField, "D", "Units Complete"
        CustomFieldValueListAdd lngField, "E", "50-50"
        CustomFieldValueListAdd lngField, "F", "0-100"
        CustomFieldValueListAdd lngField, "G", "100-0"
        CustomFieldValueListAdd lngField, "H", "User Defined"
        CustomFieldValueListAdd lngField, "J", "Apportioned"
        CustomFieldValueListAdd lngField, "K", "Planning Package"
        CustomFieldValueListAdd lngField, "L", "Assignment % Complete"
        CustomFieldValueListAdd lngField, "M", "Calculated Apportionment"
        CustomFieldValueListAdd lngField, "N", "Steps"
        CustomFieldValueListAdd lngField, "O", "Earned As Spent"
        CustomFieldValueListAdd lngField, "P", "% Complete Manual Entry"
                
      Case Else
      
    End Select 'strField
    
  ElseIf strTool = "Empower" Then
  
    Select Case strField
      Case "Task Type"
        
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          If MsgBox("Keep custom field name '" & strFieldName & "'?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
            strFieldName = "Empower Task Type"
          End If
        Else
          strFieldName = "Empower Task Type"
        End If
      
        'add formulae
        CustomFieldDelete lngField
        CustomFieldRename lngField, strFieldName
        CustomFieldPropertiesEx lngField, pjFieldAttributeFormula, pjCalcNone
        Application.CustomFieldSetFormula lngField, "IIf([Summary],""SUM"",Iif([Milestone],""MS"",""AC""))"
      
      Case "Plan Level Code"
        
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          If MsgBox("Keep custom field name '" & strFieldName & "'?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
            strFieldName = "Empower Plan Level Code"
          End If
        Else
          strFieldName = "Empower Plan Level Code"
        End If
        
        CustomFieldDelete lngField
        CustomFieldRename lngField, strFieldName
        CustomFieldPropertiesEx FieldID:=lngField, Attribute:=pjFieldAttributeValueList, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
        CustomFieldValueListAdd lngField, "ACT", "Activity"
        CustomFieldValueListAdd lngField, "CA", "Control Account"
        CustomFieldValueListAdd lngField, "PP", "Planning Package"
        CustomFieldValueListAdd lngField, "SLP", "Summary Level Planning Package"
        CustomFieldValueListAdd lngField, "WP", "Work Package"
        CustomFieldValueListAdd lngField, "ZZZ*", "Mutually Defined"
      
      Case Else
      
    End Select
  
  ElseIf strTool = "MPM" Then
    
    Select Case strField
      Case "Curve Code"
        
      Case "EVM"
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          If MsgBox("Keep custom field name '" & strFieldName & "'?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
            strFieldName = "MPM EVM"
          End If
        Else
          strFieldName = "MPM EVM"
        End If
        
        CustomFieldDelete lngField
        CustomFieldRename lngField, strFieldName
        CustomFieldPropertiesEx FieldID:=lngField, Attribute:=pjFieldAttributeValueList, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
        CustomFieldValueListAdd lngField, "0", "No EVM required"
        CustomFieldValueListAdd lngField, "1", "0/100"
        CustomFieldValueListAdd lngField, "2", "25/75"
        CustomFieldValueListAdd lngField, "3", "40/60"
        CustomFieldValueListAdd lngField, "4", "50/50"
        CustomFieldValueListAdd lngField, "5", "% Complete"
        CustomFieldValueListAdd lngField, "6", "LOE"
        CustomFieldValueListAdd lngField, "7", "Earned Standards"
        CustomFieldValueListAdd lngField, "8", "Milestone Weights"
        CustomFieldValueListAdd lngField, "9", "BCWP Entry"
        CustomFieldValueListAdd lngField, "A", "Apportioned"
        CustomFieldValueListAdd lngField, "P", "Milestone Weights with % Complete"
        CustomFieldValueListAdd lngField, "K", "Key Event"
        
    End Select
    
  ElseIf strTool = "Guru" Then
    Select Case strField
      Case "FTE"
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          If MsgBox("Keep custom field name '" & strFieldName & "'?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbNo Then
            strFieldName = "FTE"
          End If
        Else
          strFieldName = "FTE"
        End If
        CustomFieldDelete lngField
        CustomFieldRename lngField, strFieldName
        CustomFieldPropertiesEx FieldID:=lngField, Attribute:=pjFieldAttributeFormula, SummaryCalc:=pjCalcNone
        Application.CustomFieldSetFormula lngField, "[Remaining Work]/Iif([Remaining Duration]>0,[Remaining Duration],1)"
        
      Case "Task Status"
        
        
      Case "Task Status Indicator"
        'todo: cannot access indicators
        
    End Select
    
  End If 'strTool

  cptFieldBuilder_frm.lboFields.List(cptFieldBuilder_frm.lboFields.ListIndex, 1) = CustomFieldGetName(lngField) & " (" & FieldConstantToFieldName(lngField) & ")"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFieldBuilder_bas", "cptBuildField", Err, Erl)
  Resume exit_here
End Sub

