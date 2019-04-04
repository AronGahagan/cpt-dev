Attribute VB_Name = "basExportCustomFields"
Sub ExportCustomFields()
Dim lgFieldType As Variant, lgField As Long

  For Each lgFieldType In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
    'For Each strFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Number", "Start", "Text", "Outline Code")
    For Each strFieldType In Array("Outline Code")
      On Error GoTo err_here
      For intField = 1 To 30
        If intField > 1 Then Exit For
        lgField = FieldNameToFieldConstant(strFieldType & intField, lgFieldType)
        strFieldName = CustomFieldGetName(lgField)
        If Len(strFieldName) > 0 Then
          
          'If MsgBox("Delete " & strFieldName & "?", vbQuestion + vbYesNoCancel, "Confirm") = vbYes Then
          '  Application.CustomFieldDelete lgField
          'End If
          
          Debug.Print IIf(lgFieldType = 0, "Text ", "Resource ") & FieldConstantToFieldName(lgField) & ": " & strFieldName
          'If Len(CustomFieldGetFormula(lgField)) > 0 Then Debug.Print "Formula: " & Chr(34) & Application.CustomFieldGetFormula(lgField) & Chr(34)
          On Error GoTo err_here
          If Not IsError(CustomFieldValueListGetItem(lgField, pjValueListValue, 1)) Then
            Debug.Print "Lookup Table: " & strFieldName & "s"
            For intListItem = 1 To 1000
              Debug.Print vbTab & Application.CustomFieldValueListGetItem(lgField, pjValueListValue, intListItem) + " (" + Application.CustomFieldValueListGetItem(lgField, pjValueListDescription, intListItem) + ")"
            Next intListItem
          End If
        End If
next_field:
      Next intField
    Next strFieldType
  Next
  
exit_here:
  Exit Sub
  
err_here:
  If err.Number = 1101 Or err.Number = 1004 Then
    err.Clear
    Resume next_field
  Else
    MsgBox err.Number & ": " & err.Description, vbExclamation + vbOKOnly, "Error"
  End If
End Sub
