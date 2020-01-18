Attribute VB_Name = "cptSaveLocal_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowSaveLocalForm()
'objects
Dim aTypes As Object
Dim rst As ADODB.Recordset
'strings
'longs
Dim lngField As Long
Dim lngFields As Long
Dim lngType As Long
Dim lngECFCount As Long
'integers
'doubles
'booleans
'variants
Dim vType As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set rst = CreateObject("ADODB.Recordset")
  
  rst.Fields.Append "ECF_Constant", adInteger
  rst.Fields.Append "ECF_Name", adVarChar, 120
  rst.Fields.Append "LOCAL_Constant", adInteger
  rst.Fields.Append "LOCAL_Name", adVarChar, 120
  rst.Open
  
  'get enterprise custom fields
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      rst.AddNew Array("ECF_Constant", "ECF_Name"), Array(lngField, FieldConstantToFieldName(lngField))
      lngECFCount = lngECFCount + 1
    End If
  Next lngField

  rst.Sort = "ECF_Name"

  With cptSaveLocal_frm
    'populate map
    .lboMap.Clear
    rst.MoveFirst
    Do While Not rst.EOF
      .lboMap.AddItem
      .lboMap.List(.lboMap.ListCount - 1, 0) = rst(0)
      .lboMap.List(.lboMap.ListCount - 1, 1) = rst(1)
      rst.MoveNext
    Loop
    
    Set aTypes = CreateObject("System.Collections.SortedList")
    aTypes.Add "Cost", 10
    aTypes.Add "Date", 10
    aTypes.Add "Duration", 10
    aTypes.Add "Finish", 10
    aTypes.Add "Flag", 20
    aTypes.Add "Number", 20
    aTypes.Add "Start", 10
    aTypes.Add "Text", 30
    aTypes.Add "Outline Code", 10
    
    'populate field types
    .cboFieldTypes.Clear
    For lngType = 0 To aTypes.Count - 1
      .cboFieldTypes.AddItem
      .cboFieldTypes.List(.cboFieldTypes.ListCount - 1, 0) = aTypes.GetKey(lngType)
      .cboFieldTypes.List(.cboFieldTypes.ListCount - 1, 1) = aTypes.GetByIndex(lngType)
    Next lngType
  
    .cboFieldTypes.Value = "Text"
  
    .lblECFCount.Caption = Format(lngECFCount, "#,##0") & " enterprise custom fields."
    .Show False
    
  End With

exit_here:
  On Error Resume Next
  Set aTypes = Nothing
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptShowSaveLocalForm", Err, Erl)
  Resume exit_here
End Sub

Sub cptSaveLocal()
'objects
Dim Task As Task
'strings
'longs
Dim lngLocal As Long
Dim lngECF As Long
Dim lngMap As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  With cptSaveLocal_frm
    For lngMap = 0 To .lboMap.ListCount - 1
      If .lboMap.List(lngMap, 2) > 0 Then
        lngECF = .lboMap.List(lngMap, 0)
        lngLocal = .lboMap.List(lngMap, 2)
        'rename the local field
        CustomFieldRename CLng(lngLocal), .lboMap.List(lngMap, 1) & " (" & .lboMap.List(lngMap, 3) & ")"
        'populate the fields
        For Each Task In ActiveProject.Tasks
          Task.SetField lngLocal, CStr(Task.GetField(lngECF))
        Next Task
      End If
    Next lngMap
  End With

  'todo: what about using setfield with number and/or date fields, non-flags to flags, etc.

  MsgBox "Enteprise Custom Fields saved locally.", vbInformation + vbOKOnly, "Complete"

exit_here:
  On Error Resume Next
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptSaveLocal", Err, Erl)
  Resume exit_here
End Sub
