Attribute VB_Name = "cptSaveLocal_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

'todo: create view ECF:Local;ECF:Local;ECF:Local;
' -- based on saved map, current selections
'todo: determine ECF field type and auto-change cboTypes upon selection
'todo: handle resource and project custom fields
'todo: import saved map or create new on cmdSave_Click
'todo: save mapping by project GUID
'todo: process is: import with enterprise open, then save as mpp
'todo: make compatible with master/sub projects

Sub cptShowSaveLocalForm()
'objects
Dim aTypes As Object
Dim rst As ADODB.Recordset
'strings
Dim strGUID As String
Dim strECF As String
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

  strGUID = ActiveProject.GetServerProjectGuid

  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "GUID", adGUID
  rst.Fields.Append "pjType", adInteger
  rst.Fields.Append "ECF_Constant", adInteger
  rst.Fields.Append "ECF_Name", adVarChar, 120
  rst.Fields.Append "LOCAL_Constant", adInteger
  rst.Fields.Append "LOCAL_Name", adVarChar, 120
  rst.Open
  
  'get enterprise custom fields
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strECF = FieldConstantToFieldName(lngField)
      If lngField = FieldNameToFieldConstant(strECF, pjTask) Then
        lngType = pjTask
      ElseIf lngField = FieldNameToFieldConstant(strECF, pjResource) Then
        lngType = pjResource
      ElseIf lngField = FieldNameToFieldConstant(strECF, pjProject) Then
        lngType = pjProject
      Else
        lngType = 10
      End If
      rst.AddNew Array("GUID", "pjType", "ECF_Constant", "ECF_Name"), Array(strGUID, lngType, lngField, FieldConstantToFieldName(lngField))
      lngECFCount = lngECFCount + 1
    End If
  Next lngField

  rst.Sort = "ECF_Name"

  With cptSaveLocal_frm
    'populate map
    .lboMap.Clear
    If rst.RecordCount = 0 Then
      rst.Close
      MsgBox "No Enterprise Custom Fields available in this file.", vbExclamation + vbOKOnly, "No ECFs found"
      GoTo exit_here
    End If
    rst.MoveFirst
    Do While Not rst.EOF
      If UCase(rst(0)) = UCase(strGUID) And rst("pjType") = 0 Then
        .lboMap.AddItem
        .lboMap.List(.lboMap.ListCount - 1, 0) = rst(2)
        .lboMap.List(.lboMap.ListCount - 1, 1) = rst(3)
      End If
      rst.MoveNext
    Loop
    rst.Close
    
    Set aTypes = CreateObject("System.Collections.SortedList")
    'record: field type, number of available custom fields
    For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
      aTypes.Add vType, 10
    Next
    aTypes.Add "Flag", 20
    aTypes.Add "Number", 20
    aTypes.Add "Text", 30
    
    'populate field types
    .cboFieldTypes.Clear
    For lngType = 0 To aTypes.Count - 1
      .cboFieldTypes.AddItem
      .cboFieldTypes.List(.cboFieldTypes.ListCount - 1, 0) = aTypes.GetKey(lngType)
      .cboFieldTypes.List(.cboFieldTypes.ListCount - 1, 1) = aTypes.GetByIndex(lngType)
    Next lngType
  
    .cboFieldTypes.Value = "Text"
  
    .lblStatus.Caption = Format(lngECFCount, "#,##0") & " enterprise custom fields."
    .Show False
    
  End With

exit_here:
  On Error Resume Next
  Set vType = Nothing
  Set aTypes = Nothing
  If rst.State Then rst.Close
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

  'todo: only cycle through tasks once, mapped field each time

  With cptSaveLocal_frm
    For lngMap = 0 To .lboMap.ListCount - 1
      If .lboMap.List(lngMap, 2) > 0 Then
        lngECF = .lboMap.List(lngMap, 0)
        lngLocal = .lboMap.List(lngMap, 2)
        'populate the fields
        For Each Task In ActiveProject.Tasks
          On Error Resume Next
          If Len(Task.GetField(lngLocal)) > 0 Then Task.SetField lngLocal, ""
          If Len(Task.GetField(lngECF)) > 0 Then
            Task.SetField lngLocal, CStr(Task.GetField(lngECF))
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            If Task.GetField(lngLocal) <> CStr(Task.GetField(lngECF)) Then
              MsgBox "There was an error copying from ECF " & CustomFieldGetName(lngECF) & " to LCF " & CustomFieldGetName(lngLocal) & "." & vbCrLf & vbCrLf & "Please validate data types.", vbExclamation + vbOKOnly, "Fail"
              GoTo next_mapping
            End If
          End If
        Next Task
      End If
next_mapping:
    Next lngMap
  End With

  MsgBox "Enteprise Custom Fields saved locally.", vbInformation + vbOKOnly, "Complete"

exit_here:
  On Error Resume Next
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptSaveLocal", Err, Erl)
  Resume exit_here
End Sub

Function cptGetECFType()
'objects
Dim rstValues As ADODB.Recordset
Dim rstFields As ADODB.Recordset
Dim oTask As Task
'strings
Dim strRecord As String
Dim strFields As String
Dim strValues As String
Dim strCon As String
Dim strDir As String
Dim strSQL As String
Dim strFile As String
'longs
Dim lngField As Long
Dim lngFields As Long
Dim lngValues As Long
Dim lngItem As Long
Dim lngGOC As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'define dir
  strDir = Environ("USERPROFILE")
  'define fields csv
  strFields = "ecf_fields.csv"
  lngFields = FreeFile
  Open strDir & "\" & strFields For Output As #lngFields
  
  'first capture the ecf fields
  lngItem = 0
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      Print #lngFields, lngItem & "," & lngField & "," & FieldConstantToFieldName(lngField)
      lngItem = lngItem + 1
    End If
  Next lngField
  
  'define values csv
  strValues = "ecf_values.csv"
  lngValues = FreeFile
  Open strDir & "\" & strValues For Output As #lngValues
  
  For Each oTask In ActiveProject.Tasks
    lngItem = 0
    'get enterprise custom fields
    For lngField = 188776000 To 188778000 '2000 should do it for now
      If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strRecord = strRecord & oTask.GetField(lngField) & ","
        lngItem = lngItem + 1
        If lngItem = 33 Then Stop
      End If
    Next lngField
    Print #lngValues, oTask.UniqueID & "," & strRecord
    strRecord = ""
  Next oTask
  Close #lngFields
  Close #lngValues
  
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=No;IMEX=2;FMT=Delimited';"
  Set rstFields = CreateObject("ADODB.Recordset")
  rstFields.Open "SELECT * FROM [" & strFields & "]", strCon, adOpenKeyset
  Set rstValues = CreateObject("ADODB.Recordset")
  rstValues.Open "SELECT * FROM [" & strValues & "]", strCon, adOpenKeyset
  rstFields.MoveFirst
  Do While Not rstFields.EOF
    Debug.Print rstFields(0); rstFields(1); rstFields(2); rstValues.Fields(rstFields.AbsolutePosition - 1).Type
    rstFields.MoveNext
  Loop
  rstFields.Close
  rstValues.Close
  
  '3 = adInteger = Number
  '202 = adVarWChar = Text
  
  Kill strDir & "\" & strFields
  Kill strDir & "\" & strValues
  
'  'todo: use this to at least capture which ECFs are Outline Codes
'  For lngGOC = 1 To Application.GlobalOutlineCodes.Count
'
'  Next lngGOC

exit_here:
  On Error Resume Next
  Set rstValues = Nothing
  If rstFields.State Then rstFields.Close
  Set rstFields = Nothing
  Set oTask = Nothing

  Exit Function
err_here:
  'Call HandleErr("foo", "bar", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Function

Function cptInterrogateECF(lngField As Long)
  'objects
  Dim oTask As Task
  'strings
  Dim strPattern As String
  Dim strVal As String
  'longs
  Dim lngVal As Long
  'integers
  'doubles
  'booleans
  Dim blnVal As Boolean
  'variants
  'dates
  Dim dtVal As Date
  
  'todo: determine if field is task, project, or resource type
  
  On Error Resume Next
     
  'todo: create only a single task and interrogate all ECFs at once on form load
  Set oTask = ActiveProject.Tasks.Add("deleteMe")
    
  On Error Resume Next
  oTask.SetField lngField, "xxx"
  If Err.Description = "This field only supports positive numbers." Then
    cptInterrogateECF = "Cost"
  ElseIf Err.Description = "The date you entered isn't supported for this field." Then
    cptInterrogateECF = "Date"
  ElseIf Err.Description = "The duration you entered isn't supported for this field." Then
    cptInterrogateECF = "Duration"
  ElseIf Err.Description = "Select either Yes or No from the list." Then
    cptInterrogateECF = "Flag"
  ElseIf Err.Description = "This field only supports numbers." Then
    cptInterrogateECF = "Number"
  ElseIf Err.Description = "This is not a valid lookup table value." Then
    'select the first value and check it
    'fails if text picklist has cost
    'fails if text picklist has dates
    'fails if text picklist has numbers
    'fails if text picklist has currency label
    oTask.SetField lngField, Application.CustomFieldValueListGetItem(lngField, pjValueListValue, 1)
    strVal = oTask.GetField(lngField)
    GoTo enhanced_interrogation
  ElseIf Err.Description = "The argument value is not valid." Then
    'figure out formula
    If Len(CustomFieldGetFormula(lngField)) > 0 Then
      strVal = oTask.GetField(lngField)
      GoTo enhanced_interrogation
    End If
  ElseIf Err.Description = "" Then
    cptInterrogateECF = "Text"
  End If
  
  GoTo exit_here
  
enhanced_interrogation:
  
  Err.Clear
  
  'check for cost
  If InStr(strVal, ActiveProject.CurrencySymbol) > 0 Then
    cptInterrogateECF = "Cost"
    GoTo exit_here
  End If
  
  'check for number
  On Error Resume Next
  lngVal = oTask.GetField(lngField)
  If Err.Number = 0 Then
    cptInterrogateECF = "Number"
    GoTo exit_here
  End If
  
  On Error Resume Next
  dtVal = oTask.GetField(lngField)
  If Err.Number = 0 Then
    cptInterrogateECF = "Date"
    GoTo exit_here
  End If
  
  If Len(cptRegEx(strVal, "Yes|No")) > 0 Then
    cptInterrogateECF = "MaybeFlag"
    GoTo exit_here
  End If
  
  Err.Clear
  strVal = oTask.GetField(lngField)
  'could be duration
  If strVal = DurationFormat(DurationValue(strVal), ActiveProject.DefaultDurationUnits) Then
    cptInterrogateECF = "Duration"
    GoTo exit_here
  End If
  
  'could be flag
  
  
  'otherwise, it's (probably) text
  cptInterrogateECF = "Text"

exit_here:
  Err.Clear
  On Error Resume Next
  oTask.Delete
  Set oTask = Nothing
  
  Exit Function
err_here:
  Call cptHandleErr("foo", "bar", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Function


Sub cptGetAllFields()
  'objects
  Dim oWorksheet As Worksheet
  Dim oWorkbook As Workbook
  Dim rst As ADODB.Recordset
  Dim oExcel As Excel.Application
  'strings
  Dim strCustomName As String
  Dim strDir As String
  Dim strFile As String
  Dim strFieldName As String
  'longs
  Dim lngTo As Long
  Dim lngFrom As Long
  Dim lngFile As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "Constant", adBigInt
  rst.Fields.Append "Name", adVarChar, 155
  rst.Fields.Append "CustomName", adVarChar, 155
  rst.Open
  
  '184549399 = lowest = ID
  '184550803 start of <Unavailable>
  '188744879 might be last of built-ins
  '188750001 start of ecfs?
  '218103807 highest and enterprise
  
  'restart at 188800001
  lngFrom = 215000001
  lngTo = 218103807
  
  For lngField = lngFrom To lngTo
    strFieldName = FieldConstantToFieldName(lngField)
    If Len(strFieldName) > 0 And strFieldName <> "<Unavailable>" Then
      strCustomName = CustomFieldGetName(lngField)
      rst.AddNew Array(0, 1, 2), Array(lngField, strFieldName, strCustomName)
    End If
    Debug.Print "Processing " & Format(lngField, "###,###,##0") & " of " & Format(lngTo, "###,###,##0") & " (" & Format(lngField / lngTo, "0%") & ")"
  Next lngField

  If rst.RecordCount > 0 Then
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.[A1].CopyFromRecordset rst
  Else
    MsgBox "No fields found between " & lngFrom & " and " & lngTo & ".", vbInformation + vbOKOnly, "No results."
  End If
exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  rst.Close
  Set rst = Nothing
  Set oExcel = Nothing
  Close #lngFile
  Exit Sub
err_here:
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

