Attribute VB_Name = "cptSaveLocal_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

'todo: save rst as rstECF for session only
'todo: create view ECF:Local;ECF:Local;ECF:Local;
' -- based on saved map, current selections
'todo: handle resource custom fields - add toggle option and filter accordinlgy
'todo: process is: import with enterprise open, then save as mpp
'todo: make compatible with master/sub projects
'todo: handle when user changes custom fields manually -- onmouseover
'todo: code up the search filter
'todo: implement a 'suggest' feature
' -- count ECF vs Available LCF; Automap them.

Sub cptShowSaveLocalForm()
'objects
Dim oTask As Task
Dim rstSavedMap As ADODB.Recordset
Dim aTypes As Object
Dim rst As ADODB.Recordset
'strings
Dim strSaved As String
Dim strEntity As String
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
Dim blnExists As Boolean
'variants
Dim vType As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSpeed True
  
  'get project guid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If

  'prepare to capture all ECFs
  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "GUID", adGUID
  rst.Fields.Append "pjType", adInteger
  rst.Fields.Append "ENTITY", adVarChar, 50
  rst.Fields.Append "ECF_Constant", adInteger
  rst.Fields.Append "ECF_Name", adVarChar, 120
  'rst.Fields.Append "LCF_Constant", adInteger
  'rst.Fields.Append "LCF_Name", adVarChar, 120
  rst.Open
  
  'create a dummy task to interrogate the ECFs
  Set oTask = ActiveProject.Tasks.Add("<dummy for cpt-save-local>")
  Application.CalculateProject
  
  'get enterprise custom task fields
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strECF = FieldConstantToFieldName(lngField)
      strEntity = cptInterrogateECF(oTask, lngField)
      rst.AddNew Array("GUID", "pjType", "ENTITY", "ECF_Constant", "ECF_Name"), Array(strGUID, pjTask, strEntity, lngField, FieldConstantToFieldName(lngField))
      lngECFCount = lngECFCount + 1
    End If
  Next lngField

  'get enterprise custom resource fields
  For lngField = 205553664 To 205555664 '2000 should do it for now
    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strECF = FieldConstantToFieldName(lngField)
      strEntity = cptInterrogateECF(oTask, lngField)
      rst.AddNew Array("GUID", "pjType", "ENTITY", "ECF_Constant", "ECF_Name"), Array(strGUID, pjResource, strEntity, lngField, FieldConstantToFieldName(lngField))
      lngECFCount = lngECFCount + 1
    End If
  Next lngField
  
  If Dir(cptDir & "\settings\cpt-ecf.adtg") <> vbNullString Then
    Kill cptDir & "\settings\cpt-ecf.adtg"
  End If
  rst.Sort = "ECF_Name"
  rst.Save cptDir & "\settings\cpt-ecf.adtg"
  
  'check for saved map
  
  strSaved = cptDir & "\settings\cpt-save-local.adtg"
  blnExists = Dir(strSaved) <> vbNullString
  If blnExists Then
    Set rstSavedMap = CreateObject("ADODB.Recordset")
    rstSavedMap.Open strSaved
  End If
  
  'populate the form - defaults to task ECFs, text
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
      If UCase(rst("GUID")) = UCase(strGUID) And rst("pjType") = 0 Then
        .lboMap.AddItem
        .lboMap.List(.lboMap.ListCount - 1, 0) = rst("ECF_Constant")
        .lboMap.List(.lboMap.ListCount - 1, 1) = rst("ECF_Name")
        .lboMap.List(.lboMap.ListCount - 1, 2) = rst("ENTITY")
        If blnExists Then
          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & rst("ECF_Constant") '& " AND ENTITY=" & pjTask
          If Not rstSavedMap.EOF Then
            .lboMap.List(.lboMap.ListCount - 1, 3) = rstSavedMap("LCF")
            If Len(CustomFieldGetName(rstSavedMap("LCF"))) > 0 Then
              .lboMap.List(.lboMap.ListCount - 1, 4) = CustomFieldGetName(rstSavedMap("LCF"))
            Else
              .lboMap.List(.lboMap.ListCount - 1, 4) = FieldConstantToFieldName(rstSavedMap("LCF"))
            End If
          End If
          rstSavedMap.Filter = ""
        End If
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
    
    .optTasks = True
    .cboFieldTypes.Value = "Text"
  
    .lblStatus.Caption = Format(lngECFCount, "#,##0") & " enterprise custom fields."
    oTask.Delete
    cptSpeed False
    .Show False
    
  End With

exit_here:
  On Error Resume Next
  cptSpeed False
  oTask.Delete
  Set oTask = Nothing
  Set rstSavedMap = Nothing
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
Dim rstSavedMap As ADODB.Recordset
Dim oTask As Task
'strings
Dim strGUID As String
Dim strSavedMap As String
'longs
Dim lngLCF As Long
Dim lngECF As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'get project guid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'save map
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) = vbNullString Then 'create it
    rstSavedMap.Fields.Append "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adBigInt
    rstSavedMap.Fields.Append "LCF", adBigInt
    rstSavedMap.Open
  Else
    'replace existing saved map
    rstSavedMap.Filter = "GUID<>'" & strGUID & "'"
    rstSavedMap.Open strSavedMap
    rstSavedMap.Save strSavedMap, adPersistADTG
    rstSavedMap.Open strSavedMap
  End If
  
  With cptSaveLocal_frm
    For lngItem = 0 To .lboMap.ListCount - 1
      If .lboMap.List(lngItem, 3) > 0 Then
        lngECF = .lboMap.List(lngItem, 0)
        lngLCF = .lboMap.List(lngItem, 3)
        rstSavedMap.AddNew Array(0, 1, 2), Array(strGUID, lngECF, lngLCF)
        'populate the fields
        For Each oTask In ActiveProject.Tasks
          On Error Resume Next
          If Len(oTask.GetField(lngLCF)) > 0 Then oTask.SetField lngLCF, ""
          If Len(oTask.GetField(lngECF)) > 0 Then
            oTask.SetField lngLCF, CStr(oTask.GetField(lngECF))
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            If oTask.GetField(lngLCF) <> CStr(oTask.GetField(lngECF)) Then
              MsgBox "There was an error copying from ECF " & CustomFieldGetName(lngECF) & " to LCF " & CustomFieldGetName(lngLCF) & "." & vbCrLf & vbCrLf & "Please validate data types.", vbExclamation + vbOKOnly, "Fail"
              GoTo next_mapping
            End If
          End If
        Next oTask
      End If
next_mapping:
    Next lngItem
  End With

  rstSavedMap.Save strSavedMap, adPersistADTG

  MsgBox "Enteprise Custom Fields saved locally.", vbInformation + vbOKOnly, "Complete"

exit_here:
  On Error Resume Next
  rstSavedMap.Close
  Set rstSavedMap = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptSaveLocal", Err, Erl)
  Resume exit_here
End Sub

Function cptInterrogateECF(ByRef oTask As Task, lngField As Long)
  'objects
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strPattern As String
  Dim strVal As String
  'longs
  Dim lngItem As Long
  Dim lngVal As Long
  'integers
  'doubles
  'booleans
  Dim blnVal As Boolean
  'variants
  'dates
  Dim dtVal As Date
  
  On Error Resume Next
  
  'check for outlinecode requirement (has parent-child structure)
  Set oOutlineCode = Application.GlobalOutlineCodes(FieldConstantToFieldName(lngField))
  If Not oOutlineCode Is Nothing Then
    If oOutlineCode.CodeMask.Count > 1 Then
      cptInterrogateECF = "Outline Code"
      GoTo exit_here
    Else
      If oOutlineCode.CodeMask(1).Sequence = 4 Then
        cptInterrogateECF = "Date"
      ElseIf oOutlineCode.CodeMask(1).Sequence = 5 Then
        cptInterrogateECF = "Cost"
      ElseIf oOutlineCode.CodeMask(1).Sequence = 7 Then
        cptInterrogateECF = "Number"
      Else
        cptInterrogateECF = "Text"
      End If
      GoTo exit_here
    End If
  End If
   
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
  ElseIf Err.Description = "This is not a valid lookup table value." Or Err.Description = "The value you entered does not exist in the lookup table of this code" Then
    'select the first value and check it
    oTask.SetField lngField, oOutlineCode.LookupTable(1).Name
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
  If Err.Number = 0 And Len(oTask.GetField(lngField)) = Len(CStr(lngVal)) Then
    cptInterrogateECF = "Number"
    GoTo exit_here
  End If
  
  'check for date
  On Error Resume Next
  dtVal = oTask.GetField(lngField)
  If Err.Number = 0 Then
    cptInterrogateECF = "Date"
    GoTo exit_here
  End If
  
  'could be flag
  If Len(cptRegEx(strVal, "Yes|No")) > 0 Then
    On Error Resume Next
    Set oOutlineCode = GlobalOutlineCodes(FieldConstantToFieldName(lngField))
    If oOutlineCode Is Nothing Then
      cptInterrogateECF = "MaybeFlag"
    Else
      cptInterrogateECF = "Text"
    End If
    GoTo exit_here
  End If
  
  On Error Resume Next
  strVal = oTask.GetField(lngField)
  'could be duration
  If strVal = DurationFormat(DurationValue(strVal), ActiveProject.DefaultDurationUnits) Then
    If Err.Number = 0 Then
      cptInterrogateECF = "Duration"
      GoTo exit_here
    End If
  End If
  
  'otherwise, it's most likely text
  cptInterrogateECF = "Text"

exit_here:
  On Error Resume Next
  Set oOutlineCode = Nothing
  
  Exit Function
err_here:
  Call cptHandleErr("foo", "bar", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Function

Sub cptGetAllFields(lngFrom As Long, lngTo As Long)
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
  'Dim lngTo As Long
  'Dim lngFrom As Long
  Dim lngFile As Long
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  GoTo exit_here
  
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
  'lngFrom = 215000001
  'lngTo = 218103807
  
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

