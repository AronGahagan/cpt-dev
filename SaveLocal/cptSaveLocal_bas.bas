Attribute VB_Name = "cptSaveLocal_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Public strStartView As String
Public strStartTable As String
Public strStartFilter As String
Public strStartGroup As String

'todo: create view ECF:Local;ECF:Local;ECF:Local;
' -- based on saved map, current selections
'todo: handle resource custom fields - add toggle option and filter accordinlgy
'todo: process is: import with enterprise open, then save as mpp
'todo: make compatible with master/sub projects
'todo: handle when user changes custom fields manually -- onmouseover
'todo: code up the search filter
'todo: redesign form-lboECF wider; lboLocal more narrow
'todo: export map to excel

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
    
  'get project guid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'capture starting view/table/filter/group
  ActiveWindow.TopPane.Activate
  strStartView = ActiveProject.CurrentView
  strStartTable = ActiveProject.CurrentTable
  strStartFilter = ActiveProject.CurrentFilter
  strStartGroup = ActiveProject.CurrentGroup
  
  'create/overwrite the table
  cptSpeed True
  ViewApply "Gantt Chart"
  TableEditEx ".cptSaveLocal Task Table", True, True, True, , "ID", , , , , , True, , , , , , , , False
  TableEditEx ".cptSaveLocal Task Table", True, False, , , , "Unique ID", "UID", , , , True
  ActiveProject.Views(".cptSaveLocal Task View").Delete
  ViewEditSingle ".cptSaveLocal Task View", True, , pjTaskSheet, , , ".cptSaveLocal Task Table", "All Tasks", "No Group"
  ViewApply ".cptSaveLocal Task View"
  cptSpeed False
  
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
  
  'show the form?
  Set aTypes = CreateObject("System.Collections.SortedList")
  'record: field type, number of available custom fields
  For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
    aTypes.Add vType, 10
  Next
  aTypes.Add "Flag", 20
  aTypes.Add "Number", 20
  aTypes.Add "Text", 30
  
  'populate field types
  With cptSaveLocal_frm
    .cboFieldTypes.Clear
    For lngType = 0 To aTypes.Count - 1
      .cboFieldTypes.AddItem
      .cboFieldTypes.List(.cboFieldTypes.ListCount - 1, 0) = aTypes.getKey(lngType)
      .cboFieldTypes.List(.cboFieldTypes.ListCount - 1, 1) = aTypes.GetByIndex(lngType)
    Next lngType
    
    .cmdAutoMap.Visible = False
    .tglAutoMap = False
    .txtAutoMap.Visible = False
    .chkAutoSwitch = True
    .optTasks = True
    .cboFieldTypes.Value = "Text"
    
    .Show False
  End With
  
  'get enterprise custom task fields
  For lngField = 188776000 To 188778000 '2000 should do it for now
    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strECF = FieldConstantToFieldName(lngField)
      strEntity = cptInterrogateECF(oTask, lngField)
      rst.AddNew Array("GUID", "pjType", "ENTITY", "ECF_Constant", "ECF_Name"), Array(strGUID, pjTask, strEntity, lngField, FieldConstantToFieldName(lngField))
      lngECFCount = lngECFCount + 1
    End If
    cptSaveLocal_frm.lblStatus.Caption = "Analyzing Task ECFs...(" & Format(((lngField - 188776000) / (188778000 - 188776000)), "0%") & ")"
    cptSaveLocal_frm.lblProgress.Width = ((lngField - 188776000) / (188778000 - 188776000)) * cptSaveLocal_frm.lblStatus.Width
    DoEvents
  Next lngField

  'get enterprise custom resource fields
  For lngField = 205553664 To 205555664 '2000 should do it for now
    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strECF = FieldConstantToFieldName(lngField)
      strEntity = cptInterrogateECF(oTask, lngField)
      rst.AddNew Array("GUID", "pjType", "ENTITY", "ECF_Constant", "ECF_Name"), Array(strGUID, pjResource, strEntity, lngField, FieldConstantToFieldName(lngField))
      lngECFCount = lngECFCount + 1
    End If
    cptSaveLocal_frm.lblStatus.Caption = "Analyzing Resource ECFs...(" & Format((lngField - 205553664) / (205555664 - 205553664), "0%") & ")"
    cptSaveLocal_frm.lblProgress.Width = ((lngField - 205553664) / (205555664 - 205553664)) * cptSaveLocal_frm.lblStatus.Width
    DoEvents
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
    .lboECF.Clear
    If rst.RecordCount = 0 Then
      rst.Close
      MsgBox "No Enterprise Custom Fields available in this file.", vbExclamation + vbOKOnly, "No ECFs found"
      GoTo exit_here
    End If
    rst.MoveFirst
    Do While Not rst.EOF
      If UCase(rst("GUID")) = UCase(strGUID) And rst("pjType") = 0 Then
        .lboECF.AddItem
        .lboECF.List(.lboECF.ListCount - 1, 0) = rst("ECF_Constant")
        .lboECF.List(.lboECF.ListCount - 1, 1) = rst("ECF_Name")
        .lboECF.List(.lboECF.ListCount - 1, 2) = rst("ENTITY")
        If blnExists Then
          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & rst("ECF_Constant") '& " AND ENTITY=" & pjTask
          If Not rstSavedMap.EOF Then
            .lboECF.List(.lboECF.ListCount - 1, 3) = rstSavedMap("LCF")
            If Len(CustomFieldGetName(rstSavedMap("LCF"))) > 0 Then
              .lboECF.List(.lboECF.ListCount - 1, 4) = CustomFieldGetName(rstSavedMap("LCF"))
            Else
              .lboECF.List(.lboECF.ListCount - 1, 4) = FieldConstantToFieldName(rstSavedMap("LCF"))
            End If
            TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(rstSavedMap("ECF")), , , , , True, , , , , , , , False
            TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(rstSavedMap("LCF")), , , , , True, , , , , , , , False
            TableApply ".cptSaveLocal Task Table"
          End If
          rstSavedMap.Filter = ""
        End If
      End If
      rst.MoveNext
    Loop
    rst.Close
      
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
Dim oTasks As Object
Dim rstSavedMap As ADODB.Recordset
Dim oTask As Task
'strings
Dim strErrors As String
Dim strGUID As String
Dim strSavedMap As String
'longs
Dim lngTasks As Long
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
  
  'get total task count
  ActiveWindow.TopPane.Activate
  Call ViewApply("cpt_SaveLocal")
  FilterClear
  GroupClear
  OutlineShowAllTasks
  SelectAll
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTasks Is Nothing Then
    lngTasks = oTasks.Count
  Else
    MsgBox "There are no tasks in this schedule.", vbCritical + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If
  
  With cptSaveLocal_frm
    'todo: need to filter the lboECF for tasks
    For Each oTask In ActiveProject.Tasks
      On Error Resume Next
      For lngItem = 0 To .lboECF.ListCount - 1
        If .lboECF.List(lngItem, 3) > 0 Then
          lngECF = .lboECF.List(lngItem, 0)
          lngLCF = .lboECF.List(lngItem, 3)
          'no duplicates
          'todo: does Filter = X AND (Y OR Z) work?
          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF
          If rstSavedMap.RecordCount = 1 Then
            'overwrite it
            rstSavedMap.Delete adAffectCurrent
          End If
          rstSavedMap.Filter = ""
          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND LCF=" & lngLCF
          If rstSavedMap.RecordCount = 1 Then
            'overwrite it
            rstSavedMap.Delete adAffectCurrent
          End If
          rstSavedMap.Filter = ""
          'add the new record
          rstSavedMap.AddNew Array(0, 1, 2), Array(strGUID, lngECF, lngLCF)
          'first clear the values
          If Len(oTask.GetField(lngLCF)) > 0 Then oTask.SetField lngLCF, ""
          'if ECF is formula, then skip it
          If Len(CustomFieldGetFormula(lngECF)) > 0 Then GoTo next_mapping
          If Len(oTask.GetField(lngECF)) > 0 Then
            oTask.SetField lngLCF, CStr(oTask.GetField(lngECF))
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            If oTask.GetField(lngLCF) <> CStr(oTask.GetField(lngECF)) Then
              If MsgBox("There was an error copying from ECF " & CustomFieldGetName(lngECF) & " to LCF " & CustomFieldGetName(lngLCF) & " on Task UID " & oTask.UniqueID & "." & vbCrLf & vbCrLf & "Please validate data type mapping." & vbCrLf & vbCrLf & "Proceed anyway?", vbExclamation + vbYesNo, "Failed!") = vbNo Then
                GoTo exit_here
              End If
            End If
          End If
        End If
next_mapping:
      Next lngItem
      
    Next oTask
  End With

  'todo: resource ECF > LCF

  rstSavedMap.Save strSavedMap, adPersistADTG

  MsgBox "Enteprise Custom Fields saved locally.", vbInformation + vbOKOnly, "Complete"

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
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

Sub cptAnalyzeAutoMap()
  'objects
  Dim rstAvailable As ADODB.Recordset
  Dim aTypes As SortedList
  'strings
  Dim strMsg As String
  'longs
  Dim lngItem2 As Long
  Dim lngAvailable As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vType As Variant
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set rstAvailable = CreateObject("ADODB.Recordset")
  With rstAvailable
    .Fields.Append "TYPE", adVarChar, 50
    .Fields.Append "ECF", adInteger
    .Fields.Append "LCF", adInteger
    .Open

    Set aTypes = CreateObject("System.Collections.SortedList")
    'record: field type, number of available custom fields
    For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
      aTypes.Add vType, 10
      .AddNew Array(0, 1, 2), Array(vType, 0, 10)
    Next
    aTypes.Add "Flag", 20
    .AddNew Array(0, 1, 2), Array("Flag", 0, 20)
    aTypes.Add "Number", 20
    .AddNew Array(0, 1, 2), Array("Number", 0, 20)
    aTypes.Add "Text", 30
    .AddNew Array(0, 1, 2), Array("Text", 0, 30)
    .Update
    .Sort = "TYPE"
    
    'todo: start->date;finish->date;date->date
    
    'get available LCF
    For lngItem = 0 To aTypes.Count - 1
      For lngItem2 = 1 To aTypes.getValueList()(lngItem)
        'todo: account for both pjTask and pjResource
        If Len(CustomFieldGetName(FieldNameToFieldConstant(aTypes.getKey(lngItem) & lngItem2))) > 0 Then
          .MoveFirst
          .Find "TYPE='" & aTypes.getKey(lngItem) & "'"
          If Not .EOF Then
            .Fields(2) = .Fields(2) - 1
          End If
        End If
      Next lngItem2
    Next lngItem
    
    'get total ECF
    For lngItem = 0 To cptSaveLocal_frm.lboECF.ListCount - 1
      If cptSaveLocal_frm.lboECF.Selected(lngItem) Then
        .MoveFirst
        .Find "TYPE='" & Replace(cptSaveLocal_frm.lboECF.List(lngItem, 2), "Maybe", "") & "'"
        If Not .EOF Then
          If IsNull(cptSaveLocal_frm.lboECF.List(cptSaveLocal_frm.lboECF.ListIndex, 3)) Then
            'only count unmapped
            .Fields(1) = .Fields(1) + 1
          End If
        End If
      End If
    Next lngItem
    
    'return result
    strMsg = strMsg & String(34, "-") & vbCrLf
    strMsg = strMsg & "| " & "TYPE" & String(10, " ") & "|"
    strMsg = strMsg & " ECF |"
    strMsg = strMsg & " LCF |"
    strMsg = strMsg & " <> |" & vbCrLf
    strMsg = strMsg & String(34, "-") & vbCrLf
    .MoveFirst
    Do While Not .EOF
      strMsg = strMsg & "| " & rstAvailable(0) & String(14 - Len(rstAvailable(0)), " ") & "|"
      If rstAvailable(0) = "Start" Or rstAvailable(0) = "Finish" Then
        strMsg = strMsg & "   - |"
      Else
        strMsg = strMsg & String(4 - Len(CStr(rstAvailable(1))), " ") & rstAvailable(1) & " |"
      End If
      strMsg = strMsg & String(4 - Len(CStr(rstAvailable(2))), " ") & rstAvailable(2) & " |"
      strMsg = strMsg & IIf(rstAvailable(2) >= rstAvailable(1), " ok ", "  X ") & "|" & vbCrLf
      .MoveNext
    Loop
    strMsg = strMsg & String(34, "-") & vbCrLf
    cptSaveLocal_frm.cmdAutoMap.Visible = False
    If InStr(strMsg, "  X ") > 0 Then
      strMsg = strMsg & "AutoMap is NOT available."
    Else
      strMsg = strMsg & "AutoMap IS available."
      If cptSaveLocal_frm.tglAutoMap Then
        cptSaveLocal_frm.cmdAutoMap.Visible = True
      Else
        cptSaveLocal_frm.cmdAutoMap.Visible = False
      End If
    End If
    
    cptSaveLocal_frm.txtAutoMap.Value = strMsg
    
    .Close
    
  End With
  
exit_here:
  On Error Resume Next
  rstAvailable.Close
  Set rstAvailable = Nothing
  aTypes.Clear
  Set aTypes = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptAutoMap", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptAutoMap()
  'objects
  'strings
  'longs
  Dim lngECF As Long
  Dim lngLCF As Long
  Dim lngLCFs As Long
  Dim lngECFs As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'todo: unselect after complete - if fails, leave selected
  'todo: update rstSavedMap after each
  'todo: update lboECF after each...or lboLCF?
  'todo: for dates > 10 cycle through date, start, finish

  With cptSaveLocal_frm
    .lblStatus.Caption = "AutoMapping..."
    'loop through ECFs looking for selected ECFs to map
    For lngECFs = 0 To .lboECF.ListCount - 1
      If .lboECF.Selected(lngECFs) Then
        lngECF = .lboECF.List(lngECFs, 0)
        'switch cbo types to get list of lngLCFs
        If .cboFieldTypes <> .lboECF.List(lngECFs, 2) Then .cboFieldTypes = .lboECF.List(lngECFs, 2)
        'loop through LCFs looking for one available
        For lngLCFs = 0 To .lboLCF.ListCount - 1
          lngLCF = .lboLCF.List(lngLCFs, 0)
          If Len(CustomFieldGetName(lngLCF)) = 0 Then
            Call cptMapECFtoLCF(lngECF, lngLCF)
            .lboECF.List(lngECFs, 3) = lngLCF
            .lboECF.List(lngECFs, 4) = CustomFieldGetName(lngLCF)
            Call cptAnalyzeAutoMap
            Exit For
          End If
        Next lngLCFs
      End If
      .lblProgress.Width = (lngECFs / (.lboECF.ListCount - 1)) * .lblStatus.Width
    Next lngECFs
    .lblStatus.Caption = "AutoMap complete."
    .lblProgress.Width = .lblStatus.Width
    
    If MsgBox("Fields AutoMapped. Import field data now?", vbQuestion + vbYesNo, "Save Local") = vbYes Then
      .cmdSaveLocal.SetFocus
    End If
    
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptAutoMap", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptMapECFtoLCF(lngECF As Long, lngLCF As Long)
  'objects
  Dim rstSavedMap As ADODB.Recordset
  Dim oLookupTableEntry As LookupTableEntry
  Dim oOutlineCodeLocal As OutlineCode
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strGUID As String
  Dim strSavedMap As String
  Dim strECF As String
  'longs
  Dim lngItem As Long
  Dim lngDown As Long
  Dim lngCodeNumber As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  With cptSaveLocal_frm
    'if already mapped then prompt with ECF name and ask to remap
    For lngItem = 0 To .lboECF.ListCount - 1
      If .lboECF.List(lngItem, 3) = lngLCF Then
        If MsgBox(FieldConstantToFieldName(lngLCF) & " is already mapped to " & .lboECF.List(lngItem, 1) & " - reassign it?", vbExclamation + vbYesNo, "Already Mapped") = vbYes Then
          CustomFieldDelete lngLCF
          .lboECF.List(lngItem, 3) = ""
          .lboECF.List(lngItem, 4) = ""
        Else
          GoTo exit_here
        End If
      End If
    Next lngItem
    
    'capture rename
    If Len(CustomFieldGetName(lngLCF)) > 0 Then
      If MsgBox("Rename " & FieldConstantToFieldName(lngLCF) & " to " & FieldConstantToFieldName(lngECF) & "?", vbQuestion + vbYesNo, "Please confirm") = vbYes Then
        'rename it
        CustomFieldRename CLng(lngLCF), CustomFieldGetName(lngECF) & " (" & FieldConstantToFieldName(lngLCF) & ")"
        'rename in lboLCF
        If Not .tglAutoMap Then .lboLCF.List(.lboLCF.ListIndex, 1) = FieldConstantToFieldName(lngLCF) & " (" & CustomFieldGetName(lngLCF) & ")"
      Else
        GoTo exit_here
      End If
    Else
      ActiveWindow.TopPane.Activate
      'rename it in msp
      CustomFieldRename lngLCF, CustomFieldGetName(lngECF) & " (" & FieldConstantToFieldName(lngLCF) & ")"
      'rename it in lboLCF
      If Not .tglAutoMap Then .lboLCF.List(.lboLCF.ListIndex, 1) = FieldConstantToFieldName(.lboLCF) & " (" & CustomFieldGetName(.lboLCF) & ")"
    End If
    
    'get formula
    If Len(CustomFieldGetFormula(lngECF)) > 0 Then
      CustomFieldSetFormula lngLCF, CustomFieldGetFormula(lngECF)
    End If
    
    'get indicators
    'todo: warn user these are not exposed/available
    
    'get pick list
    strECF = CustomFieldGetName(lngECF)
    On Error Resume Next
    Set oOutlineCode = GlobalOutlineCodes(strECF)
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Not oOutlineCode Is Nothing Then
      'make it a picklist
      CustomFieldPropertiesEx lngLCF, pjFieldAttributeValueList
      If oOutlineCode.CodeMask.Count > 1 Then 'import outline code and all settings
'        MsgBox "If copying down an Outline Code, please use the 'Import Field' function of the Custom Fields dialog before clicking Save Local.", vbInformation + vbOKOnly, "Nota Bene"
'        VBA.SendKeys "%r", True
'        VBA.SendKeys "f", True
'        VBA.SendKeys "%y", True
'        VBA.SendKeys "o", True
'        VBA.SendKeys "{TAB}"
'        'repeat the down key based on which code
'        lngCodeNumber = CLng(Replace(FieldConstantToFieldName(lngLCF), "Outline Code", ""))
'        If lngCodeNumber > 1 Then
'          For lngDown = 1 To lngCodeNumber - 1
'            VBA.SendKeys "{DOWN}", True
'          Next lngDown
'        End If
'        VBA.SendKeys "%i", True
'        VBA.SendKeys "%f", True
'        VBA.SendKeys "%{DOWN}", True
'        VBA.SendKeys Left(FieldConstantToFieldName(.lboECF.List(.lboECF.ListIndex, 0)), 1), True
        'capture code mask
        With oOutlineCode.CodeMask
          For lngItem = 1 To .Count
            CustomOutlineCodeEditEx lngLCF, .Item(lngItem).Level, .Item(lngItem).Sequence, .Item(lngItem).Length, .Item(lngItem).Separator
          Next lngItem
        End With
        'capture picklist
        Set oOutlineCodeLocal = ActiveProject.OutlineCodes(CustomFieldGetName(lngLCF))
        With oOutlineCode.LookupTable
          'load items bottom to top
          For lngItem = .Count To 1 Step -1
            Set oLookupTableEntry = oOutlineCodeLocal.LookupTable.AddChild(.Item(lngItem).Name)
            oLookupTableEntry.Description = .Item(lngItem).Description
          Next lngItem
          'indent top to bottom
          For lngItem = 1 To .Count
            oOutlineCodeLocal.LookupTable.Item(lngItem).Level = .Item(lngItem).Level
          Next lngItem
        End With
        'capture other options
        CustomOutlineCodeEditEx FieldID:=lngLCF, OnlyLookUpTableCodes:=oOutlineCode.OnlyLookUpTableCodes, OnlyCompleteCodes:=oOutlineCode.OnlyCompleteCodes
        CustomOutlineCodeEditEx FieldID:=lngLCF, OnlyLeaves:=oOutlineCode.OnlyLeaves
        'todo: next line for RequiredCode throws an error, not sure why
        'CustomOutlineCodeEditEx FieldID:=lngLCF, RequiredCode:=oOutlineCode.RequiredCode
        If oOutlineCode.DefaultValue <> "" Then CustomOutlineCodeEditEx FieldID:=lngLCF, DefaultValue:=oOutlineCode.DefaultValue
        CustomOutlineCodeEditEx FieldID:=lngLCF, SortOrder:=oOutlineCode.SortOrder
      Else 'import just the pick list
        For lngItem = 1 To oOutlineCode.LookupTable.Count
          CustomFieldValueListAdd lngLCF, oOutlineCode.LookupTable(lngItem).Name, oOutlineCode.LookupTable(lngItem).Description
        Next lngItem
        
      End If
    End If
    If Not .tglAutoMap Then
      .lboECF.List(.lboECF.ListIndex, 3) = lngLCF
      .lboECF.List(.lboECF.ListIndex, 4) = CustomFieldGetName(lngLCF)
    End If
  End With
  
  'update rstSavedMap
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) <> vbNullString Then
    rstSavedMap.Open strSavedMap
    rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF
    If Not rstSavedMap.EOF Then
      rstSavedMap.Fields(2) = lngLCF
    Else
      rstSavedMap.AddNew Array(0, 1, 2), Array(strGUID, lngECF, lngLCF)
    End If
    rstSavedMap.Filter = ""
    rstSavedMap.Save strSavedMap, adPersistADTG
  Else 'create it
    rstSavedMap.Fields.Append "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adInteger
    rstSavedMap.Fields.Append "LCF", adInteger
    rstSavedMap.Open
    rstSavedMap.AddNew Array(0, 1, 2), Array(strGUID, lngECF, lngLCF)
    rstSavedMap.Save strSavedMap, adPersistADTG
  End If
  rstSavedMap.Close
  
  'update the table
  TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False
  TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False
  TableApply ".cptSaveLocal Task Table"
  
exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing
  Set oLookupTableEntry = Nothing
  Set oOutlineCodeLocal = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptMapECFtoLCF", Err, Erl)
  Resume exit_here
End Sub

