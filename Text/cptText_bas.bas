Attribute VB_Name = "cptText_bas"
'<cpt_version>v1.1.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptReplicateProcess()
  MsgBox "feature not yet released", vbOKOnly + vbInformation, "todo"
  'select a process (group of tasks)
  'define the sequence
  'provide number of units
  'replicate the process
  'define products; define steps; define count; replicate
End Sub

Sub cptBulkAppend()
Dim Tasks As Tasks, Task As Task, strAppend As String

  On Error Resume Next
  Set Tasks = ActiveSelection.Tasks
  If Tasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strAppend = InputBox("Append what text to selected tasks?", "Append Text")
  
  If Len(strAppend) = 0 Then Exit Sub
  
  Application.OpenUndoTransaction "Bulk Append"
  
  For Each Task In Tasks
    If Task.ExternalTask Then GoTo next_task
    If Not Task Is Nothing Then
      Task.Name = Trim(Task.Name) & " " & Trim(strAppend)
    End If
next_task:
  Next Task
  
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set Task = Nothing
  Set Tasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptBulkAppend", err)
  Resume exit_here
  
End Sub

Sub cptBulkPrepend()
Dim Tasks As Tasks, Task As Task, strPrepend As String

  On Error Resume Next
  Set Tasks = ActiveSelection.Tasks
  If Tasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strPrepend = InputBox("Prepend what text to selected tasks?", "Prepend Text")
  
  If Len(strPrepend) = 0 Then Exit Sub
  
  Application.OpenUndoTransaction "Bulk Prepend"
  
  For Each Task In ActiveSelection.Tasks
    If Task.ExternalTask Then GoTo next_task
    If Not Task Is Nothing Then
       Task.Name = Trim(strPrepend) & " " & Trim(Task.Name)
    End If
next_task:
  Next Task
  
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set Task = Nothing
  Set Tasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("basTextTool", "cptBulkPrepend", err)
  Resume exit_here
  
End Sub

Sub cptEnumerate()
Dim Tasks As Tasks, Task As Task, lgDigits As Long
Dim vbResponse As Variant, lgEnumerate As Long

  On Error Resume Next
  Set Tasks = ActiveSelection.Tasks
  If Tasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  vbResponse = InputBox("How many digits (number input only)?", "Format Enueration", 3)
  If StrPtr(vbResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vbResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  End If
  lgDigits = CLng(vbResponse)
  
  vbResponse = InputBox("Start at what number (number input only)?", "Format Enumeration", 1)
  If StrPtr(vbResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vbResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  End If
  lgEnumerate = CLng(vbResponse)
  
  cptSpeed True
  
  Application.OpenUndoTransaction "Enumeration"
  
  If Tasks.count > 2 Then
    For Each Task In Tasks
      If Task.ExternalTask Then GoTo next_task
      If Not Task Is Nothing Then
        Task.Name = Task.Name & " (" & Format(lgEnumerate, String(lgDigits, "0")) & ")"
        lgEnumerate = lgEnumerate + 1
      End If
next_task:
    Next
  End If
  cptSpeed False


exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set Task = Nothing
  Set Tasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptEnumerate", err)
  Resume exit_here

End Sub

Sub cptMyReplace()
'fields affected: Marked, Task Name, Text Fields, Outline Code Fields
'objects
Dim arrReplaced As Object
Dim Tasks As Tasks, Task As Task
'strings
Dim strMsg As String
'longs
Dim lngItem As Long
Dim lngFound As Long
'integers
'doubles
'booleans
'variants
Dim vField As Variant, vFind As Variant, vReplace As Variant
'dates

  On Error Resume Next
  cptSpeed True
  Set Tasks = ActiveSelection.Tasks
  If Tasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'get string to find
  vFind = InputBox("Find what text:", "Replace")
  If StrPtr(vFind) = 0 Then GoTo exit_here 'user hit cancel
  vFind = Trim(vFind)
  
  'get string to replace it with
  vReplace = InputBox("Replace '" & CStr(vFind) & "' with what text:", "Replace")
  If StrPtr(vReplace) = 0 Then GoTo exit_here 'user hit cancel
  vReplace = Trim(vReplace)
  
  Application.OpenUndoTransaction "MyReplace"

  Set arrReplaced = CreateObject("System.Collections.SortedList")

  For Each Task In Tasks
    If Task Is Nothing Then GoTo next_task
    If Task.ExternalTask Then GoTo next_task
    For Each vField In ActiveSelection.FieldIDList
      'limit to text fields
      If Len(cptRegEx(FieldConstantToFieldName(vField), "Text|Name")) > 0 Then
        If InStr(Task.GetField(vField), CStr(vFind)) > 0 Then
          Task.SetField vField, Replace(Task.GetField(vField), CStr(vFind), CStr(vReplace))
          arrReplaced.Add Task.UniqueID, Task.UniqueID
          lngFound = lngFound + 1
        End If
      End If
    Next vField
next_task:
  Next Task

  If lngFound = 0 Then
    MsgBox "No instances of '" & CStr(vFind) & "' found in selected cells.", vbExclamation + vbOKOnly, "MyReplace"
  Else
    FilterEdit "cptMyReplace", True, True, True, False, , "Unique ID", , "equals", arrReplaced.getKey(0), "Or", True
    For lngItem = 1 To arrReplaced.count - 1
      FilterEdit "cptMyReplace", Taskfilter:=True, FieldName:="", NewFieldName:="Unique ID", test:="equals", Value:=arrReplaced.getKey(lngItem), operation:="Or", ShowInMenu:=True
    Next lngItem
    FilterApply "cptMyReplace", True
    Application.Find "Unique ID", "equals", arrReplaced.getKey(0)
    cptSpeed False
    strMsg = "Replaced " & Format(lngFound, "#,##0") & " instance" & IIf(lngFound = 1, "", "s") & " of '" & CStr(vFind) & "' with '" & CStr(vReplace) & "'" & vbCrLf & vbCrLf
    strMsg = strMsg & "Keep highlighted?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Replace") = vbNo Then
      cptSpeed True
      FilterApply "All Tasks", True
      Application.Find "Unique ID", "equals", arrReplaced.getKey(0)
      cptSpeed False
    End If
  End If
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Set arrReplaced = Nothing
  Set Tasks = Nothing
  Set Task = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptMyReplace", err)
  Resume exit_here

End Sub

Sub cptFindDuplicateTaskNames()
'requires: msexcel
'objects
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Excel.Range, ListObject As ListObject
'string
Dim strFileName As String
'boolean
Dim blnMaster As Boolean
'longs
Dim lgNameCol As Long

  If ActiveProject.Tasks.count = 0 Then GoTo exit_here
  If ActiveProject.Subprojects.count > 0 Then blnMaster = True
  
  If Not cptCheckReference("Excel") Then GoTo exit_here

  On Error GoTo err_here
  MapEdit Name:="ExportTaskNames", Create:=True, OverwriteExisting:=True, DataCategory:=0, CategoryEnabled:=True, TableName:="Task_Table1", FieldName:="Unique ID", ExternalFieldName:="Unique_ID", ExportFilter:="All Tasks", ImportMethod:=0, headerRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(9), TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
  If blnMaster Then
    MapEdit Name:="ExportTaskNames", DataCategory:=0, FieldName:="Project", ExternalFieldName:="Project"
  End If
  MapEdit Name:="ExportTaskNames", DataCategory:=0, FieldName:="Summary", ExternalFieldName:="Summary"
  MapEdit Name:="ExportTaskNames", DataCategory:=0, FieldName:="Name", ExternalFieldName:="Name"
  strFileName = Environ("USERPROFILE") & "\Desktop\DuplicateTaskNames.xlsx"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  FileSaveAs Name:=strFileName, FormatID:="MSProject.ACE", Map:="ExportTaskNames"
  
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Open(strFileName)
  Set Worksheet = Workbook.Sheets(1)
  
  Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), , xlYes)
  
  xlApp.ActiveWindow.Zoom = 85
  ListObject.Range.Columns.AutoFit
  ListObject.TableStyle = ""
  Set rng = Worksheet.Range("Table1[Name]")
  rng.FormatConditions.AddUniqueValues
  rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
  rng.FormatConditions(1).DupeUnique = xlDuplicate
  With rng.FormatConditions(1).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With rng.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False
  'filter for duplicates
  lgNameCol = Worksheet.Rows(1).Find("Name", lookat:=xlWhole).Column
  ListObject.Range.AutoFilter Field:=lgNameCol, Criteria1:=RGB(255, 199, 206), Operator:=xlFilterCellColor
  'sort by task name (to put duplicates together)
  ListObject.Sort.SortFields.Clear
  ListObject.Sort.SortFields.Add2 key:=Worksheet.Range("Table1[[#All],[Name]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
  With ListObject.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With

exit_here:
  On Error Resume Next
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set rng = Nothing
  If Not xlApp Is Nothing Then xlApp.Visible = True
  Set xlApp = Nothing
  Set ListObject = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptFindDuplicateTaskNames", err)
  Resume exit_here

End Sub

Sub cptTrimTaskNames()
Dim Task As Task, lgBefore As Long, lgAfter As Long, lgCount As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.OpenUndoTransaction "Trim Task Names"

  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then
      lgBefore = Len(Task.Name)
      Task.Name = Trim(Task.Name)
      lgAfter = Len(Task.Name)
      If lgBefore > lgAfter Then lgCount = lgCount + 1
    End If
  Next Task


  MsgBox Format(lgCount, "#,##0") & " task names trimmed.", vbInformation + vbOKOnly, "Trim Task Names"

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptTrimTaskNames", err)
  Resume exit_here

End Sub

Sub ShowCptText_frm()
'objects
Dim Tasks As Tasks, Task As Task
'strings
'longs
Dim lngItem As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptModuleExists("cptText_frm") Then GoTo exit_here

  On Error Resume Next
  Set Tasks = ActiveSelection.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not Tasks Is Nothing Then
    cptText_frm.lboOutput.Clear
    For Each Task In Tasks
      cptText_frm.lboOutput.AddItem
      cptText_frm.lboOutput.List(lngItem, 0) = Task.UniqueID
      cptText_frm.lboOutput.List(lngItem, 1) = Task.Name
      lngItem = lngItem + 1
    Next Task
  End If

  Call cptStartEvents
  cptText_frm.Show
  
exit_here:
  On Error Resume Next
  Set Task = Nothing
  Set Tasks = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "ShowcptText_frm", err)
  Resume exit_here
  
End Sub

Sub cptUpdatePreview(Optional strPrepend As String, Optional strAppend As String, Optional strPrefix As String, Optional lgCharacters As Long, Optional lgStartAt As Long, _
                  Optional lgCountBy As Long, Optional strSuffix As String, Optional strReplaceWhat As String, Optional strReplaceWith As String)
'objects
Dim Task As Object
'strings
Dim strTaskName As String, strEnumerate As String
'longs
Dim lngItem As Long, lgEnumerate As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lngItem = 0 To cptText_frm.lboOutput.ListCount - 1
    On Error Resume Next
    Set Task = ActiveProject.Tasks.UniqueID(cptText_frm.lboOutput.List(lngItem, 0))
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Task Is Nothing Then
      If MsgBox("UID " & cptText_frm.lboOutput.List(lngItem, 0) & " not found in " & UCase(ActiveProject.Name) & "! Proceed?", vbCritical + vbYesNo, "Task Not Found") = vbNo Then
        err.Clear
        GoTo exit_here
      Else
        GoTo next_item
      End If
    End If
    strTaskName = Task.Name
    If Len(strPrepend) > 0 Then
      strTaskName = Trim(strPrepend) & " " & strTaskName
    ElseIf Len(cptText_frm.txtPrepend.Value) > 0 Then
      strTaskName = Trim(cptText_frm.txtPrepend.Value) & " " & strTaskName
    End If
    If Len(strAppend) > 0 Then
      strTaskName = Trim(strTaskName) & " " & Trim(strAppend)
    ElseIf Len(cptText_frm.txtAppend.Value) > 0 Then
      strTaskName = Trim(strTaskName) & " " & Trim(cptText_frm.txtAppend.Value)
    End If
    cptText_frm.chkIsDirty = cptText_frm.CheckDirty
    If cptText_frm.chkIsDirty Then
      strEnumerate = IIf(Len(strPrefix) > 0, strPrefix, cptText_frm.txtPrefix.Value)

      If lgStartAt = 0 Then
        If cptText_frm.txtStartAt.Value = "" Then
          lgStartAt = 1
          'cptText_frm.txtStartAt.Value = 1
        Else
          lgStartAt = CLng(cptText_frm.txtStartAt.Value)
        End If
      End If

      If lgCountBy = 0 Then
        If cptText_frm.txtCountBy.Value = "" Then
          lgCountBy = 1
          'cptText_frm.txtCountBy.Value = 1
        Else
          lgCountBy = CLng(cptText_frm.txtCountBy.Value)
        End If
      End If

      lgEnumerate = lgStartAt + (lngItem * lgCountBy)

      If lgCharacters = 0 Then
        If cptText_frm.txtCharacters.Value = "" Then
          lgCharacters = 1
          'cptText_frm.txtCharacters.Value = 1
        Else
          lgCharacters = CLng(cptText_frm.txtCharacters.Value)
        End If
      End If

      strEnumerate = strEnumerate & Format(lgEnumerate, String(lgCharacters, "0"))
      strEnumerate = strEnumerate & IIf(Len(strSuffix) > 0, strSuffix, cptText_frm.txtSuffix.Value)
      cptText_frm.lboOutput.List(lngItem, 1) = strTaskName & " " & strEnumerate
    Else
      cptText_frm.lboOutput.List(lngItem, 1) = strTaskName
    End If

    'replace
    If Len(strReplaceWhat) > 0 And Len(strReplaceWith) > 0 Then
      strTaskName = Replace(strTaskName, strReplaceWhat, strReplaceWith)
      cptText_frm.lboOutput.List(lngItem, 1) = strTaskName
    End If
next_item:
  Next lngItem

exit_here:
  On Error Resume Next
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptUpdatePreview", err)
  Resume exit_here

End Sub



