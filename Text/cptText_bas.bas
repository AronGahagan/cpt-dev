Attribute VB_Name = "cptText_bas"
'<cpt_version>v1.3.4</cpt_version>
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
  'objects
  Dim oTasks As Tasks, oTask As Task
  'strings
  Dim strAppend As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strAppend = InputBox("Append what text to selected tasks?", "Append Text")
  
  If Len(strAppend) = 0 Then Exit Sub
  
  Application.OpenUndoTransaction "Bulk Append"
  
  For Each oTask In oTasks
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask Is Nothing Then
      oTask.Name = Trim(oTask.Name) & " " & Trim(strAppend)
    End If
next_task:
  Next oTask
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptBulkAppend", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptBulkPrepend()
  'objects
  Dim oTasks As Tasks, oTask As Task
  'strings
  Dim strPrepend As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strPrepend = InputBox("Prepend what text to selected tasks?", "Prepend Text")
  
  If Len(strPrepend) = 0 Then Exit Sub
  
  Application.OpenUndoTransaction "Bulk Prepend"
  
  For Each oTask In oTasks
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask Is Nothing Then
       oTask.Name = Trim(strPrepend) & " " & Trim(oTask.Name)
    End If
next_task:
  Next oTask
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptBulkPrepend", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptEnumerate()
  'objects
  Dim oTasks As Tasks, oTask As Task
  'strings
  'longs
  Dim lngDigits As Long
  Dim lngEnumerate As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  'dates

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  vResponse = InputBox("How many digits (number input only)?", "Format Enumeration", 3)
  If StrPtr(vResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  End If
  lngDigits = CLng(vResponse)
  
  vResponse = InputBox("Start at what number (number input only)?", "Format Enumeration", 1)
  If StrPtr(vResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  End If
  lngEnumerate = CLng(vResponse)
  
  cptSpeed True
  
  Application.OpenUndoTransaction "Enumeration"
  
  If oTasks.Count > 2 Then
    For Each oTask In oTasks
      If oTask.ExternalTask Then GoTo next_task
      If Not oTask Is Nothing Then
        oTask.Name = oTask.Name & " (" & Format(lngEnumerate, String(lngDigits, "0")) & ")"
        lngEnumerate = lngEnumerate + 1
      End If
next_task:
    Next oTask
  End If
  cptSpeed False

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptEnumerate", Err, Erl)
  Resume exit_here

End Sub

Sub cptMyReplace()
  'fields affected: Marked, Task Name, Text Fields, Outline Code Fields
  'objects
  Dim rstReplaced As Object 'ADODB.Recordset
  Dim oTasks As Tasks, oTask As Task
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
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
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

  Set rstReplaced = CreateObject("ADODB.Recordset")
  rstReplaced.Fields.Append "UID", adBigInt
  rstReplaced.Open

  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    For Each vField In ActiveSelection.FieldIDList
      'limit to text fields
      If Len(cptRegEx(FieldConstantToFieldName(vField), "Text|Name")) > 0 Then
        If InStr(oTask.GetField(vField), CStr(vFind)) > 0 Then
          oTask.SetField vField, Replace(oTask.GetField(vField), CStr(vFind), CStr(vReplace))
          rstReplaced.AddNew Array("UID"), Array(oTask.UniqueID)
          rstReplaced.Update
          lngFound = lngFound + 1
        End If
      End If
    Next vField
next_task:
  Next oTask

  If lngFound = 0 Then
    MsgBox "No instances of '" & CStr(vFind) & "' found in selected cells.", vbExclamation + vbOKOnly, "MyReplace"
  Else
    rstReplaced.MoveFirst
    FilterEdit "cptMyReplace", True, True, True, False, , "Unique ID", , "equals", rstReplaced(0), "Or", True
    Do While Not rstReplaced.EOF
      FilterEdit "cptMyReplace", TaskFilter:=True, FieldName:="", newfieldname:="Unique ID", Test:="equals", Value:=rstReplaced(0), Operation:="Or", ShowInMenu:=True
      rstReplaced.MoveNext
    Loop
    FilterApply "cptMyReplace", True
    rstReplaced.MoveFirst
    Application.Find "Unique ID", "equals", rstReplaced(0)
    cptSpeed False
    strMsg = "Replaced " & Format(lngFound, "#,##0") & " instance" & IIf(lngFound = 1, "", "s") & " of '" & CStr(vFind) & "' with '" & CStr(vReplace) & "'" & vbCrLf & vbCrLf
    strMsg = strMsg & "Keep highlighted?"
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Replace") = vbNo Then
      cptSpeed True
      FilterApply "All Tasks", True
      Application.Find "Unique ID", "equals", rstReplaced(0)
      cptSpeed False
    End If
  End If
  
exit_here:
  On Error Resume Next
  If rstReplaced.State Then rstReplaced.Close
  Set rstReplaced = Nothing
  Application.CloseUndoTransaction
  cptSpeed False
  Set oTasks = Nothing
  Set oTask = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptMyReplace", Err, Erl)
  Resume exit_here

End Sub

Sub cptFindDuplicateTaskNames()
  'objects
  Dim oShell As Object
  Dim oExcel As Excel.Application
  Dim oWorkbook As Workbook
  Dim oWorksheet As Worksheet
  Dim oRange As Excel.Range
  Dim oListObject As ListObject
  'strings
  Dim strFileName As String
  'longs
  Dim lgNameCol As Long
  'integers
  'doubles
  'booleans
  Dim blnMaster As Boolean
  'variants
  'dates

  If ActiveProject.Tasks.Count = 0 Then GoTo exit_here
  If ActiveProject.Subprojects.Count > 0 Then blnMaster = True
  
  If Not cptCheckReference("Excel") Then GoTo exit_here

  On Error GoTo err_here
  If Not cptFilterExists("Active Tasks") Then
    FilterEdit Name:="Active Tasks", TaskFilter:=True, Create:=True, OverwriteExisting:=False, FieldName:="Active", Test:="equals", Value:="Yes", ShowInMenu:=True, showsummarytasks:=True
  End If
  MapEdit Name:="ExportTaskNames", Create:=True, OverwriteExisting:=True, DataCategory:=0, CategoryEnabled:=True, TableName:="Task_Table1", FieldName:="Unique ID", ExternalFieldName:="Unique_ID", ExportFilter:="Active Tasks", ImportMethod:=0, headerRow:=True, AssignmentData:=False, TextDelimiter:=Chr$(9), TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
  If blnMaster Then
    MapEdit Name:="ExportTaskNames", DataCategory:=0, FieldName:="Project", ExternalFieldName:="Project"
  End If
  MapEdit Name:="ExportTaskNames", DataCategory:=0, FieldName:="Summary", ExternalFieldName:="Summary"
  MapEdit Name:="ExportTaskNames", DataCategory:=0, FieldName:="Name", ExternalFieldName:="Name"
  Set oShell = CreateObject("WScript.Shell")
  strFileName = oShell.SpecialFolders("Desktop") & "\DuplicateTaskNames.xlsx"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  FileSaveAs Name:=strFileName, FormatID:="MSProject.ACE", Map:="ExportTaskNames"
  
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Open(strFileName)
  Set oWorksheet = oWorkbook.Sheets(1)
  
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(-4121)), , xlYes)
  
  oExcel.ActiveWindow.Zoom = 85
  oListObject.Range.Columns.AutoFit
  oListObject.TableStyle = ""
  Set oRange = oWorksheet.Range("Table1[Name]")
  oRange.FormatConditions.AddUniqueValues
  oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
  oRange.FormatConditions(1).DupeUnique = xlDuplicate
  With oRange.FormatConditions(1).Font
    .Color = -16383844
    .TintAndShade = 0
  End With
  With oRange.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
  End With
  oRange.FormatConditions(1).StopIfTrue = False
  'filter for duplicates
  lgNameCol = oWorksheet.Rows(1).Find("Name", lookat:=xlWhole).Column
  oListObject.Range.AutoFilter Field:=lgNameCol, Criteria1:=RGB(255, 199, 206), Operator:=xlFilterCellColor
  'sort by task name (to put duplicates together)
  oListObject.Sort.SortFields.Clear
  oListObject.Sort.SortFields.Add2 key:=oWorksheet.Range("Table1[[#All],[Name]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
  With oListObject.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With

exit_here:
  On Error Resume Next
  Set oShell = Nothing
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing
  Set oRange = Nothing
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Set oExcel = Nothing
  Set oListObject = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptText_bas", "cptFindDuplicateTaskNames", Err, Erl)
  Resume exit_here

End Sub

Sub cptTrimTaskNames()
  'objects
  Dim oTask As Task
  'strings
  'longs
  Dim lngBefore As Long
  Dim lngAfter As Long
  Dim lngCount As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True

  Application.OpenUndoTransaction "Trim Task Names"

  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then
      If oTask.ExternalTask Then GoTo next_task
      Application.StatusBar = "Trimming Task ID " & oTask.ID
      DoEvents
      lngBefore = Len(oTask.Name)
      'replace multi-spaces with single space
      oTask.Name = Replace(oTask.Name, cptRegEx(oTask.Name, "\s{2,}"), " ")
      'trim leading and trailing spaces
      oTask.Name = Trim(oTask.Name)
      lngAfter = Len(oTask.Name)
      If lngBefore > lngAfter Then lngCount = lngCount + 1
    End If
next_task:
  Next oTask

  Application.StatusBar = Format(lngCount, "#,##0") & " task names trimmed."

  MsgBox Format(lngCount, "#,##0") & " task names trimmed.", vbInformation + vbOKOnly, "Trim Task Names"

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Application.StatusBar = ""
  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptTrimTaskNames", Err, Erl)
  Resume exit_here

End Sub

Sub cptShowText_frm()
'objects
Dim oTasks As Tasks
Dim oTask As Task
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
  Set oTasks = ActiveSelection.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTasks Is Nothing Then
    cptText_frm.lboOutput.Clear
    For Each oTask In oTasks
      cptText_frm.lboOutput.AddItem
      cptText_frm.lboOutput.List(lngItem, 0) = oTask.UniqueID
      cptText_frm.lboOutput.List(lngItem, 1) = oTask.Name
      lngItem = lngItem + 1
    Next oTask
  End If

  Call cptStartEvents
  cptText_frm.Show
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing
  Set oTasks = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptShowText_frm", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptUpdatePreview(Optional strPrepend As String, Optional strAppend As String, Optional strPrefix As String, Optional lngCharacters As Long, Optional lngStartAt As Long, _
                  Optional lngCountBy As Long, Optional strSuffix As String, Optional strReplaceWhat As String, Optional strReplaceWith As String)
  'objects
  Dim oTask As Object
  'strings
  Dim strTaskName As String
  Dim strEnumerate As String
  'longs
  Dim lngItem As Long
  Dim lngEnumerate As Long
  'integers
  'booleans
  'variants
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lngItem = 0 To cptText_frm.lboOutput.ListCount - 1
    If IsNull(cptText_frm.lboOutput.List(lngItem, 0)) Then GoTo exit_here
    On Error Resume Next
    Set oTask = ActiveProject.Tasks.UniqueID(cptText_frm.lboOutput.List(lngItem, 0))
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If oTask Is Nothing Then
      If MsgBox("UID " & cptText_frm.lboOutput.List(lngItem, 0) & " not found in " & UCase(ActiveProject.Name) & "! Proceed?", vbCritical + vbYesNo, "Task Not Found") = vbNo Then
        Err.Clear
        GoTo exit_here
      Else
        GoTo next_item
      End If
    End If
    
    'start with the task name
    strTaskName = oTask.Name
    
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

      If lngStartAt = 0 Then
        If cptText_frm.txtStartAt.Value = "" Then
          lngStartAt = 1
          'cptText_frm.txtStartAt.Value = 1
        Else
          lngStartAt = CLng(cptText_frm.txtStartAt.Value)
        End If
      End If

      If lngCountBy = 0 Then
        If cptText_frm.txtCountBy.Value = "" Then
          lngCountBy = 1
          'cptText_frm.txtCountBy.Value = 1
        Else
          lngCountBy = CLng(cptText_frm.txtCountBy.Value)
        End If
      End If

      lngEnumerate = lngStartAt + (lngItem * lngCountBy)

      If lngCharacters = 0 Then
        If cptText_frm.txtCharacters.Value = "" Then
          lngCharacters = 1
          'cptText_frm.txtCharacters.Value = 1
        Else
          lngCharacters = CLng(cptText_frm.txtCharacters.Value)
        End If
      End If

      strEnumerate = strEnumerate & Format(lngEnumerate, String(lngCharacters, "0"))
      strEnumerate = strEnumerate & IIf(Len(strSuffix) > 0, strSuffix, cptText_frm.txtSuffix.Value)
      cptText_frm.lboOutput.List(lngItem, 1) = strTaskName & " " & strEnumerate
    Else
      cptText_frm.lboOutput.List(lngItem, 1) = strTaskName
    End If
    
    'replace
    '<issue27> added
    If Len(strReplaceWhat) = 0 Then strReplaceWhat = cptText_frm.txtReplaceWhat.Value
    If Len(strReplaceWith) = 0 Then strReplaceWith = cptText_frm.txtReplaceWith.Value
    If Len(strReplaceWhat) > 0 And Len(strReplaceWith) > 0 Then
      strTaskName = Replace(strTaskName, strReplaceWhat, strReplaceWith)
      cptText_frm.lboOutput.List(lngItem, 1) = strTaskName & " " & strEnumerate '</issue27>
    End If
next_item:
  Next lngItem

exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptUpdatePreview", Err, Erl)
  Resume exit_here

End Sub

Sub cptResetRowHeight()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

    '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "ResetRowHeight"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  SetRowHeight 1, "all"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptText_bas", "cptResetRowHeight", Err, Erl)
  Resume exit_here
End Sub
