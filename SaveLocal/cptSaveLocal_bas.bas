Attribute VB_Name = "cptSaveLocal_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Public strStartView As String
Public strStartTable As String
Public strStartFilter As String
Public strStartGroup As String

'todo: process is: import with enterprise open, then save as mpp
'todo: make compatible with master/sub projects
'todo: handle when user changes custom fields manually

Sub cptShowSaveLocalForm()
'objects
Dim oRange As Excel.Range
Dim oListObject As ListObject
Dim aProjects As ArrayList
Dim oSubproject As SubProject
Dim oMasterProject As Project
Dim oWorksheet As Worksheet
Dim oWorkbook As Workbook
Dim oExcel As Excel.Application
Dim oTask As Task
Dim rstSavedMap As ADODB.Recordset
Dim aTypes As Object
Dim rst As ADODB.Recordset
'strings
Dim strURL As String
Dim strSaved As String
Dim strEntity As String
Dim strGUID As String
Dim strECF As String
'longs
Dim lngECF As Long
Dim lngLCF As Long
Dim lngMismatchCount As Long
Dim lngLastRow As Long
Dim lngSubproject As Long
Dim lngProject As Long
Dim lngSubprojectCount As Long
Dim lngField As Long
Dim lngFields As Long
Dim lngType As Long
Dim lngECFCount As Long
'integers
'doubles
'booleans
'variants
Dim vLine As Variant
Dim vEntity As Variant
Dim vType As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'get server URL
  If Len(ActiveProject.ServerURL) = 0 Then
    MsgBox "You are not connected to a Project Server.", vbCritical + vbOKOnly, "No Server"
    GoTo exit_here
  Else
    strURL = ActiveProject.ServerURL
  End If
  
  'setup array of types/counts
  Set aTypes = CreateObject("System.Collections.SortedList")
  'record: field type, number of available custom fields
  For Each vType In Array("Cost", "Date", "Duration", "Finish", "Start", "Outline Code")
    aTypes.Add vType, 10
  Next
  aTypes.Add "Flag", 20
  aTypes.Add "Number", 20
  aTypes.Add "Text", 30
  
  'if master/sub then ensure LCFs match
  lngSubprojectCount = ActiveProject.Subprojects.Count
  If lngSubprojectCount > 0 Then
    If MsgBox(lngSubprojectCount & " subproject(s) found." & vbCrLf & vbCrLf & "It is highly recommended that you analyze master/sub LCF matches." & vbCrLf & vbCrLf & "Do it now?", vbExclamation + vbYesNo, "Master/Sub Detected") = vbNo Then GoTo skip_it
    Set oMasterProject = ActiveProject
    Application.StatusBar = "Setting up Excel..."
    'set up Excel
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = False
    Set oWorkbook = oExcel.Workbooks.Add
    oExcel.ScreenUpdating = False
    oExcel.Calculation = xlCalculationManual
    Set oWorksheet = oWorkbook.Sheets(1)
    oExcel.ActiveWindow.Zoom = 85
    oExcel.ActiveWindow.SplitRow = 1
    oExcel.ActiveWindow.SplitColumn = 4
    oExcel.ActiveWindow.FreezePanes = True
    oWorksheet.Name = "Sync"
    'set up headers
    oWorksheet.[A1:D1] = Array("ENTITY", "TYPE", "CONSTANT", "NAME")
    'capture master and subproject names
    oWorksheet.Cells(1, 5) = oMasterProject.Name
    oWorksheet.Columns.AutoFit
    cptSpeed True
    Application.StatusBar = "Opening subprojects..."
    Set aProjects = CreateObject("System.Collections.ArrayList")
    aProjects.Add oMasterProject.Name
    For Each oSubproject In oMasterProject.Subprojects
      FileOpenEx oSubproject.SourceProject.FullName, True
      aProjects.Add ActiveProject.Name
    Next oSubproject
    For lngProject = 0 To aProjects.Count - 1
      Application.StatusBar = "Analyzing " & aProjects(lngProject) & "..."
      DoEvents
      lngLastRow = 1
      Projects(aProjects(lngProject)).Activate
      oWorksheet.Cells(1, 5 + lngProject) = aProjects(lngProject)
      For Each vEntity In Array(pjTask, pjResource)
        For lngType = 0 To aTypes.Count - 1
          For lngField = 1 To aTypes.getByIndex(lngType)
            lngLastRow = lngLastRow + 1
            If lngProject = 0 Then
              oWorksheet.Cells(lngLastRow, 1) = Choose(vEntity + 1, "Task", "Resource")
              oWorksheet.Cells(lngLastRow, 2) = aTypes.getKey(lngType)
            End If
            lngLCF = FieldNameToFieldConstant(aTypes.getKey(lngType) & lngField, vEntity)
            If lngProject = 0 Then
              oWorksheet.Cells(lngLastRow, 3) = lngLCF
              oWorksheet.Cells(lngLastRow, 4) = FieldConstantToFieldName(lngLCF)
            End If
            oExcel.ActiveWindow.ScrollRow = lngLastRow
            oWorksheet.Cells(lngLastRow, 5 + lngProject) = CustomFieldGetName(lngLCF)
            oWorksheet.Cells.Columns.AutoFit
          Next lngField
        Next lngType
      Next vEntity
    Next lngProject
    'add a formula
    oWorksheet.Cells(1, 5 + lngProject) = "MATCH"
    oWorksheet.Range(oWorksheet.Cells(2, 5 + lngProject), oWorksheet.Cells(lngLastRow, 5 + lngProject)).FormulaR1C1 = "=AND(EXACT(RC[-5],RC[-4]),EXACT(RC[-4],RC[-3]),EXACT(RC[-3],RC[-2]),EXACT(RC[-2],RC[-1]))"
    Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), , xlYes)
    oListObject.TableStyle = ""
    oListObject.HeaderRowRange.Font.Bold = True
    'throw some shade
    With oListObject.HeaderRowRange.Interior
      .Pattern = xlSolid
      .PatternColorIndex = xlAutomatic
      .ThemeColor = xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
    End With

    oListObject.Range.Borders(xlDiagonalDown).LineStyle = xlNone
    oListObject.Range.Borders(xlDiagonalUp).LineStyle = xlNone
    For Each vLine In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
      With oListObject.Range.Borders(vLine)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
      End With
    Next vLine
    oExcel.Calculation = xlCalculationAutomatic
    'add conditional formatting
    Set oRange = oListObject.ListColumns("MATCH").DataBodyRange
    oRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE"
    oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
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
    oRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TRUE"
    oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority
    With oRange.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With oRange.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    oRange.FormatConditions(1).StopIfTrue = False
    'autofilter it
    oListObject.Range.AutoFilter oRange.Column, False
    oWorksheet.Columns.AutoFit
    oExcel.ActiveWindow.ScrollRow = 1
    oMasterProject.Activate
    'todo: keep them open for changes, close on form close?
    For lngProject = 0 To aProjects.Count - 1
      If aProjects(lngProject) <> oMasterProject.Name Then
        Projects(aProjects(lngProject)).Activate
        Application.FileCloseEx pjDoNotSave
      End If
    Next lngProject
    cptSpeed False
    oExcel.ScreenUpdating = True
    oExcel.Visible = True
    oExcel.WindowState = xlMaximized
    lngMismatchCount = oRange.SpecialCells(xlCellTypeVisible).Count
    If lngMismatchCount > 0 Then
      oExcel.ActivateMicrosoftApp xlMicrosoftProject
      MsgBox lngMismatchCount & " Local Custom Fields do not match between Master and all Subprojects!", vbCritical + vbOKOnly, "Warning"
      Application.ActivateMicrosoftApp pjMicrosoftExcel
      GoTo exit_here
    Else
      oWorkbook.Close False
      oExcel.Quit
    End If
    
  End If
  
skip_it:

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
  
  'apply the ECF to LCF view
  cptUpdateSaveLocalView
  
  'prepare to capture all ECFs
  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "URL", adVarChar, 255
  rst.Fields.Append "GUID", adGUID
  rst.Fields.Append "pjType", adInteger
  rst.Fields.Append "ENTITY", adVarChar, 50
  rst.Fields.Append "ECF", adInteger
  rst.Fields.Append "ECF_Name", adVarChar, 120
  rst.Fields.Append "LCF", adInteger
  rst.Fields.Append "LCF_Name", adVarChar, 120
  rst.Open
  
  'create a dummy task to interrogate the ECFs
  Set oTask = ActiveProject.Tasks.Add("<dummy for cpt-save-local>")
  Application.CalculateProject
  
  'populate field types
  With cptSaveLocal_frm
    .cboLCF.Clear
    .cboECF.Clear
    .cboECF.AddItem "All Types"
    For lngType = 0 To aTypes.Count - 1
      If aTypes.getKey(lngType) = "Start" Or aTypes.getKey(lngType) = "Finish" Then GoTo next_type
      .cboLCF.AddItem
      .cboLCF.List(.cboLCF.ListCount - 1, 0) = aTypes.getKey(lngType)
      .cboLCF.List(.cboLCF.ListCount - 1, 1) = aTypes.getByIndex(lngType)
      .cboECF.AddItem
      .cboECF.List(.cboECF.ListCount - 1, 0) = aTypes.getKey(lngType)
      .cboECF.List(.cboECF.ListCount - 1, 1) = aTypes.getByIndex(lngType)
next_type:
    Next lngType
    .cboECF.AddItem "undetermined"
    .cmdAutoMap.Enabled = False
    .tglAutoMap = False
    .txtAutoMap.Visible = False
    .chkAutoSwitch = True
    .cboLCF.Value = "Text"
    .optTasks = True
    
    .Show False
  
    'get enterprise custom task fields
    For lngField = 188776000 To 188778000 '2000 should do it for now
      If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strECF = FieldConstantToFieldName(lngField)
        strEntity = cptInterrogateECF(oTask, lngField)
        rst.AddNew Array("URL", "GUID", "pjType", "ENTITY", "ECF", "ECF_Name"), Array(strURL, strGUID, pjTask, strEntity, lngField, FieldConstantToFieldName(lngField))
        lngECFCount = lngECFCount + 1
      End If
      .lblStatus.Caption = "Analyzing Task ECFs...(" & Format(((lngField - 188776000) / (188778000 - 188776000)), "0%") & ")"
      .lblProgress.Width = ((lngField - 188776000) / (188778000 - 188776000)) * .lblStatus.Width
      DoEvents
    Next lngField

    'get enterprise custom resource fields
    For lngField = 205553664 To 205555664 '2000 should do it for now
      If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strECF = FieldConstantToFieldName(lngField)
        strEntity = cptInterrogateECF(oTask, lngField)
        rst.AddNew Array("URL", "GUID", "pjType", "ENTITY", "ECF", "ECF_Name"), Array(strURL, strGUID, pjResource, strEntity, lngField, FieldConstantToFieldName(lngField))
        lngECFCount = lngECFCount + 1
      End If
      .lblStatus.Caption = "Analyzing Resource ECFs...(" & Format((lngField - 205553664) / (205555664 - 205553664), "0%") & ")"
      .lblProgress.Width = ((lngField - 205553664) / (205555664 - 205553664)) * .lblStatus.Width
      DoEvents
    Next lngField
  
    oTask.Delete
  
    If Dir(cptDir & "\settings\cpt-ecf.adtg") <> vbNullString Then
      Kill cptDir & "\settings\cpt-ecf.adtg"
    End If
    rst.Sort = "ECF_Name"
    rst.Save cptDir & "\settings\cpt-ecf.adtg"
    rst.Close
  
    'trigger lboECF refresh
    .cboECF.Value = "All Types"
    
    'update the table
    For lngField = 0 To .lboECF.ListCount - 1
      If Not IsNull(.lboECF.List(lngField, 3)) Then
        lngECF = .lboECF.List(lngField, 0)
        lngLCF = .lboECF.List(lngField, 3)
        cptUpdateSaveLocalView lngECF, lngLCF
      End If
    Next lngField
    
    If BLN_TRAP_ERRORS Then
      .Hide
      cptSpeed False
      .Show 'modal to control changes to custom fields
    Else
      cptSpeed False
      .Show False
    End If
  End With

exit_here:
  On Error Resume Next
  Set oRange = Nothing
  Set oListObject = Nothing
  Set aProjects = Nothing
  Set oSubproject = Nothing
  Set oMasterProject = Nothing
  Set aProjects = Nothing
  Application.StatusBar = ""
  Set oWorksheet = Nothing
  oExcel.Calculation = xlCalculationAutomatic
  oExcel.ScreenUpdating = True
  oExcel.Visible = True
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  cptSpeed False
  oTask.Delete
  Set oTask = Nothing
  Set rstSavedMap = Nothing
  Set vType = Nothing
  aTypes.Clear
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
  
  'todo: ensure there is some mapping?
  
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
    rstSavedMap.Fields.Append "URL", adVarChar, 255
    rstSavedMap.Fields.Append "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adBigInt
    rstSavedMap.Fields.Append "LCF", adBigInt
    rstSavedMap.Open
  Else
    'replace existing saved map
    rstSavedMap.Filter = "GUID<>'" & strGUID & "'"
    rstSavedMap.Open strSavedMap
    rstSavedMap.Save strSavedMap, adPersistADTG
    'rstSavedMap.Open strSavedMap
  End If
  
  'get total task count
  ActiveWindow.TopPane.Activate
  'todo: task vs resource
  If ActiveProject.CurrentView <> ".cptSaveLocal Task View" Then
    ViewApply ".cpt_SaveLocal Task View"
  End If
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
'          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & lngECF
'          If rstSavedMap.RecordCount = 1 Then
            'overwrite it
'            rstSavedMap.Delete adAffectCurrent
'          End If
'          rstSavedMap.Filter = ""
'          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND LCF=" & lngLCF
'          If rstSavedMap.RecordCount = 1 Then
            'overwrite it
'            rstSavedMap.Delete adAffectCurrent
'          End If
'          rstSavedMap.Filter = ""
          'add the new record
'          rstSavedMap.AddNew Array(0, 1, 2), Array(strGUID, lngECF, lngLCF)
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

'  rstSavedMap.Save strSavedMap, adPersistADTG

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
    cptSaveLocal_frm.cmdAutoMap.Enabled = False
    If InStr(strMsg, "  X ") > 0 Then
      strMsg = strMsg & "AutoMap is NOT available." & vbCrLf
      strMsg = strMsg & "Free up some fields and try again."
    Else
      strMsg = strMsg & "AutoMap IS available." & vbCrLf
      strMsg = strMsg & "Click GO! to AutoMap now."
      If cptSaveLocal_frm.tglAutoMap Then
        cptSaveLocal_frm.cmdAutoMap.Enabled = True
      Else
        cptSaveLocal_frm.cmdAutoMap.Enabled = False
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
        If .cboLCF <> .lboECF.List(lngECFs, 2) Then .cboLCF = .lboECF.List(lngECFs, 2)
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
  Dim strURL As String
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

  'todo: deletecustomfield if overwriting

  With cptSaveLocal_frm
    'if already mapped then prompt with ECF name and ask to remap
    For lngItem = 0 To .lboECF.ListCount - 1
      If .lboECF.List(lngItem, 3) = lngLCF Then
        If .lboECF.List(lngItem, 0) = lngECF Then GoTo exit_here 'already mapped
        If MsgBox(FieldConstantToFieldName(lngLCF) & " is already mapped to " & .lboECF.List(lngItem, 1) & " - reassign it?", vbExclamation + vbYesNo, "Already Mapped") = vbYes Then '
          .lboECF.List(lngItem, 3) = ""
          .lboECF.List(lngItem, 4) = ""
        Else
          GoTo exit_here
        End If
      End If
next_ecf:
    Next lngItem
    
    CustomFieldDelete lngLCF
    
    'capture rename
    If CustomFieldGetName(lngECF) = Replace(CustomFieldGetName(lngLCF), " (" & FieldConstantToFieldName(lngLCF) & ")", "") Then GoTo skip_rename
    If Len(CustomFieldGetName(lngLCF)) > 0 Then
      'todo: skip if name already matches
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
    
skip_rename:
    
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
  
  strURL = ActiveProject.ServerURL
  
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
      rstSavedMap.AddNew Array(0, 1, 2, 3), Array(strURL, strGUID, lngECF, lngLCF)
    End If
    rstSavedMap.Filter = ""
    rstSavedMap.Save strSavedMap, adPersistADTG
  Else 'create it
    rstSavedMap.Fields.Append "URL", adVarChar, 255
    rstSavedMap.Fields.Append "GUID", adGUID
    rstSavedMap.Fields.Append "ECF", adInteger
    rstSavedMap.Fields.Append "LCF", adInteger
    rstSavedMap.Open
    rstSavedMap.AddNew Array(0, 1, 2, 3), Array(strURL, strGUID, lngECF, lngLCF)
    rstSavedMap.Save strSavedMap, adPersistADTG
  End If
  rstSavedMap.Close
  
  'update the table
  cptUpdateSaveLocalView lngECF, lngLCF
  
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

Sub cptExportCFMap()
  'objects
  Dim rstSavedMap As ADODB.Recordset
  'strings
  Dim strMsg As String
  Dim strSavedMapExport As String
  Dim strGUID As String
  Dim strSavedMap As String
  'longs
  Dim lngFile As Long
  Dim lngProjectCount As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  strMsg = "Your maps are only valid for other users on this server:" & vbCrLf & vbCrLf
  strMsg = strMsg & ActiveProject.ServerURL & vbCrLf & vbCrLf
  strMsg = strMsg & "Proceed with export?"
  If MsgBox(strMsg, vbExclamation + vbYesNo, "Important") = vbNo Then GoTo exit_here
  
  'ensure file exists
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  If Dir(strSavedMap) = vbNullString Then
    MsgBox "You have no saved map for this project.", vbExclamation + vbOKOnly, "No Map"
    GoTo exit_here
  End If
  
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'prepare an export csv file
  lngFile = FreeFile
  strSavedMapExport = Environ("USERPROFILE") & "\Downloads\"
  If Dir(strSavedMapExport, vbDirectory) = vbNullString Then
    strSavedMapExport = Environ("USERPROFILE")
  End If
  strSavedMapExport = strSavedMapExport & "cpt-saved-map.csv"
  If Dir(strSavedMapExport) <> vbNullString Then Kill strSavedMapExport
  Open strSavedMapExport For Output As #lngFile
  'open the filtered recordset and export it
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  With rstSavedMap
    .Open strSavedMap, "Provider=MSPersist", , , adCmdFile
    .Filter = "GUID='" & UCase(strGUID) & "'"
    If .RecordCount = 0 Then
      MsgBox "You have no saved map for this project.", vbExclamation + vbOKOnly, "No Map"
    Else
      Print #lngFile, .GetString(adClipString, , ",", vbCrLf, vbNullString)
    End If
    Close #lngFile
    .Filter = ""
    .Close
  End With
  
  MsgBox "Map saved to '" & strSavedMapExport & "'", vbInformation + vbOKOnly, "Export Complete"
    
exit_here:
  On Error Resume Next
  rstSavedMap.Close
  Set rstSavedMap = Nothing
  Close #lngFile
  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptSaveLocal_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportCFMap()
  'objects
  Dim rstSavedMap As ADODB.Recordset
  Dim oStream As Scripting.TextStream
  Dim oFile As Scripting.File
  Dim oFSO As Scripting.FileSystemObject
  Dim oExcel As Excel.Application
  Dim oFileDialog As FileDialog
  'strings
  Dim strGUID As String
  Dim strConn As String
  Dim strSavedMapImport As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim aLine As Variant
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    
  'get guid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
    
  'borrow Excel's FileDialogFilePicker
  Set oExcel = CreateObject("Excel.Application")
  Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
  With oFileDialog
    .AllowMultiSelect = False
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .InitialFileName = Environ("USERPROFILE") & "\Downloads\"
    .Title = "Select cpt-saved-map.csv:"
    .Filters.Add "Comma Separated Values (csv)", "*.csv"
    If .Show = -1 Then
      strSavedMapImport = .SelectedItems(1)
    End If
  End With
  'close Excel, with thanks...
  oExcel.Quit
  
  'stream the csv
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFile = oFSO.GetFile(strSavedMapImport)
  Set oStream = oFile.OpenAsTextStream(ForReading)
  
  'open user's saved map
  Set rstSavedMap = CreateObject("ADODB.Recordset")
  With rstSavedMap
    If Dir(cptDir & "\settings\cpt-save-local.adtg") = vbNullString Then 'create it
      .Fields.Append "ServerURL", adVarChar, 255
      .Fields.Append "GUID", adGUID
      .Fields.Append "ECF", adInteger
      .Fields.Append "LCF", adInteger
      .Save strSavedMapImport, adPersistADTG
    End If
    .Open cptDir & "\settings\cpt-save-local.adtg", "Provider=MSPersist", , , adCmdFile
    
    Do Until oStream.AtEndOfStream
      aLine = Split(oStream.ReadLine, ",")
      If UBound(aLine) > 0 Then
        'ensure same server
        If aLine(0) <> ActiveProject.ServerURL Then
          MsgBox "Maps can only be imported from this server:" & vbCrLf & vbCrLf & ActiveProject.ServerURL, vbCritical + vbOKOnly, "Invalid Map"
          GoTo exit_here
        Else
          'ensure ecf exists
          If FieldConstantToFieldName(CLng(aLine(2))) = "<Unavailable>" Then
            'todo: strECF
            MsgBox "The imported ECF (" & CLng(aLine(2)) & ") is <Unavailable> on this server.", vbCritical + vbOKOnly, "Invalid ECF"
          Else
            'todo: ensure same ECF settings?
            cptSaveLocal_frm.lboECF.Value = CLng(aLine(2))
            cptSaveLocal_frm.lboLCF.Value = CLng(aLine(3))
            Call cptMapECFtoLCF(CLng(aLine(2)), CLng(aLine(3)))
          End If
        End If
      End If
    Loop
    
  End With

exit_here:
  On Error Resume Next
  Set rstSavedMap = Nothing
  Set oStream = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing
  oExcel.Quit
  Set oExcel = Nothing
  Set oFileDialog = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptImportCFMap", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateECF(Optional strFilter As String)
  'objects
  Dim rstECF As ADODB.Recordset
  Dim rstSavedMap As ADODB.Recordset
  'strings
  Dim strGUID As String
  Dim strSavedMap As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'get project guid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'open the ecf recordset
  Set rstECF = CreateObject("ADODB.Recordset")
  rstECF.Open cptDir & "\settings\cpt-ecf.adtg"
    
  'check for saved map
  strSavedMap = cptDir & "\settings\cpt-save-local.adtg"
  blnExists = Dir(strSavedMap) <> vbNullString
  If blnExists Then
    Set rstSavedMap = CreateObject("ADODB.Recordset")
    rstSavedMap.Open strSavedMap
  End If
  
  'populate the form - defaults to task ECFs, text
  With cptSaveLocal_frm
    'populate map
    .lboECF.Clear
    If rstECF.RecordCount = 0 Then
      rstECF.Close
      MsgBox "No Enterprise Custom Fields available in this file.", vbExclamation + vbOKOnly, "No ECFs found"
      GoTo exit_here
    End If
    'apply
    If .cboECF <> "All Types" Then
      If .cboECF = "undetermined" Then
        rstECF.Filter = "ENTITY='undetermined' OR ENTITY LIKE '%Maybe%'"
      Else
        rstECF.Filter = "ENTITY='" & .cboECF & "' OR ENTITY='undetermined'"
      End If
    End If
    If rstECF.RecordCount = 0 Then GoTo no_records
    rstECF.MoveFirst
    Do While Not rstECF.EOF
      If UCase(rstECF("GUID")) = UCase(strGUID) And rstECF("pjType") = IIf(.optTasks, pjTask, pjResource) Then
        If strFilter <> "" Then
          If InStr(UCase(rstECF("ECF_Name")), UCase(strFilter)) = 0 Then GoTo next_field
        End If
        .lboECF.AddItem
        .lboECF.List(.lboECF.ListCount - 1, 0) = rstECF("ECF")
        .lboECF.List(.lboECF.ListCount - 1, 1) = rstECF("ECF_Name")
        .lboECF.List(.lboECF.ListCount - 1, 2) = rstECF("ENTITY")
        If blnExists Then
          rstSavedMap.Filter = "GUID='" & UCase(strGUID) & "' AND ECF=" & rstECF("ECF") '& " AND ENTITY=" & IIf(.optTasks, pjTask, pjResource)
          If Not rstSavedMap.EOF Then
            .lboECF.List(.lboECF.ListCount - 1, 3) = rstSavedMap("LCF")
            If Len(CustomFieldGetName(rstSavedMap("LCF"))) > 0 Then
              .lboECF.List(.lboECF.ListCount - 1, 4) = CustomFieldGetName(rstSavedMap("LCF"))
            Else
              .lboECF.List(.lboECF.ListCount - 1, 4) = FieldConstantToFieldName(rstSavedMap("LCF"))
            End If
          End If
          rstSavedMap.Filter = ""
        End If
      End If
next_field:
      rstECF.MoveNext
    Loop
no_records:
    .lblStatus.Caption = Format(.lboECF.ListCount, "#,##0") & " enterprise custom field(s)."

  End With

exit_here:
  On Error Resume Next
  If rstECF.State Then rstECF.Close
  Set rstECF = Nothing
  If rstSavedMap.State Then rstSavedMap.Close
  Set rstSavedMap = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptUpdateECF", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateLCF(Optional strFilter As String)
'objects
'strings
Dim strCustomName As String
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

  With cptSaveLocal_frm
    .lboLCF.Clear
    lngFields = .cboLCF.Column(1)
    For lngField = 1 To lngFields
      strFieldName = .cboLCF.Column(0) & lngField
      lngFieldID = FieldNameToFieldConstant(strFieldName, IIf(.optTasks, pjTask, pjResource))
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant(.cboLCF.Column(0) & lngField, IIf(.optTasks, pjTask, pjResource)))
      If strFilter <> "" Then
        If InStr(UCase(strCustomName), UCase(strFilter)) = 0 Or Len(strCustomName) = 0 Then GoTo next_field
      End If
      If Len(strCustomName) > 0 Then
        .lboLCF.AddItem
        .lboLCF.List(.lboLCF.ListCount - 1, 0) = lngFieldID
        .lboLCF.List(.lboLCF.ListCount - 1, 1) = strFieldName & " (" & CustomFieldGetName(lngFieldID) & ")" 'Me.lboLCF.List(Me.lboLCF.ListCount - 1, 0) = CustomFieldGetName(FieldNameToFieldConstant(Me.cboLCF.Column(0) & lngField))
      Else
        .lboLCF.AddItem
        .lboLCF.List(.lboLCF.ListCount - 1, 0) = lngFieldID
        .lboLCF.List(.lboLCF.ListCount - 1, 1) = strFieldName
      End If
next_field:
    Next lngField
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cboLCF_Change", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateSaveLocalView(Optional lngECF As Long, Optional lngLCF As Long)
  'objects
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  With cptSaveLocal_frm
    If Not .Visible Then
      'create/overwrite the tables
      cptSpeed True
      ViewApply "Gantt Chart"
      TableEditEx ".cptSaveLocal Task Table", True, True, True, , "ID", , , , , , True, , , , , , , , False
      TableEditEx ".cptSaveLocal Task Table", True, False, , , , "Unique ID", "UID", , , , True
      TableEditEx ".cptSaveLocal Resource Table", False, True, True, , "ID", , , , , , True, , , , , , , , False
      TableEditEx ".cptSaveLocal Resource Table", False, False, , , , "Unique ID", "UID", , , , True
      On Error Resume Next
      ActiveProject.Views(".cptSaveLocal Task View").Delete
      ActiveProject.Views(".cptSaveLocal Resource View").Delete
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      ViewEditSingle ".cptSaveLocal Task View", True, , pjTaskSheet, , , ".cptSaveLocal Task Table", "All Tasks", "No Group"
      ViewEditSingle ".cptSaveLocal Resource View", True, , pjResourceSheet, , , ".cptSaveLocal Resource Table", "All Resources", "No Group"
      'update the table
      For lngField = 0 To .lboECF.ListCount - 1
        If Not IsNull(.lboECF.List(lngField, 3)) Then
          lngECF = .lboECF.List(lngField, 0)
          lngLCF = .lboECF.List(lngField, 3)
          TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False
          TableEditEx ".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False
        End If
      Next lngField
      ViewApply ".cptSaveLocal Task View"
      cptSpeed False
    Else
      If .optTasks Then
        If ActiveProject.CurrentView <> ".cptSaveLocal Task View" Then ViewApply ".cptSaveLocal Task View"
        If lngECF > 0 And lngLCF > 0 Then
          If Not TableEditEx(".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          If Not TableEditEx(".cptSaveLocal Task Table", True, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          TableApply ".cptSaveLocal Task Table"
        End If
      ElseIf .optResources Then
        If ActiveProject.CurrentView <> ".cptSaveLocal Resource View" Then ViewApply ".cptSaveLocal Resource View"
        If lngECF > 0 And lngLCF > 0 Then
          If Not TableEditEx(".cptSaveLocal Resource Table", False, False, False, , , FieldConstantToFieldName(lngECF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          If Not TableEditEx(".cptSaveLocal Resource Table", False, False, False, , , FieldConstantToFieldName(lngLCF), , , , , True, , , , , , , , False) Then
            MsgBox "Failed to add column " & FieldConstantToFieldName(lngECF) & "!", vbExclamation + vbOKOnly, "Fail"
          End If
          TableApply ".cptSaveLocal Resource Table"
        End If
      End If
    End If
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptSaveLocal_bas", "cptUpdateSaveLocalView", Err, Erl)
  Resume exit_here
End Sub
