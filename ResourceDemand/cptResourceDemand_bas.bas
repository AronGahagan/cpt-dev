Attribute VB_Name = "cptResourceDemand_bas"
'<cpt_version>v1.1.3</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Const adVarChar As Long = 200

Sub cptExportResourceDemand(Optional lngTaskCount As Long)
'objects
Dim Task As Task, Resource As Resource, Assignment As Assignment
Dim TSV As TimeScaleValue
Dim TSVS_WORK As TimeScaleValues, TSVS_ACTUAL As TimeScaleValues
Dim xlApp As Excel.Application, Worksheet As Worksheet, Workbook As Workbook
Dim rng As Excel.Range
Dim PivotTable As PivotTable, ListObject As ListObject
'dates
Dim dtStart As Date, dtMin As Date, dtMax As Date
'doubles
Dim dblWork As Double
'strings
Dim strMsg As String
Dim strView As String
Dim strFile As String, strRange As String
Dim strTitle As String, strHeaders As String
Dim strRecord As String, strFileName As String
'longs
Dim lgFile As Long, lgTasks As Long, lgTask As Long
Dim lgCol As Long, lgExport As Long, lgField As Long
'variants
Dim aUserFields() As Variant
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
    If ActiveProject.ProjectStart > dtStart Then dtStart = ActiveProject.ProjectStart
  Else
    MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    GoTo exit_here
  End If
  
  'save settings
  strFileName = Environ("tmp") & "\cpt-export-resource-userfields.adtg."
  aUserFields = cptResourceDemand_frm.lboExport.List()
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", adVarChar, 255
    .Fields.Append "Custom Field Name", adVarChar, 255
    .Open
    For lgExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
      .AddNew Array(0, 1), Array(aUserFields(lgExport, 0), aUserFields(lgExport, 1))
    Next lgExport
    .Update
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    .Save strFileName
    .Close
  End With
  
  lgFile = FreeFile
  strFile = Environ("USERPROFILE") & "\Desktop\" & Replace(Replace(ActiveProject.Name, ".mpp", ""), " ", "_") & "_ResourceDemand.csv"
  
  If Dir(strFile) <> vbNullString Then Kill strFile
  
  Open strFile For Output As #lgFile
  strHeaders = "PROJECT,[UID] TASK,RESOURCE_NAME,HOURS,WEEK"
  For lgExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
    lgField = cptResourceDemand_frm.lboExport.List(lgExport, 0)
    strHeaders = strHeaders & "," & CustomFieldGetName(lgField)
  Next lgExport
  Print #lgFile, strHeaders
  
  If ActiveProject.Subprojects.count = 0 Then
    lgTasks = ActiveProject.Tasks.count
  Else
    cptSpeed True
    strView = ActiveWindow.TopPane.View.Name
    ViewApply "Gantt Chart"
    FilterClear
    GroupClear
    SelectAll
    OptionsViewEx displaysummarytasks:=True
    OutlineShowAllTasks
    SelectAll
    lgTasks = ActiveSelection.Tasks.count
    ViewApply strView
    cptSpeed False
  End If
  
  'iterate over tasks
  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then 'skip blank lines
    If Task.ExternalTask Then GoTo next_task 'skip external tasks
    If Not Task.Summary And Task.RemainingDuration > 0 And Task.Active Then 'skip summary, complete tasks/milestones, and inactive
    If Task.Start > ActiveProject.StatusDate Then dtStart = Task.Start Else dtStart = ActiveProject.StatusDate
      'examine every assignment on the task
      For Each Assignment In Task.Assignments
        'limit export to labor resources only
        If Assignment.ResourceType = pjResourceTypeWork Then
          'capture timephased work (ETC)
          Set TSVS_WORK = Assignment.TimeScaleData(DateAdd("d", -7, dtStart), Task.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
          For Each TSV In TSVS_WORK
            'capture (and subtract) actual work, leaving ETC/Remaining Work
            Set TSVS_ACTUAL = Assignment.TimeScaleData(TSV.startDate, TSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
            dblWork = Val(TSV.Value) - Val(TSVS_ACTUAL(1))
            'write a record to the CSV
            '<issue14-15>strRecord = Task.Project & ",[" & Task.UniqueID & "] " & Replace(Task.Name, ",", "") & "," & Assignment.ResourceName & "," & dblWork / 60 & "," & DateAdd("d", 1, TSV.startDate) - removed </issue14-15>
            '<issue14-15> added
            strRecord = Task.Project & "," & Chr(34) & "[" & Task.UniqueID & "] " & Replace(Task.Name, Chr(34), Chr(39)) & Chr(34) & "," & Assignment.ResourceName & "," & dblWork / 60 & "," & DateAdd("d", 1, TSV.startDate)
            '</issue14-15>
            For lgExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
              lgField = cptResourceDemand_frm.lboExport.List(lgExport, 0)
              strRecord = strRecord & "," & Task.GetField(lgField)
            Next lgExport
            Print #lgFile, strRecord
          Next TSV
        End If
next_assignment:
        Next Assignment
      End If
    End If
next_task:
    lgTask = lgTask + 1
    Application.StatusBar = "Exporting " & Format(lgTask, "#,##0") & " of " & Format(lgTasks, "#,##0") & "...(" & Format(lgTask / lgTasks, "0%") & ")"
    cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
    cptResourceDemand_frm.lblProgress.Width = (lgTask / lgTasks) * cptResourceDemand_frm.lblStatus.Width
    DoEvents
  Next Task
  
  'close the CSV
  Close #lgFile
   
  Application.StatusBar = "Creating Workbook..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  'set reference to Excel
  '<issue14-15> added
  On Error Resume Next
  Set xlApp = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
  End If
  
  'is previous run still open>
  On Error Resume Next
  Set Workbook = xlApp.Workbooks(strFile)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not Workbook Is Nothing Then Workbook.Close False
  On Error Resume Next
  Set Workbook = xlApp.Workbooks(Environ("TEMP") & "\ExportResourceDemand.xlsx")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not Workbook Is Nothing Then 'add timestamp to existing file
    If Workbook.Application.Visible = False Then Workbook.Application.Visible = True
    strMsg = "'" & strFile & "' already exists and is open."
    strFile = Replace(strFile, ".xlsx", "_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx")
    strMsg = strMsg & "Your new file will be saved as:" & vbCrLf & strFile
    MsgBox strMsg, vbExclamation + vbOKOnly, "File Exists and is Open"
  End If '</issue14-15>
  
  'create a new workbook
  Set Workbook = xlApp.Workbooks.Open(strFile)
  
  '<issue14-15> added
  On Error Resume Next
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then Kill Environ("TEMP") & "\ExportResourceDemand.xlsx"
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then 'kill failed, rename it
    Workbook.SaveAs Environ("TEMP") & "\ExportResourceDemand_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx", 51
  Else
    Workbook.SaveAs Environ("TEMP") & "\ExportResourceDemand.xlsx", 51
  End If
  If Dir(strFile) <> vbNullString Then Kill strFile '</issue14-15>
  
  'set reference to worksheet to manipulate it
  Set Worksheet = Workbook.Sheets(1)
  'rename the worksheet
  Worksheet.Name = "SourceData"
  lgCol = Worksheet.Rows(1).Find(what:="WEEK").Column
  dtMin = xlApp.WorksheetFunction.Min(Worksheet.Columns(lgCol))
  dtMax = xlApp.WorksheetFunction.Max(Worksheet.Columns(lgCol))
  
  'import fiscal weeks
'  Set Worksheet = Workbook.Sheets.Add(After:=Workbook.Sheets("SourceData"))
'  Worksheet.Name = "Fiscal Weeks"
'  strSQL = "SELECT WEEK_ENDING,LEFT(MONTH_LABEL,3),2000+cast(RIGHT(MONTH_LABEL,2) as int) FROM FISCAL_WEEKS "
'  strSQL = strSQL & "WHERE WEEK_ENDING>='" & Format(dtMin, "yyyy-mm-dd") & "' "
'  strSQL = strSQL & "AND WEEK_ENDING<='" & Format(DateAdd("d", 7, dtMax), "yyyy-mm-dd") & "' "
'  strSQL = strSQL & "ORDER BY WEEK_ENDING"
'  Set rst = New ADODB.Recordset
'  rst.Open strSQL, STR_CON, adOpenKeyset
'  Worksheet.[A1].Value = "WEEK_ENDING"
'  Worksheet.[B1].Value = "MONTH_LABEL"
'  Worksheet.[C1].Value = "YEAR_LABEL"
'  Worksheet.[A2].CopyFromRecordset rst
'  Worksheet.Columns(1).Replace "-", "/"
'  xlApp.ActiveWindow.Zoom = 85
'  Worksheet.Cells.EntireColumn.AutoFit
'  Set rng = Worksheet.Range(Worksheet.[A1].End(xlDown), Worksheet.[A1].End(xlToRight))
'  Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
'  ListObject.Name = "FISCAL_WEEKS"
'  ListObject.TableStyle = ""
  
  Set Worksheet = Workbook.Sheets("SourceData")
  'Worksheet.Activate
'  'add fiscal year calc
'  lgCol = Worksheet.Rows(1).Find(what:="FISCAL_YEAR").Column
'  Set rng = Worksheet.Range(Worksheet.Cells(2, lgCol), Worksheet.Cells(2, lgCol - 1).End(xlDown).Offset(0, 1))
'  rng.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1]+4,FISCAL_WEEKS,3,FALSE),""<n/a>"")"
'  rng.Copy
'  rng.PasteSpecial xlPasteValues
'  'add fiscal month labels
'  lgCol = Worksheet.Rows(1).Find(what:="FISCAL_MONTH").Column
'  Set rng = Worksheet.Range(Worksheet.Cells(2, lgCol), Worksheet.Cells(2, lgCol - 1).End(xlDown).Offset(0, 1))
'  rng.FormulaR1C1 = "=IFERROR(LEFT(VLOOKUP(RC[-2]+4,FISCAL_WEEKS,2,FALSE),3),""<n/a>"")"
'  rng.Copy
'  rng.PasteSpecial xlPasteValues
  
  'capture the range of data to feed as variable to PivotTable
  Set rng = Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown))
  strRange = Worksheet.Name & "!" & Replace(rng.Address, "$", "")
  'add a new worksheet for the PivotTable
  Set Worksheet = Workbook.Sheets.Add
  'rename the new worksheet
  Worksheet.Name = "ResourceDemand"
  
  Application.StatusBar = "Creating PivotTable..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  'create the PivotTable
  Workbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        strRange, Version:= _
        xlPivotTableVersion12).CreatePivotTable TableDestination:="ResourceDemand!R3C1", _
        TableName:="RESOURCE_DEMAND", DefaultVersion:=xlPivotTableVersion12
  Set PivotTable = Worksheet.PivotTables(1)
  PivotTable.AddFields Array("RESOURCE_NAME", "PROJECT", "[UID] TASK"), Array("WEEK") 'Array("FISCAL_YEAR", "FISCAL_MONTH", "WEEK")
  PivotTable.AddDataField PivotTable.PivotFields("HOURS"), "HOURS ", -4157
  'format the PivotTable
  PivotTable.PivotFields("RESOURCE_NAME").ShowDetail = False
  'PivotTable.PivotFields("FISCAL_MONTH").ShowDetail = False
  'PivotTable.PivotFields("FISCAL_MONTH").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
  'PivotTable.PivotFields("FISCAL_YEAR").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
  PivotTable.TableStyle2 = "PivotStyleMedium2"
  PivotTable.PivotSelect "", xlDataOnly, True
  xlApp.Selection.Style = "Comma"
  
  Application.StatusBar = "Building header..."
  cptResourceDemand_frm.lblStatus = Application.StatusBar
  
  'add a title
  Worksheet.[A2] = "Status Date: " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  Worksheet.[A2].EntireColumn.AutoFit
  Worksheet.[A1] = "REMAINING WORK IN IMS: " & Replace(ActiveProject.Name, " ", "_")
  Worksheet.[A1].Font.Bold = True
  Worksheet.[A1].Font.Italic = True
  Worksheet.[A1].Font.Size = 14
  Worksheet.[A1:F1].Merge
'  If blnReturn Then
'    Worksheet.[A2] = "Status Date: " & FormatDateTime(DateAdd("d", -7, dtStart), vbShortDate)
'  Else
'  End If
  Worksheet.[B2] = "Weeks Beginning" '"Fiscal Years / Fiscal Months / Weeks Beginning"
  Worksheet.[B4].Select
  'group the weeks into years/months and collapse
  'xlApp.Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, True, True, False, True)
  Worksheet.[B5].Select
  'PivotTable.PivotFields("Months").ShowDetail = False
  
  'make it nice
  xlApp.ActiveWindow.Zoom = 85
  
  Application.StatusBar = "Creating PivotChart..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  'create a PivotChart
  Set Worksheet = Workbook.Sheets("SourceData")
  Worksheet.Activate
  Worksheet.[A2].Select
  Worksheet.[A2].EntireColumn.AutoFit
  xlApp.ActiveWindow.Zoom = 85
  xlApp.ActiveWindow.FreezePanes = True
  Worksheet.Cells.EntireColumn.AutoFit
  Set Worksheet = Workbook.Sheets.Add
  Worksheet.Name = "PivotChart_Source"
  Workbook.Worksheets("ResourceDemand").PivotTables("RESOURCE_DEMAND"). _
        PivotCache.CreatePivotTable TableDestination:="PivotChart_Source!R1C1", TableName:= _
        "PivotTable1", DefaultVersion:=xlPivotTableVersion12
  Set Worksheet = Workbook.Sheets("PivotChart_Source")
  Worksheet.[A1].Select
  xlApp.ActiveSheet.Shapes.AddChart.Select
  Set rng = Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown))
  xlApp.ActiveChart.SetSourceData Source:=rng
  Workbook.ShowPivotChartActiveFields = True
  'xlApp.ActiveChart.ChartType = xlColumnClustered
  xlApp.ActiveChart.ChartType = xlAreaStacked
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK")
    .Orientation = xlRowField
    .Position = 1
  End With
  xlApp.ActiveChart.PivotLayout.PivotTable.AddDataField xlApp.ActiveChart.PivotLayout. _
        PivotTable.PivotFields("HOURS"), "Sum of HOURS", xlSum
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("RESOURCE_NAME")
    .Orientation = xlColumnField
    .Position = 1
  End With

'  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("FISCAL_YEAR")
'    .Orientation = xlRowField
'    .Position = 2
'  End With
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK")
    .Orientation = xlRowField
    .Position = 1
  End With
  xlApp.ActiveSheet.PivotTables("PivotTable1").PivotFields("WEEK").PivotFilters.Add _
        Type:=xlAfter, Value1:=ActiveProject.StatusDate
  xlApp.ActiveChart.ClearToMatchStyle
  xlApp.ActiveChart.ChartStyle = 34
  xlApp.ActiveChart.ClearToMatchStyle
  xlApp.ActiveChart.SetElement (msoElementChartTitleAboveChart)
  xlApp.ActiveSheet.ChartObjects(1).Activate
'  If blnReturn Then
'    strTitle = Replace(ActiveProject.Name, ".mpp", "") & " - Resource Demand" & Chr(13) & "As of WE " & FormatDateTime(DateAdd("d", -7, dtStart), vbShortDate)
'  Else
    strTitle = Replace(ActiveProject.Name, ".mpp", "") & " - Resource Demand" & Chr(13) & "As of WE " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
'  End If
  xlApp.ActiveChart.ChartTitle.Text = strTitle
  xlApp.ActiveChart.Location xlLocationAsNewSheet, "PivotChart"
  Set Worksheet = Workbook.Sheets("PivotChart_Source")
  Worksheet.Visible = False
  
  'add legend
  xlApp.ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
  xlApp.ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Hours"
    
  Application.StatusBar = "Saving the Workbook..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  'save the file
  '<issue14-15> dtStart = ActiveProject.StatusDate - removed
  'strFile = Environ("USERPROFILE") & "\Deskop\" & Replace(strFile, ".csv", ".xlsx") - removed
  'strFile = Replace(strFile, ".xlsx", "_" & Format(dtStart, "yyyy-mm-dd") & ".xlsx") - removed
file_save:
  'If Dir(strFile) <> vbNullString Then Kill strFile - removed </issue14-15>
  Workbook.SaveAs Environ("USERPROFILE") & "\Desktop\" & Workbook.Name, 51
  
  Application.StatusBar = "Complete."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  MsgBox "Saved to your Desktop:" & vbCrLf & vbCrLf & Workbook.Name, vbInformation + vbOKOnly, "Resource Demand Exported"
  xlApp.Visible = True
  
exit_here:
  On Error Resume Next
  If Not xlApp Is Nothing Then xlApp.Visible = True
  Application.StatusBar = ""
  cptResourceDemand_frm.lblStatus.Caption = "Ready..."
  For lgFile = 1 To FreeFile
    Close #lgFile
  Next lgFile
  cptSpeed False
  Set Task = Nothing
  Set Resource = Nothing
  Set Assignment = Nothing
  Set xlApp = Nothing
  Set PivotTable = Nothing
  Set ListObject = Nothing
  Set Workbook = Nothing
  Set Worksheet = Nothing
  If Not Workbook Is Nothing Then Workbook.Close False
  If Not xlApp Is Nothing Then xlApp.Quit
  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_bas", "cptExportResourceDemand", err)
  On Error Resume Next
  Resume exit_here
      
End Sub

Sub ShowFrmExportResourceDemand()
'objects
Dim arrResources As Object
Dim objProject As Object
Dim arrFields As Object
'strings
Dim strActiveView As String
Dim strFieldName As String, strFileName As String
'longs
Dim lngResourceCount As Long, lngResource As Long
Dim lngField As Long, lngItem As Long
'integers
'booleans
'variants
Dim vFieldType As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'requires ms excel
  If Not cptCheckReference("Excel") Then
    MsgBox "This feature requires MS Excel.", vbCritical + vbOKOnly, "Resource Demand"
    GoTo exit_here
  End If
  If ActiveProject.Subprojects.count = 0 And ActiveProject.ResourceCount = 0 Then
    MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
    GoTo exit_here
  Else
    cptSpeed True
    GoTo option_2 'delay is better than a flicker
option_1:
    strActiveView = ActiveWindow.TopPane.View.Name
    ViewApply "Resource Sheet"
    SelectAll
    lngResourceCount = ActiveSelection.Resources.count
    ViewApply strActiveView
    cptSpeed False
    If lngResourceCount = 0 Then
      MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
      GoTo exit_here
    End If
option_2:
    'option 2
    lngResourceCount = ActiveProject.ResourceCount
    Set arrResources = CreateObject("System.Collections.SortedList")
    For lngItem = 1 To ActiveProject.Subprojects.count
      Set objProject = ActiveProject.Subprojects(lngItem).SourceProject
      Application.StatusBar = "Loading " & objProject.Name & "..."
      For lngResource = 1 To objProject.Resources.count
        With arrResources
          If Not .contains(objProject.Resources(lngResource).Name) Then
            .Add objProject.Resources(lngResource).Name, objProject.Resources(lngResource).Name
            lngResourceCount = lngResourceCount + 1
          End If
        End With
      Next lngResource
      Set objProject = Nothing
    Next lngItem
    arrResources.Clear
    Application.StatusBar = ""
    cptSpeed False
  End If
  
  cptResourceDemand_frm.lboFields.Clear
  cptResourceDemand_frm.lboExport.Clear

  Set arrFields = CreateObject("System.Collections.SortedList")
  'col0 = custom field name (sortfield)
  'col1 = field constant
  
  For Each vFieldType In Array("Text", "Outline Code")
    On Error GoTo err_here
    For lngItem = 1 To 30
      lngField = FieldNameToFieldConstant(vFieldType & lngItem) ',lngFieldType)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        If Not arrFields.contains(strFieldName) Then arrFields.Add strFieldName, lngField
      End If
next_field:
    Next lngItem
  Next vFieldType

  'get enterprise custom fields
  For lngField = 188776000 To 188778000
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      arrFields.Add Application.FieldConstantToFieldName(lngField), lngField
    End If
  Next lngField
  
  'add fields to listbox
  For lngItem = 0 To arrFields.count - 1
    cptResourceDemand_frm.lboFields.AddItem
    'column 0 = field constant = arrFields col1
    'column 1 = custom field name = arrFields col0
    cptResourceDemand_frm.lboFields.List(lngItem, 0) = arrFields.getValueList()(lngItem) 'FieldNameToFieldConstant(arrFields.getKey(lngItem))
    If FieldNameToFieldConstant(arrFields.getKey(lngItem)) >= 188776000 Then
      cptResourceDemand_frm.lboFields.List(lngItem, 1) = arrFields.getKey(lngItem) & " (Enterprise)"
    Else
      cptResourceDemand_frm.lboFields.List(lngItem, 1) = arrFields.getKey(lngItem) & " (" & FieldConstantToFieldName(arrFields.getValueList()(lngItem)) & ")"
    End If
  Next lngItem
  
  'save the fields to a file
  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", adVarChar, 100
    .Fields.Append "Custom Field Name", adVarChar, 100
    .Open
    For lngItem = 0 To arrFields.count - 1 'cptResourceDemand_frm.lboFields.ListCount - 1
      'col0 = constant = arrFields col1
      'col1 = field name = arrFields col0
      .AddNew Array(0, 1), Array(arrFields.getValueList()(lngItem), arrFields.getKey(lngItem)) 'Array(cptResourceDemand_frm.lboFields.List(lngItem, 0), cptResourceDemand_frm.lboFields.List(lngItem, 1))
    Next lngItem
    .Save strFileName
    .Close
  End With
  
  'import saved fields if exists
  strFileName = Environ("tmp") & "\cpt-export-resource-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    With CreateObject("ADODB.Recordset")
      .Open strFileName
      .MoveFirst
      lngItem = 0
      Do While Not .EOF
        cptResourceDemand_frm.lboExport.AddItem
        cptResourceDemand_frm.lboExport.List(lngItem, 0) = .Fields(0)     'Field Constant
        cptResourceDemand_frm.lboExport.List(lngItem, 1) = .Fields(1)  'Custom Field Name
        lngItem = lngItem + 1
        .MoveNext
      Loop
      .Close
    End With
  End If

  cptResourceDemand_frm.Show False

exit_here:
  On Error Resume Next
  Set arrResources = Nothing
  Set objProject = Nothing
  Set arrFields = Nothing
  Exit Sub

err_here:
  If err.Number = 1101 Or err.Number = 1004 Then
    err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptResourceDemand_bas", "ShowCptResourceDemand_frm", err)
    Resume exit_here
  End If
  
End Sub
