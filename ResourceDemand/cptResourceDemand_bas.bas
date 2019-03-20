Attribute VB_Name = "cptResourceDemand_bas"
'<cpt_version>v1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Const adVarChar As Long = 200

Sub ExportResourceDemand(Optional lngTaskCount As Long)
'objects
Dim Task As Task, Resource As Resource, Assignment As Assignment
Dim TSVS As TimeScaleValues, TSV As TimeScaleValue
Dim TSVS_WORK As TimeScaleValues, TSVS_ACTUAL As TimeScaleValues
Dim xlApp As Excel.Application, Worksheet As Worksheet, Workbook As Workbook
Dim rng As Excel.Range
Dim PivotTable As PivotTable, ListObject As ListObject, PivotChart As ChartObject
Dim rst As Object 'ADODB.Recordset
'dates
Dim dtStart As Date, dtMin As Date, dtMax As Date
'doubles
Dim dblWork As Double
'strings
Dim strView As String
Dim strFile As String, strGroup As String, strRange As String, strCLIN As String
Dim strTitle As String, strMsg As String, strSQL As String, strHeaders As String
Dim strRecord As String, strFileName As String
'longs
Dim lgFile As Long, lgTasks As Long, lgTask As Long
Dim lgCol As Long, lgExport As Long, lgField As Long
'variants
Dim aUserFields() As Variant

  SpeedON
  
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
            strRecord = Task.Project & ",[" & Task.UniqueID & "] " & Replace(Task.Name, ",", "") & "," & Assignment.ResourceName & "," & dblWork / 60 & "," & DateAdd("d", 1, TSV.startDate)
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
  Set xlApp = CreateObject("Excel.Application")
  'create a new workbook
  Set Workbook = xlApp.Workbooks.Open(strFile)
  
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then Kill Environ("TEMP") & "\ExportResourceDemand.xlsx"
  Workbook.SaveAs Environ("TEMP") & "\ExportResourceDemand.xlsx", 51
  
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
  dtStart = ActiveProject.StatusDate
  strFile = Replace(strFile, ".csv", ".xlsx")
  strFile = Replace(strFile, ".xlsx", "_" & Format(dtStart, "yyyy-mm-dd") & ".xlsx")
file_save:
  If Dir(strFile) <> vbNullString Then Kill strFile
  Workbook.SaveAs strFile, 51
  
  Application.StatusBar = "Complete."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  MsgBox "Saved to your Desktop:" & vbCrLf & vbCrLf & Dir(strFile), vbInformation + vbOKOnly, "Resource Demand Exported"
  xlApp.Visible = True
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  cptResourceDemand_frm.lblStatus.Caption = "Ready..."
  If FreeFile > 0 Then Close #lgFile
  SpeedOFF
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
  Call HandleErr("basBrashear", "ExportResourceDemand", err)
  On Error Resume Next
  Resume exit_here
      
End Sub

Sub ShowFrmExportResourceDemand()
'longs
Dim lngResourceCount As Long
Dim lgFieldType As Variant, lgField As Long, lngItem As Long
'integers
Dim intField As Long
'strings
Dim strActiveView As String
Dim strFieldName As String, strFileName As String
'objects
Dim objProject As Project
Dim arrFields As Object
'variants
Dim strFieldType As Variant, st As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'requires ms excel
  If Not CheckReference("Excel") Then
    MsgBox "This feature requires MS Excel.", vbCritical + vbOKOnly, "Resource Demand"
    GoTo exit_here
  End If
  If ActiveProject.Subprojects.count = 0 And ActiveProject.ResourceCount = 0 Then
    MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
    GoTo exit_here
  Else
    SpeedON
    strActiveView = ActiveWindow.TopPane.View.Name
    ViewApply "Resource Sheet"
    SelectAll
    lngResourceCount = ActiveSelection.Resources.count
    ViewApply strActiveView
    SpeedOFF
    If lngResourceCount = 0 Then
      MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
      GoTo exit_here
    End If
  End If
  
  cptResourceDemand_frm.lboFields.Clear
  cptResourceDemand_frm.lboExport.Clear

  Set arrFields = CreateObject("System.Collections.ArrayList")

  'For Each lgFieldType In Array(0) '0 = pjTask; 1 = pjResource; 2 = pjProject
    'For Each strFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Number", "Start", "Text", "Outline Code")
    For Each strFieldType In Array("Text", "Outline Code")
      On Error GoTo err_here
      For intField = 1 To 30
        lgField = FieldNameToFieldConstant(strFieldType & intField) ',lgFieldType)
        strFieldName = CustomFieldGetName(lgField)
        If Len(strFieldName) > 0 Then arrFields.Add strFieldName
next_field:
      Next intField
    Next strFieldType

  'get enterprise custom fields
  For lgField = 188776000 To 188778000
    If Application.FieldConstantToFieldName(lgField) <> "<Unavailable>" Then
      arrFields.Add Application.FieldConstantToFieldName(lgField)
    End If
  Next lgField
  
  'add fields to listbox
  arrFields.Sort
  st = arrFields.GetRange(0, arrFields.count).ToArray
  For intField = 0 To UBound(st)
    cptResourceDemand_frm.lboFields.AddItem
    cptResourceDemand_frm.lboFields.List(lngItem, 0) = FieldNameToFieldConstant(st(intField))
    If FieldNameToFieldConstant(st(intField)) >= 188776000 Then
      cptResourceDemand_frm.lboFields.List(lngItem, 1) = st(intField) & " (Enterprise)"
    Else
      cptResourceDemand_frm.lboFields.List(lngItem, 1) = st(intField) & " (" & FieldConstantToFieldName(FieldNameToFieldConstant(st(intField))) & ")"
    End If
    lngItem = lngItem + 1
  Next intField
  
  'save the fields to a file
  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  With CreateObject("ADODB.Recordset")
    .Fields.Append "Field Constant", adVarChar, 100
    .Fields.Append "Custom Field Name", adVarChar, 100
    .Open
    For lngItem = 0 To cptResourceDemand_frm.lboFields.ListCount - 1
      .AddNew Array(0, 1), Array(cptResourceDemand_frm.lboFields.List(lngItem, 0), cptResourceDemand_frm.lboFields.List(lngItem, 1))
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
  Set objProject = Nothing
  arrFields.Clear
  Set arrFields = Nothing
  Erase st
  Set st = Nothing
  Exit Sub

err_here:
  If err.Number = 1101 Or err.Number = 1004 Then
    err.Clear
    Resume next_field
  Else
    Call HandleErr("cptResourceDemand_bas", "ShowcptResourceDemand_frm", err)
    Resume exit_here
  End If
  
End Sub
