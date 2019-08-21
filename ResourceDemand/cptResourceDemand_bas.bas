Attribute VB_Name = "cptResourceDemand_bas"
'<cpt_version>v1.1.7</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private Const adVarChar As Long = 200

Sub cptExportResourceDemand(Optional lngTaskCount As Long)
'objects
Dim Task As Task, Resource As Resource, Assignment As Assignment
Dim TSV As TimeScaleValue, TSVS_BCWS As TimeScaleValues
Dim TSVS_WORK As TimeScaleValues, TSVS_AW As TimeScaleValues
Dim TSVS_COST As TimeScaleValues, TSVS_AC As TimeScaleValues
Dim CostRateTable As CostRateTable, PayRate As PayRate
Dim xlApp As Excel.Application, Worksheet As Worksheet, Workbook As Workbook
Dim rng As Excel.Range
Dim PivotTable As PivotTable, ListObject As ListObject
'dates
Dim dtStart As Date, dtFinish As Date, dtMin As Date, dtMax As Date
'doubles
Dim dblWork As Double
'strings
Dim strTask As String
Dim strMsg As String
Dim strView As String
Dim strFile As String, strRange As String
Dim strTitle As String, strHeaders As String
Dim strRecord As String, strFileName As String
'longs
Dim lngOffset As Long
Dim lngRateSets As Long
Dim lgCol As Long
Dim lngOriginalRateSet As Long
Dim lgFile As Long, lgTasks As Long, lgTask As Long
Dim lgWeekCol As Long, lgExport As Long, lgField As Long
Dim lngRateSet As Long
'variants
Dim vRateSet As Variant
Dim aUserFields() As Variant
'booleans
Dim blnIncludeCosts As Boolean

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
    '<issue43> capture case when no custom fields are selected
    If cptResourceDemand_frm.lboExport.ListCount > 0 Then
      .Update
      If Dir(strFileName) <> vbNullString Then Kill strFileName
      .Save strFileName
    End If '</issue43>
    .Close
  End With

  lgFile = FreeFile
  strFile = Environ("USERPROFILE") & "\Desktop\" & Replace(Replace(ActiveProject.Name, ".mpp", ""), " ", "_") & "_ResourceDemand.csv"

  If Dir(strFile) <> vbNullString Then Kill strFile
  
  Open strFile For Output As #lgFile
  strHeaders = "PROJECT,[UID] TASK,RESOURCE_NAME,"
  '<issue42> get selected rate sets
  With cptResourceDemand_frm
    If .chkBaseline Then
      strHeaders = strHeaders & "BL_HOURS,BL_COST,"
    End If
    strHeaders = strHeaders & "HOURS,"
    'get rate sets
    blnIncludeCosts = .chkA Or .chkB Or .chkC Or .chkD Or .chkE
    If blnIncludeCosts Then strHeaders = strHeaders & "RATE_TABLE,COST,"
    If .chkA Then
      strHeaders = strHeaders & "COST_A,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkB Then
      strHeaders = strHeaders & "COST_B,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkC Then
      strHeaders = strHeaders & "COST_C,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkD Then
      strHeaders = strHeaders & "COST_D,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkE Then
      strHeaders = strHeaders & "COST_E,"
      lngRateSets = lngRateSets + 1
    End If
    'get custom fields
    'todo: notify if fields no longer exist
    For lgExport = 0 To .lboExport.ListCount - 1
      lgField = .lboExport.List(lgExport, 0)
      strHeaders = strHeaders & CustomFieldGetName(lgField) & ","
    Next lgExport
    strHeaders = strHeaders & "WEEK"
  End With '</issue42>
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
  'todo: include costs for all rate tables
  'todo: include material and cost resources
  'todo: include baseline, forecast [no actuals]
  Set xlApp = CreateObject("Excel.Application")
  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then 'skip blank lines
      If Task.ExternalTask Then GoTo next_task 'skip external tasks
      If Not Task.Summary And Task.RemainingDuration > 0 And Task.Active Then 'skip summary, complete tasks/milestones, and inactive
        
        'get earliest start and latest finish
        dtStart = xlApp.WorksheetFunction.Min(Task.Start, Task.BaselineStart)
        dtFinish = xlApp.WorksheetFunction.Max(Task.Finish, Task.BaselineFinish)
        
        'capture task data common to all assignments
        strTask = Task.Project & "," & Chr(34) & "[" & Task.UniqueID & "] " & Replace(Task.Name, Chr(34), Chr(39)) & Chr(34) & ","
        
        'examine every assignment on the task
        For Each Assignment In Task.Assignments
          
          'capture timephased work
          Set TSVS_WORK = Assignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
          For Each TSV In TSVS_WORK
            
            'capture common assignment data
            strRecord = strTask & Assignment.ResourceName & ","
            
            'optionally capture baseline work and cost
            If cptResourceDemand_frm.chkBaseline Then
              Set TSVS_BCWS = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
              If Assignment.ResourceType = pjResourceTypeWork Then
                strRecord = strRecord & Val(TSVS_BCWS(1).Value) / 60 & ","
              Else
                strRecord = strRecord & "0," 'Val(TSVS_BCWS(1).Value) & ","
              End If
              Set TSVS_BCWS = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks, 1)
              strRecord = strRecord & Val(TSVS_BCWS(1).Value) & ","
            End If
            'capture (and subtract) actual work, leaving ETC/Remaining Work
            Set TSVS_AW = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
            dblWork = Val(TSV.Value) - Val(TSVS_AW(1))
            If Assignment.ResourceType = pjResourceTypeWork Then
              strRecord = strRecord & dblWork / 60 & ","
            Else
              strRecord = strRecord & "0,"
            End If
            'get costs
            If blnIncludeCosts Then
              'rate set
              strRecord = strRecord & Choose(Assignment.CostRateTable + 1, "A", "B", "C", "D", "E") & ","
              Set TSVS_COST = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
              'get actual cost
              Set TSVS_AC = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
              'subtract actual cost from cost to get remaining cost
              dblCost = Val(TSVS_COST(1).Value) - Val(TSVS_AC(1))
              'get cost
              If dblWork > 0 Or dblCost > 0 Then 'there is remaining work or cost
                strRecord = strRecord & dblCost & ","
              Else
                strRecord = strRecord & "0,"
              End If
            End If
            
            'if default rate set is included then include it
            If cptResourceDemand_frm.chkA Then
              strRecord = strRecord & IIf(Assignment.CostRateTable = 0, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkB Then
              strRecord = strRecord & IIf(Assignment.CostRateTable = 1, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkC Then
              strRecord = strRecord & IIf(Assignment.CostRateTable = 2, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkD Then
              strRecord = strRecord & IIf(Assignment.CostRateTable = 3, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkE Then
              strRecord = strRecord & IIf(Assignment.CostRateTable = 4, dblCost, 0) & ","
            End If
            
            'get custom field values
            For lgExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
              lgField = cptResourceDemand_frm.lboExport.List(lgExport, 0)
              strRecord = strRecord & Task.GetField(lgField) & ","
            Next lgExport
            strRecord = strRecord & DateAdd("d", 1, TSV.StartDate) & "," 'week
            Debug.Print strRecord
            Print #lgFile, strRecord
            'todo: options for week beginning/ending and what day
            'todo: real estate - convert all custom field export listboxes to checkbox lists with +/- show all / show selected
          Next TSV
          
          'get rate set and cost
          lngOriginalRateSet = Assignment.CostRateTable
          If cptResourceDemand_frm.chkBaseline Then strRecord = strRecord & "0,0," 'BL HOURS, BL COST
          For lngRateSet = 0 To 4
            'need msproj to calculate the cost
            If cptResourceDemand_frm.Controls(Choose(lngRateSet + 1, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
              If lngRateSet = lngOriginalRateSet Then GoTo next_rate_set
              Application.StatusBar = "Exporting Rate Set " & Replace(Choose(lngRateSet + 1, "chkA", "chkB", "chkC", "chkD", "chkE"), "chk", "") & "..."
              If Assignment.CostRateTable <> lngRateSet Then Assignment.CostRateTable = lngRateSet 'recalculation not needed
              'extract timephased date
              'get work
              Set TSVS_WORK = Assignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
              For Each TSV In TSVS_WORK
                strRecord = Task.Project & "," & Chr(34) & "[" & Task.UniqueID & "] " & Replace(Task.Name, Chr(34), Chr(39)) & Chr(34) & ","
                strRecord = strRecord & Assignment.ResourceName & ","
                If cptResourceDemand_frm.chkBaseline Then strRecord = strRecord & "0,0," 'baseline placeholder
                strRecord = strRecord & "0," 'hours
                strRecord = strRecord & Choose(lngOriginalRateSet + 1, "A", "B", "C", "D", "E") & ","
                strRecord = strRecord & "0," 'cost
                'get cost
                Set TSVS_COST = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
                'get actual cost
                Set TSVS_AC = Assignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
                'subtract actual cost from cost to get remaining cost
                dblCost = Val(TSVS_COST(1).Value) - Val(TSVS_AC(1))
                With cptResourceDemand_frm
                  If .chkA Then strCost = "[0],"
                  If .chkB Then strCost = strCost & "[1],"
                  If .chkC Then strCost = strCost & "[2],"
                  If .chkD Then strCost = strCost & "[3],"
                  If .chkE Then strCost = strCost & "[4],"
                End With
                If dblCost > 0 Then
                  strCost = Replace(strCost, "[" & lngRateSet & "]", dblCost)
                  strCost = Replace(strCost, "[0]", "0")
                  strCost = Replace(strCost, "[1]", "0")
                  strCost = Replace(strCost, "[2]", "0")
                  strCost = Replace(strCost, "[3]", "0")
                  strCost = Replace(strCost, "[4]", "0")
                  strRecord = strRecord & strCost
                Else
                  strRecord = strRecord & Replace(String(lngRateSets, "0"), "0", "0,")
                End If
                
                'todo:if costs are getting pulled then pull rate table, rate, and cost
                'todo:add comment on rate set header that this is the currently applied rate table
                
                'get custom field values
                For lgExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
                  lgField = cptResourceDemand_frm.lboExport.List(lgExport, 0)
                  strRecord = strRecord & Task.GetField(lgField) & ","
                Next lgExport
                strRecord = strRecord & DateAdd("d", 1, TSV.StartDate) & "," 'week
                Debug.Print strRecord
                Print #lgFile, strRecord
              Next TSV
            End If
next_rate_set:
          Next lngRateSet
          If Assignment.CostRateTable <> lngOriginalRateSet Then Assignment.CostRateTable = lngOriginalRateSet

next_assignment:
        Next Assignment
      End If 'skip external tasks
    End If 'skip blank lines
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

  'is previous run still open?
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
  lgWeekCol = Worksheet.Rows(1).Find(what:="WEEK").Column
  dtMin = xlApp.WorksheetFunction.Min(Worksheet.Columns(lgWeekCol))
  dtMax = xlApp.WorksheetFunction.Max(Worksheet.Columns(lgWeekCol))
  
  Set Worksheet = Workbook.Sheets("SourceData")
  
  'format currencies
  For lgCol = 1 To lgWeekCol
    If InStr(Worksheet.Cells(1, lgCol), "COST") > 0 Then Worksheet.Columns(lgCol).Style = "Currency"
  Next lgCol
  
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
  'todo: revise this according to user options
  Worksheet.[B2] = "Weeks Beginning Monday"
  Worksheet.[B4].Select
  Worksheet.[B5].Select

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
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK")
    .Orientation = xlRowField
    .Position = 1
  End With
  If Not cptResourceDemand_frm.chkBaseline Then xlApp.ActiveSheet.PivotTables("PivotTable1").PivotFields("WEEK").PivotFilters.Add _
        Type:=xlAfter, Value1:=ActiveProject.StatusDate
  xlApp.ActiveChart.ClearToMatchStyle
  xlApp.ActiveChart.ChartStyle = 34
  xlApp.ActiveChart.ClearToMatchStyle
  xlApp.ActiveChart.SetElement (msoElementChartTitleAboveChart)
  xlApp.ActiveSheet.ChartObjects(1).Activate
  xlApp.ActiveChart.ChartTitle.Text = strTitle
  xlApp.ActiveChart.Location xlLocationAsNewSheet, "PivotChart"
  Set Worksheet = Workbook.Sheets("PivotChart_Source")
  Worksheet.Visible = False

  'add legend
  xlApp.ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
  xlApp.ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Hours"

  Application.StatusBar = "Saving the Workbook..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  'todo: export selected cost rate tables to worksheet
  If blnIncludeCosts Then
    Set Worksheet = Workbook.Sheets.Add(After:=Workbook.Sheets("SourceData"))
    Worksheet.Name = "Cost Rate Tables"
    Worksheet.[A1:F1].Value = Array("RESOURCE_NAME", "RATE_TABLE", "EFFECTIVE_DATE", "STANDARD_RATE", "OVERTIME_RATE", "PER_USE_COST")
    lngRow = 2
	'todo: make compatible with master-sub projects
    For Each Resource In ActiveProject.Resources
      Worksheet.Cells(lngRow, 1) = Resource.Name
      For Each CostRateTable In Resource.CostRateTables
        If cptResourceDemand_frm.Controls(Choose(CostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
          For Each PayRate In CostRateTable.PayRates
            Worksheet.Cells(lngRow, 1) = Resource.Name
            Worksheet.Cells(lngRow, 2) = CostRateTable.Name
            Worksheet.Cells(lngRow, 3) = Format(PayRate.EffectiveDate, "mm/dd/yyyy")
            Worksheet.Cells(lngRow, 4) = Replace(PayRate.StandardRate, "/h", "")
            Worksheet.Cells(lngRow, 5) = Replace(PayRate.OvertimeRate, "/h", "")
            Worksheet.Cells(lngRow, 6) = PayRate.CostPerUse
            lngRow = lngRow + 1
          Next PayRate
        End If
      Next CostRateTable
    Next Resource
  End If
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.[A1].AutoFilter
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  Worksheet.Columns.AutoFit
  
  Workbook.Sheets("ResourceDemand").Activate
  
  'save the file
  '<issue49> - file exists in location
  strFile = Environ("USERPROFILE") & "\Desktop\" & Replace(Workbook.Name, ".xlsx", "_" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".xlsx") '<issue49>
  If Dir(strFile) <> vbNullString Then '<issue49>
    If MsgBox("A file named '" & strFile & "' already exists in this location. Replace?", vbYesNo + vbExclamation, "Overwrite?") = vbYes Then '<issue49>
      Kill strFile '<issue49>
      Workbook.SaveAs strFile, 51 '<issue49>
      MsgBox "Saved to your Desktop:" & vbCrLf & vbCrLf & Dir(strFile), vbInformation + vbOKOnly, "Resource Demand Exported" '<issue49>
    End If '<issue49>
  Else '<issue49>
    Workbook.SaveAs strFile, 51  '<issue49>
  End If '</issue49>

  Application.StatusBar = "Complete."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar

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
  Call cptHandleErr("cptResourceDemand_bas", "cptExportResourceDemand", err, Erl)
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
  
  cptResourceDemand_frm.show False

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
    Call cptHandleErr("cptResourceDemand_bas", "ShowCptResourceDemand_frm", err, Erl)
    Resume exit_here
  End If

End Sub
