Attribute VB_Name = "cptGraphics_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowFrmGraphics()
  With cptGraphics_frm
    .cboMetric.AddItem "PLEASE CHOOSE A METRIC:"
    .cboMetric.AddItem "-----------------------"
    .cboMetric.AddItem "Bow Wave"
    .cboMetric.AddItem "Current Execution Index (CEI)"
    .cboMetric.AddItem "DCMA14.01 Logic"
    .cboMetric.AddItem "DCMA14.02 Leads"
    .cboMetric.AddItem "DCMA14.03 Lags"
    .cboMetric.AddItem "DCMA14.04 Relationship Types"
    .cboMetric.AddItem "DCMA14.05 Hard Constraints"
    .cboMetric.AddItem "DCMA14.06 High Float"
    .cboMetric.AddItem "DCMA14.07 Negative Float"
    .cboMetric.AddItem "DCMA14.08 High Duration"
    .cboMetric.AddItem "DCMA14.09 Invalid Dates"
    .cboMetric.AddItem "DCMA14.10 Resources"
    .cboMetric.AddItem "DCMA14.11 Missed Tasks"
    .cboMetric.AddItem "DCMA14.12 Critical Path Test"
    .cboMetric.AddItem "DCMA14.13 Critical Path Length Index (CPLI)"
    .cboMetric.AddItem "DCMA14.14 Baseline Execution Index (BEI)"
    
'    For lngItem = 0 To .cboMetric.ListCount - 1
'      Debug.Print "Case " & Chr(34) & .cboMetric.List(lngItem, 0) & Chr(34) & vbCrLf
'    Next lngItem
    
    .Show False
  End With
End Sub

Sub cptGetCharts()
'objects
Dim rst3 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst1 As ADODB.Recordset
Dim ChartObject As ChartObject
Dim PivotTable As PivotTable
Dim SubProject As Project
Dim wsDataSet3 As Worksheet
Dim wsDataSet2 As Worksheet
Dim wsDataSet1 As Worksheet
Dim xlApp As Excel.Application
Dim Workbook As Workbook
Dim Worksheet As Worksheet
Dim ListObject As ListObject
Dim rng As Excel.Range
Dim c As Excel.Range
Dim Task As MSProject.Task
Dim TaskDependency As TaskDependency
'strings
Dim strMethod As String
Dim strFileName As String
Dim strRange As String
'longs
Dim lngSubProject As Long
Dim lngTask As Long
Dim lngTasks As Long
Dim lngRow As Long
'integers
'doubles
'booleans
'variants
Dim vSheet As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  xlApp.ScreenUpdating = False
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = xlCalculationManual
  
  'dataset 1 - task uid, duration, total slack
  Set wsDataSet1 = Workbook.Sheets(1)
  wsDataSet1.Name = "DataSet1"
  'dataset 2 - task uid, baseline finish, actual finish, forecast finish - by week
  Set wsDataSet2 = Workbook.Sheets.Add
  wsDataSet2.Name = "DataSet2"
  'dataset 3 - task dependencies
  Set wsDataSet3 = Workbook.Sheets.Add
  wsDataSet3.Name = "DataSet3"
  
  'todo: include custom fields for user to filter out LOE, PP, etcs.
  strMethod = "ADODB"
  
  'set up headers
  wsDataSet1.[A1:C1] = Array("UID", "DURATION", "TOTAL_SLACK")
  wsDataSet2.[A1:E1] = Array("UID", "BL FINISH", "ACTUAL_FINISH", "FINISH", "WEEK_ENDING")
  wsDataSet3.[A1:J1] = Array("FROM_PROJECT", "FROM_ID", "FROM_UID", "FROM_TASK", "TYPE", "LAG", "TO_PROJECT", "TO_ID", "TO_UID", "TO_TASK")
  
  If strMethod = "ADODB" Then
    Set rst1 = CreateObject("ADODB.Recordset")
    With rst1
      .Fields.Append "UID", adInteger
      .Fields.Append "DURATION", adInteger
      .Fields.Append "TOTAL_SLACK", adInteger
      .Open
    End With
    Set rst2 = CreateObject("ADODB.Recordset")
    With rst2
      .Fields.Append "UID", adBigInt
      .Fields.Append "BL FINISH", adInteger
      .Fields.Append "ACTUAL_FINISH", adInteger
      .Fields.Append "FINISH", adInteger
      .Fields.Append "WEEK_ENDING", adDBDate
      .Open
    End With
    Set rst3 = CreateObject("ADODB.Recordset")
    With rst3
      .Fields.Append "FROM_PROJECT", adVarChar, 255
      .Fields.Append "FROM_ID", adInteger
      .Fields.Append "FROM_UID", adInteger
      .Fields.Append "FROM_TASK", adVarChar, 255
      .Fields.Append "TYPE", adVarChar, 2
      .Fields.Append "LAG", adInteger
      .Fields.Append "TO_PROJECT", adVarChar, 255
      .Fields.Append "TO_ID", adInteger
      .Fields.Append "TO_UID", adInteger
      .Fields.Append "TO_TASK", adVarChar, 255
      .Open
    End With
  End If
  
  'get task count
  If ActiveProject.Subprojects.Count > 0 Then
    lngTasks = ActiveProject.Tasks.Count
    For lngSubProject = 1 To ActiveProject.Subprojects.Count
      lngTasks = lngTasks + ActiveProject.Subprojects(lngSubProject).SourceProject.Tasks.Count
    Next lngSubProject
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  
  'extract data
  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then
      If Not Task.Active Then GoTo next_task
      If Task.Summary Then GoTo next_task
      If Task.BaselineWork > 0 Or Task.BaselineCost > 0 Then 'pmb task
        'dataset 1 - incomplete tasks only
        If strMethod = "Excel" Then
          lngRow = wsDataSet1.Cells(wsDataSet1.Rows.Count, 1).End(xlUp).Row + 1
          wsDataSet1.Cells(lngRow, 1) = Task.UniqueID
          wsDataSet1.Cells(lngRow, 2) = Task.Duration / (60 * 8)
          wsDataSet1.Cells(lngRow, 3) = Task.TotalSlack / (60 * 8)
        Else
          rst1.AddNew Array(0, 1, 2), Array(Task.UniqueID, Task.Duration / (60 * 8), Task.TotalSlack / (60 * 8))
          rst1.Update
        End If
        'dataset 2
        If Task.BaselineFinish <> "NA" Then
          If strMethod = "Excel" Then
            lngRow = wsDataSet2.Cells(wsDataSet2.Rows.Count, 1).End(xlUp).Row + 1
            wsDataSet2.Cells(lngRow, 1) = Task.UniqueID
            wsDataSet2.Cells(lngRow, 2) = 1 'baseline
            wsDataSet2.Cells(lngRow, 3) = 0 'actual
            wsDataSet2.Cells(lngRow, 3) = 0 'forecast
            wsDataSet2.Cells(lngRow, 5) = DateAdd("d", 6 - Weekday(Task.BaselineFinish), Task.BaselineFinish)
          Else
            rst2.AddNew Array(0, 1, 4), Array(Task.UniqueID, 1, DateAdd("d", 6 - Weekday(Task.BaselineFinish), Task.BaselineFinish))
            rst2.Update
          End If
        End If
        If Task.ActualFinish <> "NA" Then
          If strMethod = "Excel" Then
            lngRow = wsDataSet2.Cells(wsDataSet2.Rows.Count, 1).End(xlUp).Row + 1
            wsDataSet2.Cells(lngRow, 1) = Task.UniqueID
            wsDataSet2.Cells(lngRow, 2) = 0 'baseline
            wsDataSet2.Cells(lngRow, 3) = 1
            wsDataSet2.Cells(lngRow, 4) = 0 'forecast
            wsDataSet2.Cells(lngRow, 5) = DateAdd("d", 6 - Weekday(Task.ActualFinish), Task.ActualFinish)
          Else
            rst2.AddNew Array(0, 2, 4), Array(Task.UniqueID, 1, DateAdd("d", 6 - Weekday(Task.ActualFinish), Task.ActualFinish))
            rst2.Update
          End If
        End If
        If strMethod = "Excel" Then
          lngRow = wsDataSet2.Cells(wsDataSet2.Rows.Count, 1).End(xlUp).Row + 1
          wsDataSet2.Cells(lngRow, 1) = Task.UniqueID
          wsDataSet2.Cells(lngRow, 2) = 0 'baseline
          wsDataSet2.Cells(lngRow, 3) = 0 'actual
          wsDataSet2.Cells(lngRow, 4) = 1 'forecast
          wsDataSet2.Cells(lngRow, 5) = DateAdd("d", 6 - Weekday(Task.Finish), Task.Finish)
        Else
          rst2.AddNew Array(0, 3, 4), Array(Task.UniqueID, 1, DateAdd("d", 6 - Weekday(Task.Finish), Task.Finish))
          rst2.Update
        End If
        'dataset 3
        For Each TaskDependency In Task.TaskDependencies
          If strMethod = "Excel" Then
            lngRow = wsDataSet3.Cells(wsDataSet3.Rows.Count, 1).End(xlUp).Row + 1
            'this next bit gets tricky in master/sub
            If Task.Guid = TaskDependency.To.Guid Then 'use guid for master/sub
              wsDataSet3.Cells(lngRow, 1) = TaskDependency.From.Project
              wsDataSet3.Cells(lngRow, 2) = TaskDependency.From.ID
              wsDataSet3.Cells(lngRow, 3) = TaskDependency.From.UniqueID
              wsDataSet3.Cells(lngRow, 4) = TaskDependency.From.Name
              wsDataSet3.Cells(lngRow, 5) = Choose(TaskDependency.Type + 1, "FF", "FS", "SF", "SS")
              wsDataSet3.Cells(lngRow, 6) = TaskDependency.Lag / (60 * 8)
              wsDataSet3.Cells(lngRow, 7) = TaskDependency.To.Project
              wsDataSet3.Cells(lngRow, 8) = TaskDependency.To.ID
              wsDataSet3.Cells(lngRow, 9) = TaskDependency.To.UniqueID
              wsDataSet3.Cells(lngRow, 10) = TaskDependency.To.Name
            End If
          Else
            If Task.Guid = TaskDependency.To.Guid Then 'use guid for master/sub
              With TaskDependency.From
                rst3.AddNew Array(0, 1, 2, 3), Array(.Project, .ID, .UniqueID, .Name)
              End With
              rst3.Fields(4) = Choose(TaskDependency.Type + 1, "FF", "FS", "SF", "SS")
              rst3.Fields(5) = TaskDependency.Lag / (60 * 8)
              With TaskDependency.To
                rst3.Fields(6) = .Project
                rst3.Fields(7) = .ID
                rst3.Fields(8) = .UniqueID
                rst3.Fields(9) = .Name
              End With
              rst3.Update
            End If
          End If
        Next TaskDependency
      End If
    End If
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Analyzing Task " & Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & " (" & Format(lngTask / lngTasks, "0%") & ")"
    cptGraphics_frm.lblStatus = Application.StatusBar
    cptGraphics_frm.lblProgress.Width = (lngTask / lngTasks) * cptGraphics_frm.lblStatus.Width
    DoEvents
  Next Task
  cptGraphics_frm.lblProgress.Width = cptGraphics_frm.lblStatus.Width
  
  'make presentable
  For Each vSheet In Array(wsDataSet1, wsDataSet2, wsDataSet3)
    Set Worksheet = vSheet
    Worksheet.Activate
    xlApp.ActiveWindow.Zoom = 85
    Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)).Address, , xlYes)
    ListObject.TableStyle = ""
    Worksheet.Columns.AutoFit
  Next
  
  'copy in data
  If strMethod = "ADODB" Then
    wsDataSet1.[A2].CopyFromRecordset rst1
    wsDataSet2.[A2].CopyFromRecordset rst2
    wsDataSet2.ListObjects("Table2").ListColumns("WEEK_ENDING").Range.NumberFormat = "m/d/yyyy"
    wsDataSet3.[A2].CopyFromRecordset rst3
    rst1.Close
    rst2.Close
    rst3.Close
  End If
  
  xlApp.Calculation = xlCalculationAutomatic
  
  'save xlsx to cpt-backup temporarily
  cptGraphics_frm.lblStatus = "Saving Source Data..."
  strFileName = cptDir & "\metrics\"
  If Dir(strFileName, vbDirectory) = vbNullString Then MkDir strFileName
  strFileName = strFileName & Replace(Replace(ActiveProject.Name, ".mpp", ""), " ", "_") & "_Metrics_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx"
  Workbook.SaveAs strFileName, 51
  'create charts and export jpg files
  
  'bow wave
  Application.StatusBar = "Creating Bow Wave..."
  cptGraphics_frm.lblStatus = Application.StatusBar
  Set Worksheet = Workbook.Sheets.Add
  Worksheet.Name = "DataSet2_Graphs"
  strRange = "DataSet2!" & wsDataSet2.Range(wsDataSet2.[A1].End(xlToRight), wsDataSet2.[A1].End(xlDown)).Address
  Workbook.PivotCaches.Create(SourceType:=xlDatabase, _
                              SourceData:="Table2", _
                              Version:=6).CreatePivotTable TableDestination:="DataSet2_Graphs!R1C1", _
                                                            TableName:="DataSet2_PivotTable", _
                                                            DefaultVersion:=6
  Set Worksheet = Workbook.Sheets("DataSet2_Graphs")
  Set PivotTable = Worksheet.PivotTables(1)
  With PivotTable
      .ColumnGrand = True
      .HasAutoFormat = True
      .DisplayErrorString = False
      .DisplayNullString = True
      .EnableDrilldown = True
      .ErrorString = ""
      .MergeLabels = False
      .NullString = ""
      .PageFieldOrder = 2
      .PageFieldWrapCount = 0
      .PreserveFormatting = True
      .RowGrand = True
      .SaveData = True
      .PrintTitles = False
      .RepeatItemsOnEachPrintedPage = True
      .TotalsAnnotation = False
      .CompactRowIndent = 1
      .InGridDropZones = False
      .DisplayFieldCaptions = True
      .DisplayMemberPropertyTooltips = False
      .DisplayContextTooltips = True
      .ShowDrillIndicators = True
      .PrintDrillIndicators = False
      .AllowMultipleFilters = False
      .SortUsingCustomLists = True
      .FieldListSortAscending = False
      .ShowValuesRow = False
      .CalculatedMembersInFilters = False
      .RowAxisLayout 0 'xlCompactRow
  End With
  With PivotTable.PivotCache
      .RefreshOnFileOpen = False
      .MissingItemsLimit = -1 'xlMissingItemsDefault
  End With
  PivotTable.RepeatAllLabels 2 'xlRepeatLabels
  Worksheet.Shapes.AddChart2(201, 51).Select '51 = xlColumnClustered
  Set ChartObject = Worksheet.ChartObjects(1)
  ChartObject.Chart.SetSourceData Source:=Worksheet.Range("DataSet2_Graphs!$A$1:$C$18")
  Worksheet.Shapes("Chart 1").IncrementLeft 192
  Worksheet.Shapes("Chart 1").IncrementTop 15
  ChartObject.Chart.PivotLayout.PivotTable.AddDataField ChartObject.Chart.PivotLayout.PivotTable.PivotFields("BL FINISH"), "Sum of BL FINISH", xlSum
  ChartObject.Chart.PivotLayout.PivotTable.AddDataField ChartObject.Chart.PivotLayout.PivotTable.PivotFields("FINISH"), "Sum of FINISH", xlSum
  With ChartObject.Chart.PivotLayout.PivotTable.PivotFields("WEEK_ENDING")
      .Orientation = 1 'xlRowField
      .Position = 1
  End With
  ChartObject.Chart.PivotLayout.PivotTable.PivotFields("WEEK_ENDING").AutoGroup
  With ChartObject.Chart.PivotLayout.PivotTable.PivotFields("Sum of BL FINISH")
      .Calculation = 5 'xlRunningTotal
      .BaseField = "Years"
  End With
  With ChartObject.Chart.PivotLayout.PivotTable.PivotFields("Sum of FINISH")
      .Calculation = 5 'xlRunningTotal
      .BaseField = "Years"
  End With
  ChartObject.Chart.ChartType = 4 'xlLine
  ChartObject.Chart.ShowAllFieldButtons = False
  ChartObject.Chart.SetElement (2) '2 = msoElementChartTitleAboveChart
  ChartObject.Chart.ChartTitle.Caption = "Bow Wave Chart"
  ChartObject.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = -1 'msoTrue
  
  If Dir(cptDir & "\metrics\", vbDirectory) = vbNullString Then MkDir cptDir & "\metrics\"
  Worksheet.ChartObjects(1).Chart.Export cptDir & "\metrics\bow_wave.jpg", filtername:="JPG"
  
  'get DCMA14 Relationship Types
  Application.StatusBar = "Analyzing DCMA14-04 Relationship Types..."
  cptGraphics_frm.lblStatus = Application.StatusBar
  Set Worksheet = xlApp.Sheets.Add
  Worksheet.Name = "DataSet2_Graph"
  Workbook.PivotCaches.Create(SourceType:=xlDatabase, _
                              SourceData:="Table3", Version:=6).CreatePivotTable _
                              TableDestination:="DataSet2_Graph!R1C1", _
                              TableName:="DataSet2_PivotTable", _
                              DefaultVersion:=6
  Set PivotTable = Worksheet.PivotTables("DataSet2_PivotTable")
  With PivotTable
    .ColumnGrand = True
    .HasAutoFormat = True
    .DisplayErrorString = False
    .DisplayNullString = True
    .EnableDrilldown = True
    .ErrorString = ""
    .MergeLabels = False
    .NullString = ""
    .PageFieldOrder = 2
    .PageFieldWrapCount = 0
    .PreserveFormatting = True
    .RowGrand = True
    .SaveData = True
    .PrintTitles = False
    .RepeatItemsOnEachPrintedPage = True
    .TotalsAnnotation = False
    .CompactRowIndent = 1
    .InGridDropZones = False
    .DisplayFieldCaptions = True
    .DisplayMemberPropertyTooltips = False
    .DisplayContextTooltips = True
    .ShowDrillIndicators = True
    .PrintDrillIndicators = False
    .AllowMultipleFilters = False
    .SortUsingCustomLists = True
    .FieldListSortAscending = False
    .ShowValuesRow = False
    .CalculatedMembersInFilters = False
    .RowAxisLayout xlCompactRow
  End With
  With PivotTable.PivotCache
    .RefreshOnFileOpen = False
    .MissingItemsLimit = xlMissingItemsDefault
  End With
  PivotTable.RepeatAllLabels xlRepeatLabels
  Worksheet.Shapes.AddChart2(201, xlColumnClustered).Select
  Set ChartObject = Worksheet.ChartObjects(1)
  ChartObject.Chart.SetSourceData Source:=Worksheet.Range("DataSet2_Graph!$A$1:$C$18")
  Worksheet.Shapes("Chart 1").IncrementLeft 192
  Worksheet.Shapes("Chart 1").IncrementTop 15
  ChartObject.Chart.ChartType = xlPie
  ChartObject.Chart.PivotLayout.PivotTable.AddDataField ChartObject.Chart.PivotLayout.PivotTable.PivotFields("TYPE"), "Count of TYPE", xlCount
  With ChartObject.Chart.PivotLayout.PivotTable.PivotFields("TYPE")
      .Orientation = xlColumnField
      .Position = 1
  End With
  With ChartObject.Chart.PivotLayout.PivotTable.PivotFields("TYPE")
      .Orientation = xlRowField
      .Position = 1
  End With
  ChartObject.Chart.ApplyLayout (4) 'needed?
  
  ChartObject.Chart.SetElement (msoElementDataLabelBestFit)
  With PivotTable.PivotFields("Count of TYPE")
      .Calculation = xlPercentOfTotal
      .NumberFormat = "0.00%"
  End With
  ChartObject.Chart.SetElement (msoElementDataLabelCenter)
  ChartObject.Chart.SetElement (msoElementChartTitleAboveChart)
  ChartObject.Chart.ChartTitle.Caption = "DCMA14.04 - Relationship Types"
  ChartObject.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
  ChartObject.Chart.ShowAllFieldButtons = False
  PivotTable.PivotFields("TYPE").AutoSort xlDescending, "Count of TYPE"
  ChartObject.Chart.Export cptDir & "\metrics\dcma14-04.jpg", filtername:="JPG"
  
  'create each one
  'create a nice one-pager
  'copy/paste each chart pic into the workbook
  
exit_here:
  On Error Resume Next
  Set rst3 = Nothing
  Set rst2 = Nothing
  Set rst1 = Nothing
  Set ChartObject = Nothing
  Set PivotTable = Nothing
  Application.StatusBar = ""
  cptGraphics_frm.lblStatus = "Ready..."
  cptGraphics_frm.lblProgress.Width = cptGraphics_frm.lblStatus.Width
  'do not delete metrics dir or file
  Set SubProject = Nothing
  Set wsDataSet3 = Nothing
  Set wsDataSet2 = Nothing
  Set wsDataSet1 = Nothing
  Set wsDataSet1 = Nothing
  Set c = Nothing
  Set rng = Nothing
  Set ListObject = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  xlApp.Visible = True
  xlApp.ScreenUpdating = True
  xlApp.Calculation = xlCalculationAutomatic
  Set xlApp = Nothing
  Set Task = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptGraphics_bas", "cptGetChart", err, Erl)
  Resume exit_here
End Sub
