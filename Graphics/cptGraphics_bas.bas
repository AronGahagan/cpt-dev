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
    .Show False
  End With
End Sub

Sub cptGetChart(strMetric As String)
'objects
Dim xlApp As Excel.Application
Dim Workbook As Workbook
Dim Worksheet As Worksheet
Dim ListObject As ListObject
Dim rng As Excel.Range
Dim c As Excel.Range
Dim Task As MSProject.Task
'strings
'longs
Dim lngRow As Long
'integers
'doubles
'booleans
'variants
'dates

If strMetric <> "Bow Wave" Then
  cptGraphics_frm.imgGraph.Picture = LoadPicture("")
  GoTo exit_here
Else
  GoTo start_here
End If

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  xlApp.ScreenUpdating = False
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = xlCalculationManual
  Set Worksheet = Workbook.Sheets(1)

  Worksheet.[A1:E1] = Array("UID", "BL FINISH", "ACTUAL_FINISH", "FINISH", "WEEK_ENDING")

  'would an exported map be faster? would it work with master/subs?

  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then
      If Task.BaselineWork > 0 Or Task.BaselineCost > 0 Then 'pmb task
        lngRow = Worksheet.Cells(Worksheet.Rows.count, 1).End(xlUp).Row + 1
        Worksheet.Cells(lngRow, 1) = Task.UniqueID
        Worksheet.Cells(lngRow, 2) = IIf(Task.BaselineFinish = "NA", 0, 1) 'baseline
        Worksheet.Cells(lngRow, 3) = 0 'actual
        Worksheet.Cells(lngRow, 3) = 0 'forecast
        Worksheet.Cells(lngRow, 5) = DateAdd("d", 6 - Weekday(Task.BaselineFinish), Task.BaselineFinish)
        lngRow = Worksheet.Cells(Worksheet.Rows.count, 1).End(xlUp).Row + 1
        Worksheet.Cells(lngRow, 1) = Task.UniqueID
        Worksheet.Cells(lngRow, 2) = 0 'baseline
        Worksheet.Cells(lngRow, 3) = IIf(Task.ActualFinish = "NA", 0, 1) 'actual
        Worksheet.Cells(lngRow, 4) = 0 'forecast
        Worksheet.Cells(lngRow, 5) = DateAdd("d", 6 - Weekday(Task.ActualFinish), Task.ActualFinish)
        lngRow = Worksheet.Cells(Worksheet.Rows.count, 1).End(xlUp).Row + 1
        Worksheet.Cells(lngRow, 1) = Task.UniqueID
        Worksheet.Cells(lngRow, 2) = 0 'baseline
        Worksheet.Cells(lngRow, 3) = 0 'actual
        Worksheet.Cells(lngRow, 4) = 1 'forecast
        Worksheet.Cells(lngRow, 5) = DateAdd("d", 6 - Weekday(Task.Finish), Task.Finish)
      End If
    End If
    Debug.Print lngRow & " of " & ActiveProject.Tasks.count
  Next Task
  
  xlApp.Visible = True
  xlApp.Calculation = xlCalculationAutomatic
  xlApp.ScreenUpdating = True
  
  
  Workbook.Sheets.Add
  Workbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
      "Sheet1!R1C1:R1651C4", Version:=6).CreatePivotTable TableDestination:= _
      "Sheet2!R1C1", TableName:="PivotTable1", DefaultVersion:=6
  Set Worksheet = Workbook.Sheets("Sheet2")
  Worksheet.Cells(1, 1).Select
  With Worksheet.PivotTables("PivotTable1")
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
  With Worksheet.PivotTables("PivotTable1").PivotCache
      .RefreshOnFileOpen = False
      .MissingItemsLimit = xlMissingItemsDefault
  End With
  Worksheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
  Worksheet.Shapes.AddChart2(201, xlColumnClustered).Select
  xlApp.ActiveChart.SetSourceData Source:=Worksheet.Range("Sheet2!$A$1:$C$18")
  Worksheet.Shapes("Chart 1").IncrementLeft 192
  Worksheet.Shapes("Chart 1").IncrementTop 15
  xlApp.ActiveChart.PivotLayout.PivotTable.AddDataField xlApp.ActiveChart.PivotLayout. _
      PivotTable.PivotFields("BL FINISH"), "Sum of BL FINISH", xlSum
  xlApp.ActiveChart.PivotLayout.PivotTable.AddDataField xlApp.ActiveChart.PivotLayout. _
      PivotTable.PivotFields("FINISH"), "Sum of FINISH", xlSum
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK_ENDING")
      .Orientation = xlRowField
      .Position = 1
  End With
  xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK_ENDING").AutoGroup
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of BL FINISH")
      .Calculation = xlRunningTotal
      .BaseField = "Years"
  End With
  With xlApp.ActiveChart.PivotLayout.PivotTable.PivotFields("Sum of FINISH")
      .Calculation = xlRunningTotal
      .BaseField = "Years"
  End With
  xlApp.ActiveChart.ChartType = xlLine
  xlApp.ActiveChart.ShowAllFieldButtons = False
  xlApp.ActiveChart.SetElement (msoElementChartTitleCenteredOverlay)
  xlApp.Selection.Caption = "Bow Wave Chart"
  xlApp.Selection.Format.TextFrame2.TextRange.Font.Bold = msoTrue
  xlApp.ActiveChart.SetElement (msoElementChartTitleAboveChart)
  xlApp.ActiveChart.ChartArea.Select
  
  Worksheet.ChartObjects(1).Chart.Export Environ("USERPROFILE") & "\Desktop\test.jpg", filtername:="JPG"
start_here:
  cptGraphics_frm.imgGraph.Picture = LoadPicture(Environ("USERPROFILE") & "\Desktop\test.jpg", 488, 288)
  
exit_here:
  On Error Resume Next
  Set c = Nothing
  Set rng = Nothing
  Set ListObject = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set Task = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("basGraphics_bas", "GetChart", err, Erl)
  Resume exit_here
End Sub
