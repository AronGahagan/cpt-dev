Attribute VB_Name = "cptQuickMonte_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptQuickMonte()
'objects
Dim Chart As Excel.Chart
Dim PivotTable As PivotTable
Dim ListObject As ListObject
Dim Worksheet As Excel.Worksheet
Dim Workbook As Excel.Workbook
Dim xlApp As Excel.Application
Dim rst As ADODB.Recordset
Dim arrDurations As SortedList 'Object
Dim Task As Task
'strings
'longs
Dim lngMLDur As Long
Dim lngMaxDur As Long
Dim lngMinDur As Long
Dim lngMax As Long
Dim lngMin As Long
Dim lngIteration As Long
Dim lngIterations As Long
Dim lngItem As Long
'integers
'doubles
Dim dblStDev As Double
Dim dblMean As Double
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngIterations = 100
  
  cptSpeed True
  
  'capture most likely durations
  Set arrDurations = CreateObject("System.Collections.SortedList")
  For Each Task In ActiveProject.Tasks
    If Task Is Nothing Then GoTo next_task0
    arrDurations.Add Task.UniqueID, Task.RemainingDuration
next_task0:
  Next Task
  
  'get three-point fields
  lngMin = FieldNameToFieldConstant("MinDuration")
  lngMax = FieldNameToFieldConstant("MaxDuration")
  
  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "ITERATION", adInteger
  rst.Fields.Append "UID", adInteger
  rst.Fields.Append "FINISH", adDate
  rst.Open
  
  Randomize
  
  'run iterations and export to adtg
  For lngIteration = 1 To lngIterations
    'simulate project
    For Each Task In ActiveProject.Tasks
      'todo: if NOT min < ml < max
      If Task.RemainingDuration = 0 Then GoTo next_task1
      lngMinDur = cptRegEx(Task.GetField(lngMin), "[0-9].") * 480
      lngMaxDur = cptRegEx(Task.GetField(lngMax), "[0-9].") * 480
      'calculate mean using PERT todo: use triangular instead?
      dblMean = (lngMinDur + (Task.RemainingDuration * 4) + lngMaxDur) / 6
      dblStDev = WorksheetFunction.StDev_P(lngMinDur, arrDurations.Item(Task.UniqueID), lngMaxDur)
      dblP = Math.Rnd
      'todo: use triangular https://www.drdawnwright.com/easy-excel-inverse-triangular-distribution-for-monte-carlo-simulations/
      lngMLDur = WorksheetFunction.Norm_Inv(dblP, dblMean, dblStDev)
      If lngMLDur > 0 Then Task.RemainingDuration = lngMLDur
next_task1:
    Next Task
    
    'todo: create array of [iterations]
    'avoid
    
    'calculate project
    'Application.StatusBar = "Calculating..."
    CalculateProject
    'ScreenUpdating = True
    
    'capture simulation
    For Each Task In ActiveProject.Tasks
      rst.AddNew Array(0, 1, 2), Array(lngIteration, Task.UniqueID, Task.Finish)
      rst.Update
    Next Task
    Application.StatusBar = "Running Simulation " & lngIteration & " of " & lngIterations & "...(" & Format(lngIteration / lngIterations, "0%") & ")"
    DoEvents
  Next lngIteration
  
  'restore most likely durations
  For lngItem = 0 To arrDurations.Count - 1
    Set Task = ActiveProject.Tasks.UniqueID(arrDurations.getKey(lngItem))
    Task.RemainingDuration = arrDurations.getValueList()(lngItem)
  Next
  CalculateProject
  
  cptSpeed False
  
  'export results
  Application.StatusBar = "Creating Report..."
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = "cptQuickMonte Data"
  Worksheet.[A1:C1] = Array("ITERATION", "UID", "FINISH")
  Worksheet.[A2].CopyFromRecordset rst
  rst.Close
  xlApp.Visible = True
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.Columns.AutoFit
  Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)))
  ListObject.Name = "QuickMonte"
  
  'todo: add task name
  
'  'create chart
'  strRange = Worksheet.Name & "!" & Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)).Address(False, False, xlR1C1)
'  'todo: must create on new worksheet, then move it
'  Set Worksheet = Workbook.Sheets.Add
'  Worksheet.Name = "Chart"
'  Workbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
'        strRange, Version:=6).CreatePivotTable TableDestination:= _
'        "Chart!R1C5", TableName:="PivotTable1", DefaultVersion:=6
'  Set PivotTable = Worksheet.PivotTables("PivotTable1")
'  With PivotTable
'      .ColumnGrand = True
'      .HasAutoFormat = True
'      .DisplayErrorString = False
'      .DisplayNullString = True
'      .EnableDrilldown = True
'      .ErrorString = ""
'      .MergeLabels = False
'      .NullString = ""
'      .PageFieldOrder = 2
'      .PageFieldWrapCount = 0
'      .PreserveFormatting = True
'      .RowGrand = True
'      .SaveData = True
'      .PrintTitles = False
'      .RepeatItemsOnEachPrintedPage = True
'      .TotalsAnnotation = False
'      .CompactRowIndent = 1
'      .InGridDropZones = False
'      .DisplayFieldCaptions = True
'      .DisplayMemberPropertyTooltips = False
'      .DisplayContextTooltips = True
'      .ShowDrillIndicators = True
'      .PrintDrillIndicators = False
'      .AllowMultipleFilters = False
'      .SortUsingCustomLists = True
'      .FieldListSortAscending = False
'      .ShowValuesRow = False
'      .CalculatedMembersInFilters = False
'      .RowAxisLayout xlCompactRow
'  End With
'  With PivotTablePivotCache
''      .RefreshOnFileOpen = False
''      .MissingItemsLimit = xlMissingItemsDefault
'  End With
'  PivotTable.RepeatAllLabels xlRepeatLabels
'  Worksheet.[A1].Select
'  Worksheet.Shapes.AddChart2.Select
'
'  Set PivotChart = Worksheet.Shapes.AddChart2(201, xlColumnClustered)
'  Chart.SetSourceData Source:=Worksheet.Range("'cptQuickMonte Data'!$E$1:$G$18")
'  With Chart.PivotLayout.PivotTable.PivotFields("UID")
'      .Orientation = xlPageField
'      .Position = 1
'  End With
'  With Chart.PivotLayout.PivotTable.PivotFields("FINISH")
'      .Orientation = xlRowField
'      .Position = 1
'  End With
'  Chart.PivotLayout.PivotTable.PivotFields("FINISH").AutoGroup
'  Worksheet.Range("E5").Group Start:=True, End:=True, By:=1, Periods:=Array(False, _
'      False, False, True, False, False, False)
'  PivotTable.AddDataField PivotTable.PivotFields("FINISH"), "Count of FINISH", xlCount
'  PivotTable.PivotFields("UID").ClearAllFilters
'  PivotTable.PivotFields("UID").CurrentPage = "16"
  
  'copy charts into each task
    
  MsgBox "Simluation Complete", vbInformation + vbOKOnly, "QuickMonte"
    
  Application.StatusBar = "Complete"
  
exit_here:
  On Error Resume Next
  Set Chart = Nothing
  Set PivotTable = Nothing
  Application.StatusBar = ""
  Set ListObject = Nothing
  If Not xlApp Is Nothing Then xlApp.Visible = True
  cptSpeed False
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set rst = Nothing
  Set arrDurations = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMonte", "cptQuickMonte", Err, Erl)
  Resume exit_here
End Sub
