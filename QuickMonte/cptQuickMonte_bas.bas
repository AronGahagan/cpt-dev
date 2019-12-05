Attribute VB_Name = "cptQuickMonte_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptQuickMonte(lngMin As Long, lngMax As Long)
'objects
Dim sProject As SubProject
Dim rng As Excel.Range
Dim CE As MSProject.Exception
Dim Chart As Excel.Chart
Dim ListObject As ListObject
Dim Worksheet As Excel.Worksheet
Dim Workbook As Excel.Workbook
Dim xlApp As Excel.Application
Dim rst3p As ADODB.Recordset
Dim rstSim As ADODB.Recordset
Dim Task As Task
'strings
'longs
Dim lngTasks As Long
Dim lngItems As Long
Dim lngTask As Long
Dim lngUID As Long
Dim lngDays As Long
Dim lngX As Long
Dim lngMLDur As Long
Dim lngMaxDur As Long
Dim lngMinDur As Long
Dim lngIteration As Long
Dim lngIterations As Long
Dim lngItem As Long
'integers
'doubles
Dim dblP As Double
Dim dblCDF_ML As Double
'booleans
Dim blnChangeHighlighting As Boolean
Dim blnFail As Boolean
'variants
Dim vbResponse As Variant
'dates
Dim dtDeterministicFinish As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'get user input
  'todo: enter this on a UserForm
  'todo: capture input on UserForm and only allow numeric
  vbResponse = InputBox("How many iterations?", "QuickMonte", 1000)
  'validate input
  'todo: remove input validation after UserForm
  If StrPtr(vbResponse) = 0 Then 'user hit cancel
    GoTo exit_here
  ElseIf vbResponse = vbNullString Then 'user entered null value
    GoTo exit_here
  Else 'go with it
    lngIterations = cptRegEx(CStr(vbResponse), "[0-9]*")
    If lngIterations = 0 Then GoTo exit_here
  End If
  
  'capture selected
  On Error Resume Next
  lngUID = ActiveSelection.Tasks(1).UniqueID
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  'select last task by default if no task selected
  'todo: put note on button hover to select the target task
  If lngUID = 0 Then lngUID = ActiveProject.Tasks(ActiveProject.Tasks.Count).UniqueID
  
  'speed up processing and prevent screen flicker
  cptSpeed True
  
  'get task count
  lngTasks = 0
  If ActiveProject.Subprojects.Count = 0 Then
    For Each sProject In ActiveProject.Subprojects
      lngTasks = lngTasks + sProject.SourceProject.Tasks.Count
    Next sProject
  Else
    lngTasks = ActiveProject.Tasks.Count
  End If
  
  'todo: capture, remove, restore deadlines and constraints?
  
  'capture three points
  cptQuickMonte_frm.lblStatus.Caption = "Capturing three points..."
  Set rst3p = CreateObject("ADODB.Recordset")
  rst3p.Fields.Append "UID", adBigInt
  rst3p.Fields.Append "MIN", adBigInt
  rst3p.Fields.Append "ML", adBigInt
  rst3p.Fields.Append "MAX", adBigInt
  rst3p.Fields.Append "SM", adBoolean
  rst3p.Open
  lngTask = 0
  For Each Task In ActiveProject.Tasks
    If Task Is Nothing Then GoTo next_task0
    If Task.Summary Then GoTo next_task0
    If Task.ExternalTask Then GoTo next_task0
    If Not Task.Active Then GoTo next_task0
    If Task.RemainingDuration = 0 Then GoTo next_task0
    'catch and ignore schedule margin > 0
    If InStr(Task.Name, "SCHEDULE MARGIN") > 0 Then
      If Task.RemainingDuration > 0 Then
        cptQuickMonte_frm.lblStatus.Caption = "SCHEDULE MARGIN FOUND"
        'prompt user
        MsgBox "Task UID " & Task.UniqueID & " '" & Task.Name & "' will be ignored for the simulations.", vbInformation + vbOKOnly, "SCHEDULE MARGIN FOUND"
        'save it
        lngMLDur = Task.RemainingDuration
        rst3p.AddNew Array(0, 1, 2, 3, 4), Array(Task.UniqueID, 0, lngMLDur, 0, True)
        'overwrite it
        Task.RemainingDuration = 0
      Else
        'ignore zero-duration schedule margin
      End If
    'capture and ignore LOE
    ElseIf Task.GetField(FieldNameToFieldConstant("EVT2")) = "A" Then 'todo: user must set this
      GoTo next_task0
    ElseIf Task.GetField(FieldNameToFieldConstant("EVT2")) = "N/A" Then 'todo: user must set this
      GoTo next_task0
    ElseIf Task.GetField(FieldNameToFieldConstant("EVT2")) = "" Then 'todo: user must set this
      GoTo next_task0
      'todo: use a sortedlist to capture exceptions
    Else
      'todo: what if user selects a text or number field?
      lngMinDur = cptGetDuration(Task, lngMin)
      lngMaxDur = cptGetDuration(Task, lngMax)
      lngMLDur = Task.RemainingDuration
      rst3p.AddNew Array(0, 1, 2, 3, 4), Array(Task.UniqueID, lngMinDur, lngMLDur, lngMaxDur, False)
    End If
next_task0:
    lngTask = lngTask + 1
    cptQuickMonte_frm.lblStatus.Caption = "Capturing three points...(" & Format(lngTask / lngTasks, 0) & ")"
    cptQuickMonte_frm.lblProgress.Width = (lngTask / lngTasks) * cptQuickMonte_frm.lblStatus.Width
  Next Task
  
  'clean it up
  cptQuickMonte_frm.lblStatus.Caption = "Capturing three points...(100%)"
  cptQuickMonte_frm.lblProgress.Width = cptQuickMonte_frm.lblStatus.Width
  
  'in case schedule margin was removed:
  CalculateProject 'once
  'get deterministic finish of target tasks after margin removal
  dtDeterministicFinish = ActiveProject.Tasks.UniqueID(lngUID).Finish
  
  'prepare to capture simulation results
  Application.StatusBar = "Preparing to run simulations..."
  Set rstSim = CreateObject("ADODB.Recordset")
  rstSim.Fields.Append "ITERATION", adInteger
  rstSim.Fields.Append "UID", adBigInt
  rstSim.Fields.Append "R_DUR", adBigInt
  rstSim.Fields.Append "FINISH", adDate
  rstSim.Open
  
  Randomize
  
  'run iterations and export to adtg
  cptQuickMonte_frm.lblStatus.Caption = "Running Simulation " & 1 & " of " & lngIterations & "...(" & Format(1 / lngIterations, "0%") & ")"
  cptQuickMonte_frm.lblProgress.Width = (lngIteration / lngIterations) * cptQuickMonte_frm.lblStatus.Width
  For lngIteration = 1 To lngIterations
    'simulate project
    rst3p.MoveFirst
    Do While Not rst3p.EOF
      Set Task = ActiveProject.Tasks.UniqueID(rst3p("UID"))
      'skip schedule margin tasks
      If rst3p("SM") = True Then GoTo next_task1
      lngMinDur = rst3p("MIN")
      lngMLDur = rst3p("ML")
      lngMaxDur = rst3p("MAX")
      blnFail = False
      'validate three points
      If lngMinDur >= lngMLDur Or lngMLDur >= lngMaxDur Then
        MsgBox "Task UID " & Task.UniqueID & " '" & Task.Name & "' has invalid three point estimates.", vbCritical + vbOKOnly, "Error"
        blnFail = True
        'todo: editgoto? mark it then filter?
        GoTo restore_durations
      End If
      'determine CDF of ML value
      dblCDF_ML = (lngMLDur - lngMinDur) / (lngMaxDur - lngMLDur)
      'get random probability
      dblP = Math.Rnd
      'credit for the following goes the discussion on this website:
      'https://www.drdawnwright.com/easy-excel-inverse-triangular-distribution-for-monte-carlo-simulations/
      If dblP <= dblCDF_ML Then
        'min+sqrt(dblP*(max-min)*(ml-min))
        lngX = lngMinDur + Math.Sqr(dblP * (lngMaxDur - lngMinDur) * (lngMLDur - lngMinDur))
      Else
        'max-sqrt((1-dblP)*(max-min)*(-ml+max)))
        lngX = lngMaxDur - Math.Sqr((1 - dblP) * (lngMaxDur - lngMinDur) * (-lngMLDur + lngMaxDur))
      End If
      'Application.Find "Unique ID", "equals", rst3p("UID")
      Task.RemainingDuration = lngX
next_task1:
      rst3p.MoveNext
    Loop
    
    CalculateProject
    
    'capture simulation
    For Each Task In ActiveProject.Tasks
      If Task Is Nothing Then GoTo next_task2
      rstSim.AddNew Array(0, 1, 2, 3), Array(lngIteration, Task.UniqueID, Task.RemainingDuration, Task.Finish)
next_task2:
    Next Task
    cptQuickMonte_frm.lblStatus.Caption = "Running Simulation " & lngIteration & " of " & lngIterations & "...(" & Format(lngIteration / lngIterations, "0%") & ")"
    cptQuickMonte_frm.lblProgress.Width = (lngIteration / lngIterations) * cptQuickMonte_frm.lblStatus.Width
    DoEvents
  Next lngIteration
  
  'clean it up
  cptQuickMonte_frm.lblStatus.Caption = "Running Simulation " & lngIteration & " of " & lngIterations & "...(100%)"
  cptQuickMonte_frm.lblProgress.Width = cptQuickMonte_frm.lblStatus.Width
  
restore_durations:
  cptQuickMonte_frm.lblStatus.Caption = "Restoring remaining durations..."
  rst3p.MoveFirst
  lngItem = 0
  lngItems = rst3p.RecordCount
  Do While Not rst3p.EOF
    Set Task = ActiveProject.Tasks.UniqueID(rst3p("UID"))
'    If Task.Summary Then GoTo next_task3
'    If Task.ExternalTask Then GoTo next_task3
'    If Not Task.Active Then GoTo next_task3
    Task.RemainingDuration = CLng(rst3p("ML"))
next_task3:
    lngItem = lngItem + 1
    cptQuickMonte_frm.lblStatus.Caption = "Restoring remaining durations...(" & Format(lngItem / lngItems, "0%") & ")"
    cptQuickMonte_frm.lblProgress.Width = (lngItem / lngItems) * cptQuickMonte_frm.lblStatus.Width
    rst3p.MoveNext
  Loop
  
  'clean it up
  cptQuickMonte_frm.lblStatus.Caption = "Restoring remaining durations...(100%)"
  cptQuickMonte_frm.lblProgress.Width = cptQuickMonte_frm.lblStatus.Width
  
  'capture enable highlighting setting and turn off
  blnChangeHighlighting = Application.EnableChangeHighlighting
  Application.EnableChangeHighlighting = False
  'calculate project - goal is to prevent screen from changing at all
  'todo: prevent gantt changes
  CalculateProject
  cptSpeed False
  'restore highlighting settings
  Application.EnableChangeHighlighting = blnChangeHighlighting
  cptQuickMonte_frm.lblStatus.Caption = "Complete"
  DoEvents
  If blnFail Then GoTo exit_here
  
  If MsgBox("Simluation Complete" & vbCrLf & vbCrLf & "Create Report?", vbInformation + vbYesNo, "QuickMonte") = vbYes Then
  
    'export results
    cptQuickMonte_frm.lblStatus.Caption = "Creating Report..."
    Set xlApp = CreateObject("Excel.Application")
    xlApp.WindowState = xlMaximized
    Set Workbook = xlApp.Workbooks.Add
    Set Worksheet = Workbook.Sheets(1)
    xlApp.ActiveWindow.DisplayGridlines = False
    xlApp.Visible = True 'todo: debug only
    Worksheet.Name = "cptQuickMonte_DATA"
    xlApp.ScreenUpdating = False
    Worksheet.[A1:D1].Merge
    Worksheet.[A1:D1].Value = "SIMULATION RESULTS"
    Worksheet.[A1:D1].HorizontalAlignment = xlCenter
    Worksheet.[A2:D2] = Array("ITERATION", "UID", "REMAINING DURATION", "FINISH")
    Worksheet.[A3].CopyFromRecordset rstSim
    rstSim.Close
    xlApp.ActiveWindow.Zoom = 85
    Worksheet.Columns.AutoFit
    Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A2].End(xlToRight), Worksheet.[A2].End(xlDown)))
    ListObject.Name = "QuickMonte"
    xlApp.ScreenUpdating = True
    
    'add informational column
    xlApp.ScreenUpdating = False
    Worksheet.[F1:F12] = WorksheetFunction.Transpose(Array("UID:", "Deterministic:", "Iterations:", "Confidence:", "Confidence Date:", "Margin Rec.:", "Min:", "Max:", "Mean:", "Range:", "Bin Count:", "Bin Size:"))
    Worksheet.[F1:F12].Font.Bold = True
    Worksheet.Columns("F:F").AutoFit
    
    'add freq chart titles
    Worksheet.[G1:G12].HorizontalAlignment = xlCenter
    Worksheet.[G1].Style = "Input"
    Worksheet.[G2:G3].Style = "Calculation"
    Worksheet.[G4].Style = "Input"
    Worksheet.[G5:G6].Style = "Calculation"
    Worksheet.[G1].Value = lngUID
    Worksheet.[G2].NumberFormat = "mm/dd/yy"
    Worksheet.[G2].Value = dtDeterministicFinish
    Worksheet.[G3].Value = lngIterations
    Worksheet.[G4].Value = 0.9 'todo: get this value from user form
    Worksheet.[G7].NumberFormat = "mm/dd/yy"
    Worksheet.[G7].FormulaR1C1 = "=ROUND(MINIFS(QuickMonte[FINISH],QuickMonte[UID],R1C7),0)"
    Worksheet.[G8].NumberFormat = "mm/dd/yy"
    Worksheet.[G8].FormulaR1C1 = "=ROUND(MAXIFS(QuickMonte[FINISH],QuickMonte[UID],R1C7),0)"
    Worksheet.[G9].NumberFormat = "mm/dd/yy"
    Worksheet.[G9].FormulaR1C1 = "=ROUND(AVERAGEIFS(QuickMonte[FINISH],QuickMonte[UID],R[-8]C),0)"
    Worksheet.[G10].FormulaR1C1 = "=DAYS(R[-2]C,R[-3]C)"
    Worksheet.[G11].Value = 25
    Worksheet.[G12].FormulaR1C1 = "=R10C7/R11C7"
    xlApp.ScreenUpdating = True
    
    'capture exceptions in [Y14]
    xlApp.ScreenUpdating = False
    Worksheet.[Y13:Z13].Merge
    Worksheet.[Y13].HorizontalAlignment = xlCenter
    Worksheet.[Y13].Value2 = "CALENDAR EXCEPTIONS"
    Worksheet.[Y14:Z14] = Array("NAME", "DATE")
    For Each CE In ActiveProject.Calendar.Exceptions
      For lngDays = 0 To CE.Occurrences - 1
        Worksheet.Cells(Worksheet.[Y13].End(xlDown).Row + 1, 25) = CE.Name
        Worksheet.Cells(Worksheet.[Y13].End(xlDown).Row, 26) = DateAdd("d", lngDays, CE.Start)
      Next lngDays
    Next CE
    Worksheet.Range(Worksheet.[Z14], Worksheet.[Z14].End(xlDown)).NumberFormat = "mm/dd/yyyy"
    Worksheet.[Z14].NumberFormat = "General"
    Worksheet.Columns("Y:Z").AutoFit
    Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[Y14].End(xlToRight), Worksheet.[Y14].End(xlDown)))
    ListObject.Name = "HOLIDAYS"
    xlApp.ScreenUpdating = True
    
    'create frequency distribution chart
    xlApp.ScreenUpdating = False
    Worksheet.[F14:L14].Font.Bold = True
    Worksheet.[F14:L14] = Array("LL", "UL", "UL TITLE", "Freq", "Cum Freq", "Freq %", "Cum %")
    Worksheet.[F15].FormulaR1C1 = "=R[-8]C[1]-R12C7"
    Worksheet.[F15:F42].NumberFormat = "mm/dd/yy"
    Worksheet.[F16:F42].FormulaR1C1 = "=R[-1]C+R12C7"
    Worksheet.[G15:G41].FormulaR1C1 = "=R[1]C[-1]-0.0001"
    Worksheet.[H15:H41].FormulaR1C1 = "=ROUND(RC[-1],0)"
    'credit for the filtered frequency formula goes to ExcelJet:
    'https://exceljet.net/formula/count-unique-text-values-with-criteria
    Worksheet.[I15:I41].FormulaArray = "=FREQUENCY(IF(QuickMonte[UID]=$G$1,QuickMonte[FINISH]),$G$15:$G$41)"
    Worksheet.[J15].FormulaR1C1 = "=RC[-1]"
    Worksheet.[J16:J41].FormulaR1C1 = "=R[-1]C+RC[-1]"
    Worksheet.[K15:K41].FormulaR1C1 = "=RC[-2]/R3C7"
    Worksheet.[L15].FormulaR1C1 = "=RC[-1]"
    Worksheet.[L16:L41].FormulaR1C1 = "=R[-1]C+RC[-1]"
    'center the distribution table
    Worksheet.Range(Worksheet.[F14].End(xlToRight), Worksheet.[F14].End(xlDown)).HorizontalAlignment = xlCenter
    'now add formulae dependent on the freq
    Worksheet.[H4].NumberFormat = "mm/dd/yy"
    Worksheet.[H4].Font.ThemeColor = xlThemeColorDark1
    Worksheet.[H4].Font.TintAndShade = -0.249977111117893
    Worksheet.[H4].FormatConditions.Add Type:=xlExpression, Formula1:="=AND(WEEKDAY(H4)<>1,WEEKDAY(H3)<>7)"
    Worksheet.[H4].FormatConditions(Worksheet.[H4].FormatConditions.Count).SetFirstPriority
    With Worksheet.[H4].FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Worksheet.[H4].FormatConditions(1).StopIfTrue = False
    Worksheet.[H4].FormulaR1C1 = "=INDEX(R15C6:R41C12,MATCH(R4C7,R15C12:R41C12,1)+1,MATCH(""UL TITLE"",R14C6:R14C12,0))"
    Worksheet.[G5].NumberFormat = "mm/dd/yy"
    Worksheet.[G5].FormulaR1C1 = "=IF(WEEKDAY(R4C8)=1,R4C8+1,IF(WEEKDAY(R4C8)=7,R[-1]C[1]+2,R4C8))"
    Worksheet.[H5].FormulaR1C1 = "=IF(OR(WEEKDAY(R[-1]C)=1,WEEKDAY(R[-1]C)=7),""(Adjusted to next working day)"","""")"
    'gray out the note
    Worksheet.[H5].Font.ThemeColor = xlThemeColorDark1
    Worksheet.[H5].Font.TintAndShade = -0.249977111117893
    Worksheet.[G6].FormulaR1C1 = "=IF(R[-1]C>R[-4]C,NETWORKDAYS(R[-4]C,R[-1]C,HOLIDAYS[DATE])-1)"
    'format the table
    Set rng = Worksheet.[F14:L14]
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone
    Set rng = Worksheet.[F15:L41]
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    With Worksheet.[F42].Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
    End With
    xlApp.ScreenUpdating = True

    'create the chart
    xlApp.ScreenUpdating = False
    Worksheet.Range("H14:H41", "I14:I41").Select
    Set Chart = Worksheet.Shapes.AddChart2(, xlColumnClustered, Worksheet.[M14].Left + 10, Worksheet.[F14].Top, 525.44125984252, 318.735433070866).Chart
    'add title
    Chart.SetElement msoElementChartTitleAboveChart
    Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 20
    Chart.ChartTitle.Text = "QuickMonte - " & FormatDateTime(Now(), vbShortDate)
    Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoCTrue
    Chart.Axes(xlCategory).CategoryType = xlCategoryScale
    'add axis titles
    Chart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    Chart.Axes(xlCategory).AxisTitle.Caption = "Upper Bound"
    Chart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Caption = "Count"
    'todo: allow for earlier versions of excel v14
    'add black border
    With Chart.FullSeriesCollection(1).Format.line
      .Visible = msoTrue
      .ForeColor.ObjectThemeColor = msoThemeColorText1
      .ForeColor.TintAndShade = 0
      .ForeColor.Brightness = 0
      .Transparency = 0
    End With
    
    'add cumulative distrbution line
    Chart.SeriesCollection.NewSeries
    Chart.FullSeriesCollection(2).Name = "=cptQuickMonte_DATA!$L$14"
    Chart.FullSeriesCollection(2).Values = "=cptQuickMonte_DATA!$L$15:$L$41"
    Chart.FullSeriesCollection(2).ChartType = xlLineStacked
    'make columns fat (here to avoid the flicker from adding a second series)
    Chart.ChartGroups(1).GapWidth = 0
    Chart.FullSeriesCollection(2).AxisGroup = 2
    Chart.Axes(xlValue, xlSecondary).MaximumScale = 1
    Chart.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
    Chart.SetElement (msoElementSecondaryValueAxisTitleAdjacentToAxis)
    Chart.Axes(xlValue, xlSecondary).AxisTitle.Caption = "Cumulative Disribution"
    xlApp.ScreenUpdating = True
    'todo?: data validation on UID
    'todo: lock everything except the two input cells
    'todo: allow inspection in the form - read from excel in background
    'todo: draw probabilty extrapolation line
    'todo: add macro to workbook to redraw extrapolation line after changes?
    
    Worksheet.[G5].Select
    
    xlApp.Visible = True
    
  End If
  
  'todo: include costs;
  'todo: use number to capture percents? adjust for fixed dur/work;
  'todo: include option to output csv for mpm/propicer at confidence level
  
  cptQuickMonte_frm.lblStatus.Caption = "Complete"
  
exit_here:
  On Error Resume Next
  Set sProject = Nothing
  Set rng = Nothing
  Set Chart = Nothing
  Set CE = Nothing
  Set rst3p = Nothing
  Set Chart = Nothing
  Set ListObject = Nothing
  If Not xlApp Is Nothing Then xlApp.Visible = True
  If Application.ScreenUpdating = False Or Application.Calculation <> pjAutomatic Then cptSpeed False
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set rstSim = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMonte", "cptQuickMonte", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowQuickMonteFrm()
'objects
'strings
Dim strFieldName As String
'longs
Dim lngField As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'update form options
  With cptQuickMonte_frm
    For lngItem = 1 To 10
      lngField = FieldNameToFieldConstant("Duration" & lngItem, pjTask)
      .cboMin.AddItem
      .cboMin.List(lngItem - 1, 0) = lngField
      If Len(CustomFieldGetName(lngField)) > 0 Then
        strFieldName = "Duration" & lngItem & " (" & CustomFieldGetName(lngField) & ")"
      Else
        strFieldName = "Duration" & lngItem
      End If
      .cboMin.List(lngItem - 1, 1) = strFieldName
      .cboMax.AddItem
      .cboMax.List(lngItem - 1, 0) = lngField
      If Len(CustomFieldGetName(lngField)) > 0 Then
        strFieldName = "Duration" & lngItem & " (" & CustomFieldGetName(lngField) & ")"
      Else
        strFieldName = "Duration" & lngItem
      End If
      .cboMax.List(lngItem - 1, 1) = strFieldName
    Next lngItem
    .cboML.AddItem
    .cboML.List(0, 0) = FieldNameToFieldConstant("Remaining Duration")
    .cboML.List(0, 1) = "Remaining Duration"
    .cboML.Value = FieldNameToFieldConstant("Remaining Duration")
    .cboML.Enabled = False
    
    .cboDistribution.AddItem
    .cboDistribution.List(0, 0) = "Triangular"
    .cboDistribution.Value = "Triangular"
    .cboDistribution.Enabled = False
    
    'import saved settings if any  exist
    If Dir(cptDir & "settings\cpt-quickMonte-settings.adtg") <> vbNullString Then
      'todo: import saved settings
    End If
    
    .Show False
    
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMonte_bas", "cptShowQuickMonte", Err, Erl)
  Resume exit_here
End Sub

Sub quickSet()
Dim Task As Task
  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then
      If Task.Summary Then GoTo next_task
      If Task.ExternalTask Then GoTo next_task
      If Not Task.Active Then GoTo next_task
      If Task.RemainingDuration > 0 Then
        'task.SetField fieldnametofieldconstant("Duration1"),cstr(
      End If
    End If
next_task:
  Next Task
End Sub

Sub cptQuickPERT(lngMinField As Long, lngMaxField As Long, lngTargetTaskUID As Long)
'objects
Dim sProject As SubProject
Dim Worksheet As Object
Dim Workbook As Object
Dim xlApp As Object
Dim rst As ADODB.Recordset
Dim Task As Task
'strings
Dim strProjectName As String
Dim strHeader As String
Dim strCSV As String
Dim strMsg As String
'longs
Dim lngItem As Long
Dim lngFile As Long
Dim lngEVT As Long
Dim lngTask As Long
Dim lngTasks As Long
Dim lngML As Long
Dim lngMax As Long
Dim lngMin As Long
'integers
'doubles
Dim dblPERT As Double
'booleans
Dim blnDirty As Boolean
'variants
'dates
Dim dtPERT As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'I feel the need for speed
  'cptSpeed True
  
  'capture task count
  If ActiveProject.Subprojects.Count = 0 Then
    lngTasks = ActiveProject.Tasks.Count
  Else
    For Each sProject In ActiveProject.Subprojects
      lngTasks = lngTasks + sProject.SourceProject.Tasks.Count
    Next
  End If
  
  'prepare to capture existing values
  'todo: capture/zero/restore LOE and SM
  Set rst = CreateObject("ADODB.Recordset")
  rst.Fields.Append "UID", adBigInt
  rst.Fields.Append "MIN", adBigInt
  rst.Fields.Append "MAX", adBigInt
  rst.Fields.Append "ML", adBigInt
  rst.Fields.Append "PERT", adBigInt
  rst.Open
  
  'todo: user must set this
  'todo: why do we need this? LOE should be zero'd out
  lngEVT = FieldNameToFieldConstant("EVT")
  
  'capture current remaining duration, set PERT duration
  For Each Task In ActiveProject.Tasks
    If Not Task Is Nothing Then
      If Task.Summary Then GoTo next_task
      If Task.ExternalTask Then GoTo next_task
      If Not Task.Active Then GoTo next_task
      'todo: user must set these
      If Task.GetField(lngEVT) = "N/A" Then GoTo next_task
      If Task.GetField(lngEVT) = "A" Then GoTo next_task
      'todo: ignore tasks based on userform criteria e.g., LOE, Schedule Margin
      If Task.RemainingDuration > 0 Then
        lngMin = cptGetDuration(Task, lngMinField) / 480
        lngML = Task.RemainingDuration / 480
        lngMax = cptGetDuration(Task, lngMaxField) / 480
        dblPERT = (lngMin + (4 * lngML) + lngMax) / 6
        rst.AddNew Array(0, 1, 2, 3, 4), Array(Task.UniqueID, lngMin, lngMax, lngML, dblPERT)
        'Task.RemainingDuration = dblPERT
        Task.SetField pjTaskRemainingDuration, CStr(dblPERT)
        blnDirty = True
      End If
    End If
next_task:
    lngTask = lngTask + 1
    cptQuickMonte_frm.lblStatus.Caption = "Calculating PERT durations...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptQuickMonte_frm.lblProgress.Width = (lngTask / lngTasks) * cptQuickMonte_frm.lblStatus.Width
  Next Task

  'cleanup
  cptQuickMonte_frm.lblStatus.Caption = "Calculating PERT durations...(100%)"
  cptQuickMonte_frm.lblProgress.Width = cptQuickMonte_frm.lblStatus.Width
  
  'create csv
  strCSV = Environ("temp") & "\quickPERT.csv"
  If Dir(strCSV) <> vbNullString Then Kill strCSV
  lngFile = FreeFile
  Open strCSV For Output As #lngFile
  rst.MoveFirst
  For lngItem = 0 To rst.Fields.Count - 1
    strHeader = strHeader & rst.Fields(lngItem).Name & ","
  Next
  strHeader = Left(strHeader, Len(strHeader) - 1)
  Print #lngFile, strHeader
  Print #lngFile, rst.GetString(adClipString, , ",", vbCrLf, vbNullString)
  Close #lngFile

  'create map and import
'  strProjectName = ActiveProject.FullName
'  MapEdit Name:="QuickPERT", Create:=True, OverwriteExisting:=True, DataCategory:=0, CategoryEnabled:=True, TableName:="Task_Table1", FieldName:="Unique ID", ExternalFieldName:="UID", ExportFilter:="All Tasks", ImportMethod:=2, MergeKey:="Unique ID", headerRow:=True, AssignmentData:=False, TextDelimiter:=",", TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
'  MapEdit Name:="QuickPERT", DataCategory:=0, FieldName:="Remaining Duration", ExternalFieldName:="PERT"
'  FileOpenEx Name:=strCSV, ReadOnly:=True, Merge:=1, FormatID:="MSProject.CSV", Map:="QuickMonte"
  
  'calculate new network
  cptQuickMonte_frm.lblStatus.Caption = "Recalculating..."
  CalculateProject
  
  'capture PERT finish
  dtPERT = ActiveProject.Tasks.UniqueID(lngTargetTaskUID).Finish
'
'  FileCloseEx pjDoNotSave
'
'  FileOpenEx strProjectName
'
'  'create map and import
'  MapEdit Name:="QuickPERT", Create:=True, OverwriteExisting:=True, DataCategory:=0, CategoryEnabled:=True, TableName:="Task_Table1", FieldName:="Unique ID", ExternalFieldName:="UID", ExportFilter:="All Tasks", ImportMethod:=2, MergeKey:="Unique ID", headerRow:=True, AssignmentData:=False, TextDelimiter:=",", TextFileOrigin:=0, UseHtmlTemplate:=False, IncludeImage:=False
'  MapEdit Name:="QuickPERT", DataCategory:=0, FieldName:="Remaining Duration", ExternalFieldName:="ML"
'  FileOpenEx Name:=strCSV, ReadOnly:=True, Merge:=1, FormatID:="MSProject.CSV", Map:="QuickMonte"
'
'  GoTo skip_to_here
  
  'restore settings
  cptQuickMonte_frm.lblStatus.Caption = "Restoring durations..."
  rst.MoveFirst
  lngTask = 0
  lngTasks = rst.RecordCount
  Do While Not rst.EOF
    ActiveProject.Tasks.UniqueID(rst(0)).SetField pjTaskRemainingDuration, rst("ML")
    rst.MoveNext
    lngTask = lngTask + 1
    cptQuickMonte_frm.lblStatus.Caption = "Restoring durations...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptQuickMonte_frm.lblProgress.Width = (lngTask / lngTasks) * cptQuickMonte_frm.lblStatus.Width
  Loop
  
  'cleanup
  cptQuickMonte_frm.lblStatus.Caption = "Restoring durations...(100%)"
  cptQuickMonte_frm.lblProgress.Width = cptQuickMonte_frm.lblStatus.Width
  
skip_to_here:
  cptSpeed False
  
  'if we made it to this point then
  'original remaining durations have been restored
  blnDirty = False
  
  cptQuickMonte_frm.lblStatus.Caption = "Returning PERT result..."
  Set Task = ActiveProject.Tasks.UniqueID(lngTargetTaskUID)
  strMsg = "UID " & lngTargetTaskUID & ": " & Task.Name & vbCrLf & vbCrLf
  strMsg = strMsg & "Deterministic Finish: " & FormatDateTime(Task.Finish, vbShortDate) & vbCrLf
  strMsg = strMsg & "Estimated using PERT: " & FormatDateTime(dtPERT, vbShortDate) & vbCrLf & vbCrLf
  strMsg = strMsg & "Recommended Margin: " & Round(Application.DateDifference(Task.Finish, dtPERT, ActiveProject.Calendar) / (60 * ActiveProject.HoursPerDay), 0) & " days" & vbCrLf & vbCrLf
  strMsg = strMsg & "Would you like to review the durations used?"
  If MsgBox(strMsg, vbInformation + vbYesNo, "PERT Estimate") = vbYes Then
    Application.StatusBar = "Creating Excel Workbook..."
    Set xlApp = CreateObject("Excel.Application")
    Set Workbook = xlApp.Workbooks.Add
    Set Worksheet = Workbook.Sheets(1)
    Worksheet.Name = "PERT"
    Worksheet.[A1:E1] = Array("UID", "MIN", "MAX", "ML", "PERT")
    Worksheet.[A2].CopyFromRecordset rst
    xlApp.ActiveWindow.Zoom = 85
    Worksheet.[A2].AutoFilter
    Worksheet.Columns.AutoFit
    Worksheet.[A2].Select
    xlApp.ActiveWindow.FreezePanes = True
    xlApp.Visible = True
  End If

  cptQuickMonte_frm.lblStatus.Caption = "QuickPERT Complete"

exit_here:
  On Error Resume Next
  Set sProject = Nothing
  cptQuickMonte_frm.lblStatus.Caption = "Ready"
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  If blnDirty Then
    MsgBox "Durations not restored! Close without saving to avoid loss of information.", vbCritical + vbOKOnly, "Restore Process Failed"
  End If
  cptSpeed False
  If rst.State Then rst.Close
  Set rst = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMonte_bas", "cptQuickPERT", Err, Erl)
  If blnDirty Then MsgBox "Durations not restored! Close without saving to avoid loss of information.", vbCritical + vbOKOnly, "Restore Process Failed"
  Resume exit_here
End Sub

Function cptGetDuration(ByRef Task As Task, lngField As Long) As Long
  Select Case lngField
    Case pjTaskDuration1
      cptGetDuration = Task.Duration1
    Case pjTaskDuration2
      cptGetDuration = Task.Duration2
    Case pjTaskDuration3
      cptGetDuration = Task.Duration3
    Case pjTaskDuration4
      cptGetDuration = Task.Duration4
    Case pjTaskDuration5
      cptGetDuration = Task.Duration5
    Case pjTaskDuration6
      cptGetDuration = Task.Duration6
    Case pjTaskDuration7
      cptGetDuration = Task.Duration7
    Case pjTaskDuration8
      cptGetDuration = Task.Duration8
    Case pjTaskDuration9
      cptGetDuration = Task.Duration9
    Case pjTaskDuration10
      cptGetDuration = Task.Duration10
  End Select
End Function

Sub cptTestImport()
  
End Sub
