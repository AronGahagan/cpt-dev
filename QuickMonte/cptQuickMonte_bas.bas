Attribute VB_Name = "cptQuickMonte_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptQuickMonte()
'objects
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
Dim lngDays As Long
Dim lngX As Long
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
Dim dblP As Double
Dim dblCDF_ML As Double
'booleans
Dim blnChangeHighlighting As Boolean
Dim blnFail As Boolean
'variants
Dim vbResponse As Variant
'dates

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
  
  'speed up processing and prevent screen flicker
  cptSpeed True
  
  'get three-point fields
  'todo: user must set these
  lngMin = FieldNameToFieldConstant("MinDuration")
  lngMax = FieldNameToFieldConstant("MaxDuration")
  
  'todo: capture, remove, restore deadlines and constraints?
  
  'capture three points
  Application.StatusBar = "Capturing three points..."
  Set rst3p = CreateObject("ADODB.Recordset")
  rst3p.Fields.Append "UID", adBigInt
  rst3p.Fields.Append "MIN", adBigInt
  rst3p.Fields.Append "ML", adBigInt
  rst3p.Fields.Append "MAX", adBigInt
  rst3p.Fields.Append "SM", adBoolean
  rst3p.Open
  For Each Task In ActiveProject.Tasks
    If Task Is Nothing Then GoTo next_task0
    'catch and ignore schedule margin > 0
    If InStr(Task.Name, "SCHEDULE MARGIN") > 0 Then
      If Task.RemainingDuration > 0 Then
        Application.StatusBar = "SCHEDULE MARGIN FOUND"
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
    Else
      Application.StatusBar = "Capturing three points..."
      'todo: what if user selects a text or number field?
      'convert custom duration text to long
      If InStr(Task.GetField(lngMin), "d") = 0 Then
        lngMinDur = cptGetLngFromDurText(Task.GetField(lngMin))
      Else
        lngMinDur = cptRegEx(Task.GetField(lngMin), "[0-9]*") * ActiveProject.HoursPerDay * 60
      End If
      'convert custom duration text to long
      If InStr(Task.GetField(lngMax), "d") = 0 Then
        lngMaxDur = cptGetLngFromDurText(Task.GetField(lngMax))
      Else
        lngMaxDur = cptRegEx(Task.GetField(lngMax), "[0-9]*") * ActiveProject.HoursPerDay * 60
      End If
      lngMLDur = Task.RemainingDuration
      rst3p.AddNew Array(0, 1, 2, 3, 4), Array(Task.UniqueID, lngMinDur, lngMLDur, lngMaxDur, False)
    End If
next_task0:
  Next Task
  
  'in case schedule margin was removed:
  CalculateProject 'once
  
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
  For lngIteration = 1 To lngIterations
    'simulate project
    For Each Task In ActiveProject.Tasks
      If Task.RemainingDuration = 0 Then GoTo next_task1
      rst3p.MoveFirst
      rst3p.Find "UID=" & Task.UniqueID, , adSearchForward
      If Not rst3p.EOF Then
        'skip schedule margin tasks
        If rst3p("SM") = True Then GoTo next_task1
        lngMinDur = rst3p("MIN")
        lngMLDur = rst3p("ML")
        lngMaxDur = rst3p("MAX")
        blnFail = False
        'validate three points
        If lngMinDur >= lngMLDur Or lngMLDur >= lngMaxDur Then
          MsgBox "Task UID '" & Task.Name & "' has invalid three point estimates.", vbCritical + vbOKOnly, "Error"
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
        Task.RemainingDuration = lngX
      Else
        MsgBox "Task information not found for UID " & Task.UniqueID & "!" & vbCrLf & vbCrLf & "Process will terminate.", vbCritical + vbOKOnly, "ERROR"
        blnFail = True
        GoTo restore_durations
      End If
next_task1:
    Next Task
        
    CalculateProject
    
    'capture simulation
    For Each Task In ActiveProject.Tasks
      If Task Is Nothing Then GoTo next_task2
      rstSim.AddNew Array(0, 1, 2, 3), Array(lngIteration, Task.UniqueID, Task.RemainingDuration, Task.Finish)
next_task2:
    Next Task
    Application.StatusBar = "Running Simulation " & lngIteration & " of " & lngIterations & "...(" & Format(lngIteration / lngIterations, "0%") & ")"
    DoEvents
  Next lngIteration
  'rstSim.Update 'not sure we need this
  
restore_durations:
  Application.StatusBar = "Restoring remaining durations..."
  rst3p.MoveFirst
  Do While Not rst3p.EOF
    ActiveProject.Tasks.UniqueID(rst3p("UID")).RemainingDuration = CLng(rst3p("ML"))
    rst3p.MoveNext
  Loop
  
  'capture enable highlighting setting and turn off
  blnChangeHighlighting = Application.EnableChangeHighlighting
  Application.EnableChangeHighlighting = False
  'calculate project - goal is to prevent screen from changing at all
  'todo: prevent gantt changes
  CalculateProject
  cptSpeed False
  'restore highlighting settings
  Application.EnableChangeHighlighting = blnChangeHighlighting
  Application.StatusBar = "Complete"
  DoEvents
  If blnFail Then GoTo exit_here
  
  If MsgBox("Simluation Complete" & vbCrLf & vbCrLf & "Create Report?", vbInformation + vbYesNo, "QuickMonte") = vbYes Then
  
    'export results
    Application.StatusBar = "Creating Report..."
    Set xlApp = CreateObject("Excel.Application")
    Set Workbook = xlApp.Workbooks.Add
    Set Worksheet = Workbook.Sheets(1)
    Worksheet.Name = "cptQuickMonte_DATA"
    Worksheet.[A1:D1] = Array("ITERATION", "UID", "REMAINING DURATION", "FINISH")
    Worksheet.[A2].CopyFromRecordset rstSim
    rstSim.Close
    xlApp.ActiveWindow.Zoom = 85
    Worksheet.Columns.AutoFit
    Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)))
    ListObject.Name = "QuickMonte"
    
    'add informational column
    Worksheet.[F1:F12] = WorksheetFunction.Transpose(Array("UID", "Deterministic:", "Iterations:", "Confidence:", "Confidence Date:", "Margin Rec.:", "Min:", "Max:", "Mean:", "Range:", "Bin Count:", "Bin Size:"))
    Worksheet.[F1:F12].Font.Bold = True
    Worksheet.Columns("F:F").AutoFit
    'add freq chart titles
    Worksheet.[F14:L14] = Array("LL", "UL", "UL TITLE", "Freq", "Cum Freq", "Freq %", "Cum %")
    Worksheet.[F14:L14].Font.Bold = True
    Worksheet.[G1:G12].HorizontalAlignment = xlCenter
    Worksheet.[G1:G4].Style = "Input"
    Worksheet.[G5:G6].Style = "Calculation"
    Worksheet.[G1].Value = 14 'todo: select a uid
    Worksheet.[G2].Value = ActiveProject.Tasks.UniqueID(14).Finish  'todo: get deterministic finish of selected uid
    Worksheet.[G2].NumberFormat = "mm/dd/yy"
    Worksheet.[G3].Value = lngIterations
    Worksheet.[G4].Value = 0.9 'todo: get this value from user form
    Worksheet.[G7].FormulaR1C1 = "=ROUND(MINIFS(QuickMonte[FINISH],QuickMonte[UID],R1C7),0)"
    Worksheet.[H7].FormulaR1C1 = "=RC[-1]"
    Worksheet.[H7].NumberFormat = "mm/dd/yy"
    Worksheet.[G8].FormulaR1C1 = "=ROUND(MAXIFS(QuickMonte[FINISH],QuickMonte[UID],R1C7),0)"
    Worksheet.[H8].FormulaR1C1 = "=RC[-1]"
    Worksheet.[H8].NumberFormat = "mm/dd/yy"
    Worksheet.[G9].FormulaR1C1 = "=ROUND(AVERAGEIFS(QuickMonte[FINISH],QuickMonte[UID],R[-8]C),0)"
    Worksheet.[H9].FormulaR1C1 = "=RC[-1]"
    Worksheet.[H9].NumberFormat = "mm/dd/yy"
    Worksheet.[G10].FormulaR1C1 = "=DAYS(R[-2]C,R[-3]C)"
    Worksheet.[G11].Value = 25
    Worksheet.[G12].FormulaR1C1 = "=R10C7/R11C7"
    
    'capture exceptions in [Q2]
    Worksheet.[Q1:R1].Merge
    Worksheet.[Q1].Value2 = "EXCEPTIONS"
    Worksheet.[Q1].HorizontalAlignment = xlCenter
    Worksheet.[Q2:R2] = Array("NAME", "DATE")
    For Each CE In ActiveProject.Calendar.Exceptions
      For lngDays = 0 To CE.Occurrences - 1
        Worksheet.Cells(Worksheet.[Q1].End(xlDown).Row + 1, 17) = CE.Name
        Worksheet.Cells(Worksheet.[Q1].End(xlDown).Row, 18) = DateAdd("d", lngDays, CE.Start)
      Next lngDays
    Next CE
    Worksheet.Range(Worksheet.[R3], Worksheet.[R3].End(xlDown)).NumberFormat = "mm/dd/yyyy"
    Worksheet.Columns("Q:R").AutoFit
    Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[Q2].End(xlToRight), Worksheet.[Q2].End(xlDown)))
    ListObject.Name = "HOLIDAYS"
    
    'create frequency distribution chart
    Worksheet.[F15].FormulaR1C1 = "=R[-8]C[1]-R12C7"
    Worksheet.[F16:F42].FormulaR1C1 = "=R[-1]C+R12C7"
    Worksheet.[F15:F42].NumberFormat = "mm/dd/yy"
    Worksheet.[G15:G41].FormulaR1C1 = "=R[1]C[-1]-0.0001"
    Worksheet.[H15:H41].FormulaR1C1 = "=ROUND(RC[-1],0)"
    Worksheet.[I15:I41].FormulaArray = "=FREQUENCY(IF(QuickMonte[UID]=$G$1,QuickMonte[FINISH]),$G$15:$G$41)"
    Worksheet.[J15].FormulaR1C1 = "=RC[-1]"
    Worksheet.[J16:J41].FormulaR1C1 = "=R[-1]C+RC[-1]"
    Worksheet.[K15:K41].FormulaR1C1 = "=RC[-2]/R3C7"
    Worksheet.[L15].FormulaR1C1 = "=RC[-1]"
    Worksheet.[L16:L41].FormulaR1C1 = "=R[-1]C+RC[-1]"
    
    'now add formulae dependent on the freq
    Worksheet.[G5].FormulaR1C1 = "=INDEX(R15C6:R41C12,MATCH(R4C7,R15C12:R41C12,1)+1,MATCH(""UL TITLE"",R14C6:R14C12,0))"
    Worksheet.[G5].NumberFormat = "mm/dd/yy"
    Worksheet.[G6].FormulaR1C1 = "=IF(R[-1]C>R[-4]C,NETWORKDAYS(R[-4]C,R[-1]C,HOLIDAYS[DATE]))"

    'center the distribution table
    Worksheet.Range(Worksheet.[F14].End(xlToRight), Worksheet.[F14].End(xlDown)).HorizontalAlignment = xlCenter
    
    'create the chart
    
    
    'todo: lock everything except the two input cells
    
    'credit for the filtered frequency formula goes to ExcelJet:
    'https://exceljet.net/formula/count-unique-text-values-with-criteria
    
    'todo: allow inspection in the form - read from excel in background
    
    xlApp.Visible = True
    
  End If
  
  'todo: include costs;
  'todo: use number to capture percents? adjust for fixed dur/work;
  'todo: include option to output csv for mpm/propicer at confidence level
  
  Application.StatusBar = "Complete"
  
exit_here:
  On Error Resume Next
  Set CE = Nothing
  Set rst3p = Nothing
  Set Chart = Nothing
  Application.StatusBar = ""
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

Function cptGetLngFromDurText(strDuration As String)
'objects
'strings
'longs
Dim lngFactor As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'determine format
  If Len(cptRegEx(strDuration, "y|yr|year")) > 0 Then
    'not sure this is even possible...
  ElseIf Len(cptRegEx(strDuration, "mo|mon|month")) > 0 Then
    'multiply by days/mo * hrs/day * 60
    cptGetLngFromDurText = cptRegEx(strDuration, "[0-9]*") * ActiveProject.DaysPerMonth * ActiveProject.HoursPerDay * 60
  ElseIf Len(cptRegEx(strDuration, "w|wk|week")) > 0 Then
    'multiply by hrs/wk * 60 = minutes
    cptGetLngFromDurText = cptRegEx(strDuration, "[0-9]*") * ActiveProject.HoursPerWeek * 60
  ElseIf Len(cptRegEx(strDuration, "h|hr|hour")) > 0 Then
    'multiply hours by 60 min/hr = minutes
    cptGetLngFromDurText = cptRegEx(strDuration, "[0-9]*") * 60
  ElseIf Len(cptRegEx(strDuration, "m|min|minute")) > 0 Then
    'no conversion necessary = minutes
    cptGetLngFromDurText = cptRegEx(strDuration, "[0-9]*")
  End If
  'todo: use activeproject.hoursperday instead of 480
  
exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptQuickMonte_bas", "cptGetLngFromDurText", Err, Err)
  Resume exit_here
End Function
