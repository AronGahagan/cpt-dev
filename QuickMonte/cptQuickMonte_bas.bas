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
Dim rst3p As ADODB.Recordset
Dim rstSim As ADODB.Recordset
Dim arrDurations As SortedList 'Object
Dim Task As Task
'strings
'longs
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
Dim blnFail As Boolean
'variants
Dim vbResponse As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  vbResponse = InputBox("How many iterations?", "QuickMonte", 1000)
  If StrPtr(vbResponse) = 0 Then
    'user hit cancel
    GoTo exit_here
  ElseIf vbResponse = vbNullString Then
    'user entered null value
    GoTo exit_here
  Else
    lngIterations = cptRegEx(CStr(vbResponse), "[0-9].*")
    If lngIterations = 0 Then GoTo exit_here
  End If
  
  cptSpeed True
  
'  'capture three points - method 1
'  Set arrDurations = CreateObject("System.Collections.SortedList")
'  For Each Task In ActiveProject.Tasks
'    If Task Is Nothing Then GoTo next_task0
'    arrDurations.Add Task.UniqueID, Task.RemainingDuration
'next_task0:
'  Next Task
  
  'get three-point fields
  'todo: user must set these
  lngMin = FieldNameToFieldConstant("MinDuration")
  lngMax = FieldNameToFieldConstant("MaxDuration")
  
  'capture three points - method 2
  Set rst3p = CreateObject("ADODB.Recordset")
  rst3p.Fields.Append "UID", adInteger
  rst3p.Fields.Append "MIN", adInteger
  rst3p.Fields.Append "ML", adInteger
  rst3p.Fields.Append "MAX", adInteger
  rst3p.Open
  For Each Task In ActiveProject.Tasks
    If Task Is Nothing Then GoTo next_task0
    lngMinDur = cptRegEx(Task.GetField(lngMin), "[0-9].") * 480
    lngMaxDur = cptRegEx(Task.GetField(lngMax), "[0-9].") * 480
    rst3p.AddNew Array(0, 1, 2, 3), Array(Task.UniqueID, lngMinDur, Task.RemainingDuration, lngMaxDur)
    rst3p.Update
next_task0:
  Next Task
  
  'prepare to capture simulation results
  Set rstSim = CreateObject("ADODB.Recordset")
  rstSim.Fields.Append "ITERATION", adInteger
  rstSim.Fields.Append "UID", adInteger
  rstSim.Fields.Append "R_DUR", adInteger
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
'        lngMinDur = cptRegEx(Task.GetField(lngMin), "[0-9].") * 480
'        lngMaxDur = cptRegEx(Task.GetField(lngMax), "[0-9].") * 480
'        lngMLDur = arrDurations.Item(Task.UniqueID)
        lngMinDur = rst3p("MIN")
        lngMLDur = rst3p("ML")
        lngMaxDur = rst3p("MAX")
        'todo: if NOT min < ml < max
        blnFail = False
        If Not (lngMinDur < lngMLDur) And (lngMLDur < lngMaxDur) Then
          MsgBox "Task UID '" & Task.Name & "' has invalid three point estimates.", vbCritical + vbOKOnly, "Error"
          blnFail = True
          GoTo restore_durations
        End If
        'determine CDF of ML value
        dblCDF_ML = (lngMLDur - lngMinDur) / (lngMaxDur - lngMLDur)
        'get random probability
        dblP = Math.Rnd
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
        
    'calculate project
    CalculateProject
    
    'capture simulation
    For Each Task In ActiveProject.Tasks
      If Task Is Nothing Then GoTo next_task2
      rstSim.AddNew Array(0, 1, 2, 3), Array(lngIteration, Task.UniqueID, Task.RemainingDuration, Task.Finish)
      'rstSim.Update
next_task2:
    Next Task
    Application.StatusBar = "Running Simulation " & lngIteration & " of " & lngIterations & "...(" & Format(lngIteration / lngIterations, "0%") & ")"
    DoEvents
  Next lngIteration
  rstSim.Update
  
restore_durations:
  'restore most likely durations
  'todo: change this to rst3p
  rst3p.MoveFirst
  Do While Not rst3p.EOF
    ActiveProject.Tasks.UniqueID(rst3p("UID")).RemainingDuration = CLng(rst3p("ML"))
    rst3p.MoveNext
  Loop
'  For lngItem = 0 To arrDurations.Count - 1
'    Set Task = ActiveProject.Tasks.UniqueID(arrDurations.getKey(lngItem))
'    Task.RemainingDuration = arrDurations.getValueList()(lngItem)
'  Next
  
  cptSpeed False
  
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
    xlApp.Visible = True
    xlApp.ActiveWindow.Zoom = 85
    Worksheet.Columns.AutoFit
    Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)))
    ListObject.Name = "QuickMonte"
    
    'todo: create workbook - lock everything except the two input cells
    
    'todo: allow inspection in the form - read from excel in background
    
  End If
    
  Application.StatusBar = "Complete"
  
exit_here:
  On Error Resume Next
  Set rst3p = Nothing
  Set Chart = Nothing
  Set PivotTable = Nothing
  Application.StatusBar = ""
  Set ListObject = Nothing
  If Not xlApp Is Nothing Then xlApp.Visible = True
  cptSpeed False
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set rstSim = Nothing
  Set arrDurations = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptQuickMonte", "cptQuickMonte", Err, Erl)
  Resume exit_here
End Sub
