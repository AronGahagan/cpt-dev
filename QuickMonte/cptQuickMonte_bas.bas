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
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngIterations = 500
  
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
  rst.Fields.Append "R_DUR", adInteger
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
      lngMLDur = arrDurations.Item(Task.UniqueID)
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
next_task1:
    Next Task
    
    'todo: create array of [iterations]
    
    'calculate project
    CalculateProject
    
    'capture simulation
    For Each Task In ActiveProject.Tasks
      If Task Is Nothing Then GoTo next_task2
      rst.AddNew Array(0, 1, 2, 3), Array(lngIteration, Task.UniqueID, Task.RemainingDuration, Task.Finish)
      rst.Update
next_task2:
    Next Task
    Application.StatusBar = "Running Simulation " & lngIteration & " of " & lngIterations & "...(" & Format(lngIteration / lngIterations, "0%") & ")"
    DoEvents
  Next lngIteration
  
  'restore most likely durations
  For lngItem = 0 To arrDurations.Count - 1
    Set Task = ActiveProject.Tasks.UniqueID(arrDurations.getKey(lngItem))
    Task.RemainingDuration = arrDurations.getValueList()(lngItem)
  Next
  
  cptSpeed False
  
  MsgBox "Simluation Complete", vbInformation + vbOKOnly, "QuickMonte"
  
  'export results
  Application.StatusBar = "Creating Report..."
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = "cptQuickMonte_DATA"
  Worksheet.[A1:D1] = Array("ITERATION", "UID", "REMAINING DURATION", "FINISH")
  Worksheet.[A2].CopyFromRecordset rst
  rst.Close
  xlApp.Visible = True
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.Columns.AutoFit
  Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)))
  ListObject.Name = "QuickMonte"
  
  'todo: add task name
  
  'create chart
    
    
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
