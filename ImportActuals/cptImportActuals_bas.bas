Attribute VB_Name = "cptImportActuals_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptImportActuals()
'objects
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim Assignment As Assignment
Dim Resource As Resource
Dim Task As Task
'strings
Dim strResource As String
Dim strWPCN As String
'longs
'integers
'doubles
Dim dblMatl As Double
Dim dblHours As Double
'booleans
'variants
'dates
Dim dtWeek As Date
Dim dtStart As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'user provides import file: indicate current or cumulative
  
  'validate field mapping
  'wpcn,resource,labor(dbl),matl,week

  'summarize

  'for each week in importfile
    'validate it's a friday
    'validate it's an existing resource
    'clear out actuals for current week
    'find existing actuals task or create
    'add assignments.actualwork
  'next

  cptSpeed True

  strWPCN = "06005-123-ME"
  On Error Resume Next
  Set Task = ActiveProject.Tasks(strWPCN & " - ACTUALS")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Task Is Nothing Then
    Set Task = ActiveProject.Tasks.Add(strWPCN & " - ACTUALS")
  End If
  If Task.Type <> pjFixedDuration Then Task.Type = pjFixedDuration
  If Task.EffortDriven Then Task.EffortDriven = False
  If Task.Estimated Then Task.Estimated = False

  dblHours = 0
  dblMatl = 100
  
  strResource = "TestChamber"
  On Error Resume Next
  Set Resource = ActiveProject.Resources(strResource)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Resource Is Nothing Then
    'todo: prompt user
    Set Resource = ActiveProject.Resources.Add(strResource)
    If dblHours = 0 And dblMatl > 0 Then
      Resource.Type = pjResourceTypeMaterial
      Resource.StandardRate = 1
    End If
    'todo: resource.type=pjMaterial
  End If
  
  For Each Assignment In Task.Assignments
    If Assignment.ResourceName = strResource Then Exit For
  Next Assignment
  If Assignment Is Nothing Then
    Set Assignment = Task.Assignments.Add(Task.ID, Resource.ID, 1)
  End If
  
  If Assignment.RemainingWork > 0 Then Task.RemainingWork = 0
  
  dtWeek = #12/14/2019#
  'todo: make it friday
  If Resource.Type = pjResourceTypeWork Then
    Set TSVS = Assignment.TimeScaleData(dtWeek, dtWeek, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
    TSVS(1).Value = dblHours * 60
  Else
    Set TSVS = Assignment.TimeScaleData(dtWeek, dtWeek, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
    TSVS(1).Value = dblMatl
  End If
  
  'todo: flag it somehow
  'todo: optionally import into second file and use master to export reports?
  'If Task.Active Then Task.Active = False
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Set TSV = Nothing
  Set TSVS = Nothing
  Set Assignment = Nothing
  Set Resource = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptImportActuals", Err, Erl)
  Resume exit_here
End Sub
