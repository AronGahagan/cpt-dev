Attribute VB_Name = "cptNetworkBrowser_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub ShowFrmPreds()
  Call ShowPreds
  frmNetworkBrowser.show False
End Sub

Sub ShowPreds()
'objects
Dim Pred As Object, Succ As Object, Task As Task, Tasks As Tasks
'strings
'longs
Dim lngTasks As Long
'integers
'doubles
'booleans
'variants
'dates
  
  On Error Resume Next
  Set Task = ActiveSelection.Tasks(1)
  If Task Is Nothing Then GoTo exit_here
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTasks = ActiveSelection.Tasks.count
  
  With frmNetworkBrowser
    Select Case lngTasks
      Case Is < 1
        .lboPredecessors.Clear
        .lboPredecessors.ColumnCount = 1
        .lboPredecessors.AddItem "Please select a task."
        .lboSuccessors.Clear
        .lboSuccessors.Column = 1
        .lboSuccessors.AddItem "Please select a task."
        GoTo exit_here
      Case Is > 1
        .lboPredecessors.Clear
        .lboPredecessors.ColumnCount = 1
        .lboPredecessors.AddItem "Please select only one task."
        .lboSuccessors.Clear
        .lboSuccessors.ColumnCount = 1
        .lboSuccessors.AddItem "Please select only one task."
        GoTo exit_here
    End Select
  End With
    
  'only 1 is selected
  With frmNetworkBrowser.lboPredecessors
    .Clear
    .ColumnCount = 7
    .AddItem
    .Column(0, .ListCount - 1) = "ID"
    .Column(1, .ListCount - 1) = "UID"
    .Column(2, .ListCount - 1) = "Lag"
    .Column(3, .ListCount - 1) = "Finish"
    .Column(4, .ListCount - 1) = "Slack"
    .Column(5, .ListCount - 1) = "Task"
    .Column(6, .ListCount - 1) = "Critical"
    
    For Each Pred In Task.TaskDependencies
      If Pred.From.ID <> Task.ID Then
        .AddItem
        .Column(0, .ListCount - 1) = Pred.From.ID
        .Column(1, .ListCount - 1) = Pred.From.UniqueID
        .Column(2, .ListCount - 1) = Round(Pred.Lag / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(3, .ListCount - 1) = Format(Pred.From.Finish, "mm/dd/yy")
        .Column(4, .ListCount - 1) = Round(Pred.From.TotalSlack / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(5, .ListCount - 1) = IIf(Pred.From.Marked, "[m] ", "") & IIf(Len(Pred.From.Name) > 65, Left(Pred.From.Name, 65) & "... ", Pred.From.Name)
        .Column(6, .ListCount - 1) = IIf(Pred.From.Critical, "CRITICAL", "")
      End If
    Next
  End With
  
  With frmNetworkBrowser.lboSuccessors
    .Clear
    .ColumnCount = 7
    .AddItem
    .Column(0, .ListCount - 1) = "ID"
    .Column(1, .ListCount - 1) = "UID"
    .Column(2, .ListCount - 1) = "Lag"
    .Column(3, .ListCount - 1) = "Start"
    .Column(4, .ListCount - 1) = "Slack"
    .Column(5, .ListCount - 1) = "Task"
    .Column(6, .ListCount - 1) = "Critical"
    For Each Succ In Task.TaskDependencies
      If Succ.To.ID <> Task.ID Then
        .AddItem
        .Column(0, .ListCount - 1) = Succ.To.ID
        .Column(1, .ListCount - 1) = Succ.To.UniqueID
        .Column(2, .ListCount - 1) = Round(Succ.Lag / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(3, .ListCount - 1) = Format(Succ.To.Finish, "mm/dd/yy")
        .Column(4, .ListCount - 1) = Round(Succ.To.TotalSlack / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(5, .ListCount - 1) = IIf(Succ.To.Marked, "[m] ", "") & IIf(Len(Succ.To.Name) > 65, Left(Succ.To.Name, 65) & "... ", Succ.To.Name)
        .Column(6, .ListCount - 1) = IIf(Succ.To.Critical, "CRITICAL", "")
      End If
    Next
  End With
  
exit_here:
  Exit Sub
err_here:
  If err.Number <> 424 Then MsgBox err.Number & ": " & err.Description, vbExclamation, "Dependency Browser: Error"
  Resume exit_here
  
End Sub

Sub UnmarkSelected()
Dim Task As Task

  For Each Task In ActiveSelection.Tasks
    Task.Marked = False
  Next Task
  ActiveWindow.TopPane.Activate
  FilterApply "Marked"
  SelectAll
  ActiveWindow.BottomPane.Activate
  ViewApply "Network Diagram"
End Sub

Sub cptHistoryDoubleClick()
Dim lngTaskID As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTaskID = CLng(Me.lboHistory.Value)
  WindowActivate TopPane:=True
  If IsNumeric(lngTaskID) Then EditGoTo lngTaskID

exit_here:
  Exit Sub
err_here:
  If err.Number = 1101 Then
    Call RemoveFilters(lngTaskID)
  Else
    Call cptHandleErr("cptNetworkBrowser_bas", "cptHistoryDoubleClick", err, Erl)
  End If
  Resume exit_here
End Sub

Sub RemoveFilters(lngTaskID As Long)
Dim msg As String
  msg = "ID " & lngTaskID & " is not currently visible." & vbCrLf & vbCrLf & "Remove filters and go to " & lngTaskID & "?"
  If MsgBox(msg, vbExclamation + vbYesNo, "Hidden") = vbYes Then
    Application.FilterApply "All Tasks"
    If ActiveProject.AutoFilter Then AutoFilter
    Sort Key1:="ID"
    On Error Resume Next
    If err.Number = 1100 Then SummaryTasksShow
    Application.OutlineShowAllTasks
    EditGoTo lngTaskID
  End If
End Sub
