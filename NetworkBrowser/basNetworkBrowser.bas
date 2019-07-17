Attribute VB_Name = "basNetworkBrowser"
Sub ShowFrmPreds()
  Call ShowPreds
  frmNetworkBrowser.Show False
End Sub

Sub ShowPreds()
Dim Pred As Object, Succ As Object, numTasks As Integer, t As Task
  Dim Tasks As Tasks
  
  On Error Resume Next
  Set t = ActiveSelection.Tasks(1)
  If t Is Nothing Then GoTo exit_here
  
  On Error GoTo err_here
  
  numTasks = ActiveSelection.Tasks.Count
  
  With frmNetworkBrowser
    Select Case numTasks
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
  
  If ActiveCell.Task Is Nothing Then GoTo exit_here
  
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
    
    For Each Pred In ActiveCell.Task.TaskDependencies
      If Pred.From.ID <> ActiveCell.Task.ID Then
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
    For Each Succ In ActiveCell.Task.TaskDependencies
      If Succ.To.ID <> ActiveCell.Task.ID Then
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
  
  'Application.Sort "Start"
  
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

