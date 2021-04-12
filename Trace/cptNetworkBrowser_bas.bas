Attribute VB_Name = "cptNetworkBrowser_bas"
'<cpt_version>v0.0.0</version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowNetworkBrowser_frm()
  If Not cptFilterExists("Marked") Then cptCreateFilter ("Marked")
  Call cptStartEvents
  Call cptShowPreds
  cptNetworkBrowser_frm.Show False
End Sub

Sub cptShowPreds()
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
  
  lngTasks = ActiveSelection.Tasks.Count
  
  With cptNetworkBrowser_frm
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
  With cptNetworkBrowser_frm.lboPredecessors
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
  
  With cptNetworkBrowser_frm.lboSuccessors
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
  If Err.Number <> 424 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Dependency Browser: Error"
  Resume exit_here
  
End Sub

Sub cptMarkSelected()
  'todo: separate network browser and make it cptMarkSelected(Optional blnRefilter as Boolean)
  Dim oTask As Task, oTasks As Tasks
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If Not oTasks Is Nothing Then
    For Each oTask In oTasks
      oTask.Marked = True
    Next oTask
  End If
  If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
    'todo: call cptFilterReapply
    cptSpeed True
    FilterApply "All Tasks"
    FilterApply "Marked"
    cptSpeed False
  Else
    'todo
  End If
  Set oTask = Nothing
  Set oTasks = Nothing
End Sub

Sub cptUnmarkSelected()
'todo: make cptMark(blnMark as Boolean)
'todo: separate network browser and make it cptUnmarkSelected(Optional blnRefilter as Boolean)
Dim Task As Task

  For Each Task In ActiveSelection.Tasks
    If Not Task Is Nothing Then Task.Marked = False
  Next Task
  
  If cptNetworkBrowser_frm.Visible Then
    'todo: from here down from network browser only
    ActiveWindow.TopPane.Activate
    FilterApply "Marked"
    If ActiveWindow.TopPane.View.Name <> "Network Diagram" Then
      SelectAll
      ActiveWindow.BottomPane.Activate
      ViewApply "Network Diagram"
    Else
      'todo: call cptFilterReapply
      cptSpeed True
      FilterApply "All Tasks"
      FilterApply "Marked"
      cptSpeed False
    End If
  End If
End Sub

Sub cptMarked()
  ActiveWindow.TopPane.Activate
  On Error Resume Next
  If Not FilterApply("Marked") Then
    FilterEdit "Marked", True, True, True, , , "Marked", , "equals", "Yes", , True, False
  End If
  FilterApply "Marked"
End Sub

Sub cptClearMarked()
Dim oTask As Task

  For Each oTask In ActiveProject.Tasks
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If Not oTask Is Nothing Then oTask.Marked = False
next_task:
  Next oTask
  ActiveProject.Tasks.UniqueID(0).Marked = False
  'todo: fix this
  If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
    cptSpeed True
    FilterApply "Active Tasks"
    FilterApply "Marked"
    cptSpeed False
  Else
    'todo: if lower pane
  End If
  Set oTask = Nothing

End Sub

Sub cptHistoryDoubleClick()
Dim lngTaskID As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTaskID = CLng(cptNetworkBrowser_frm.lboHistory.Value)
  WindowActivate TopPane:=True
  If IsNumeric(lngTaskID) Then
    On Error Resume Next
    If Not EditGoTo(lngTaskID, ActiveProject.Tasks(lngTaskID).Start) Then
      If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
        ActiveProject.Tasks(lngTaskID).Marked = True
        FilterApply "Marked"
        GoTo exit_here
      End If
      If MsgBox("Task is hidden - remove filters and show it?", vbQuestion + vbYesNo, "Confirm Apocalypse") = vbYes Then
        FilterClear
        OptionsViewEx displaysummarytasks:=True
        OutlineShowAllTasks
        EditGoTo lngTaskID, ActiveProject.Tasks(lngTaskID).Start
      End If
    End If
  End If
  
exit_here:
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptHistoryDoubleClick", Err, Erl)
  Resume exit_here
End Sub
