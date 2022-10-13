Attribute VB_Name = "cptNetworkBrowser_bas"
'<cpt_version>v1.1.0</cpt_version>
Option Explicit
Public oSubMap As Scripting.Dictionary

Sub cptShowNetworkBrowser_frm()
  'objects
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptFilterExists("Marked") Then cptCreateFilter ("Marked")
  
  Call cptStartEvents
  Call cptShowPreds
  With cptNetworkBrowser_frm
    .Caption = "Network Browser (" & cptGetVersion("cptNetworkBrowser_frm") & ")"
    .tglTrace = False
    .tglTrace.Caption = "Jump"
    .lboPredecessors.MultiSelect = fmMultiSelectSingle
    .lboSuccessors.MultiSelect = fmMultiSelectSingle
    .Show False
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptShowNetworkBrowser_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowPreds()
'objects
Dim oTaskDependencies As TaskDependencies
Dim oSubproject As Subproject
Dim oLink As TaskDependency, oTask As Task
'strings
Dim strProject As String
'longs
Dim lngLinkUID As Long
Dim lngItem As Long
Dim lngFactor As Long
Dim lngTasks As Long
'integers
'doubles
'booleans
Dim blnSubprojects As Boolean
'variants
Dim vControl As Variant
'dates
  
  On Error Resume Next
  Set oTask = ActiveSelection.Tasks(1)
  If oTask Is Nothing Then GoTo exit_here
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTasks = ActiveSelection.Tasks.Count
  'determine if there are subprojects loaded (this affects displayed UIDs)
  blnSubprojects = ActiveProject.Subprojects.Count > 0
  
  If blnSubprojects Then
    If oSubMap Is Nothing Then
      Set oSubMap = CreateObject("Scripting.Dictionary")
    Else
      oSubMap.RemoveAll
    End If
    For Each oSubproject In ActiveProject.Subprojects
      If InStr(oSubproject.Path, "\") > 0 Then 'offline
        oSubMap.Add Replace(Dir(oSubproject.Path), ".mpp", ""), 0
      ElseIf InStr(oSubproject.Path, "<>") > 0 Then 'online
        oSubMap.Add oSubproject.Path, 0
      End If
    Next oSubproject
    For Each oTask In ActiveProject.Tasks
      If oSubMap.Exists(oTask.Project) Then
        If oSubMap(oTask.Project) > 0 Then GoTo next_mapping_task
        If Not oTask.Summary Then
          oSubMap.Item(oTask.Project) = CLng(oTask.UniqueID / 4194304)
        End If
      End If
next_mapping_task:
    Next oTask
  End If
  
  Set oTask = ActiveSelection.Tasks(1)
  
  With cptNetworkBrowser_frm
    Select Case lngTasks
      Case Is < 1
        .lboCurrent.Clear
        .lboPredecessors.Clear
        .lboPredecessors.ColumnCount = 1
        .lboPredecessors.AddItem "Please select a task."
        .lboSuccessors.Clear
        .lboSuccessors.ColumnCount = 1
        .lboSuccessors.AddItem "Please select a task."
        GoTo exit_here
      Case Is > 1
        .lboCurrent.Clear
        .lboPredecessors.Clear
        .lboPredecessors.ColumnCount = 1
        .lboPredecessors.AddItem "Please select only one task."
        .lboSuccessors.Clear
        .lboSuccessors.ColumnCount = 1
        .lboSuccessors.AddItem "Please select only one task."
        GoTo exit_here
    End Select
    If .tglTrace Then
      .tglTrace.Caption = "Trace"
      .lboPredecessors.MultiSelect = fmMultiSelectMulti
      .lboPredecessors.MultiSelect = fmMultiSelectMulti
    Else
      .tglTrace.Caption = "Jump"
      .lboSuccessors.MultiSelect = fmMultiSelectSingle
      .lboSuccessors.MultiSelect = fmMultiSelectSingle
    End If
    With .lboCurrent
      .Clear
      .ColumnCount = 4
      .AddItem
      If blnSubprojects Then
        .ColumnWidths = "50 pt;35 pt;24.95 pt"
      Else
        .ColumnWidths = "24.95 pt;0 pt;24.95 pt"
      End If
      .Column(0, .ListCount - 1) = oTask.UniqueID
      .Column(1, .ListCount - 1) = oTask.UniqueID Mod 4194304
      .Column(2, .ListCount - 1) = oTask.ID
      .Column(3, .ListCount - 1) = IIf(oTask.Marked, "[m] ", "") & oTask.Name
    End With
  End With
    
  'only 1 is selected
  On Error Resume Next
  Set oTaskDependencies = oTask.TaskDependencies
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oTaskDependencies Is Nothing Then
    cptNetworkBrowser_frm.lboPredecessors.Clear
    cptNetworkBrowser_frm.lboSuccessors.Clear
    GoTo exit_here
  End If
    
  'reset both lbos once in an array here
  For Each vControl In Array("lboPredecessors", "lboSuccessors")
    With cptNetworkBrowser_frm.Controls(vControl)
      .Clear
      .ColumnCount = 9
      .AddItem
      If blnSubprojects Then
        .ColumnWidths = "50 pt;35 pt;24.95 pt;24.95 pt;24.95 pt;55 pt;35 pt;225 pt;35 pt"
        .Column(0, .ListCount - 1) = "UID[M]"
        .Column(1, .ListCount - 1) = "UID[S]"
      Else
        .ColumnWidths = "35 pt;0 pt;24.95 pt;24.95 pt;24.95 pt;55 pt;35 pt;225 pt;35 pt"
        .Column(0, .ListCount - 1) = "UID"
      End If
      .Column(2, .ListCount - 1) = "ID"
      .Column(3, .ListCount - 1) = "Type"
      .Column(4, .ListCount - 1) = "Lag"
      .Column(5, .ListCount - 1) = IIf(vControl = "lboPredecessors", "Finish", "Start")
      .Column(6, .ListCount - 1) = "Slack"
      .Column(7, .ListCount - 1) = "Task"
      .Column(8, .ListCount - 1) = "Critical"
    End With
  Next
  
  'capture list of preds with valid native UIDs
  For Each oLink In oTask.TaskDependencies
    'limit to only predecessors
    If oLink.To.Guid = oTask.Guid Then 'it's a predecessor to selected task
      'handle external tasks
      If blnSubprojects And oLink.From.ExternalTask Then
        'fix the returned UID
        lngLinkUID = oLink.From.GetField(185073906) Mod 4194304
        strProject = oLink.From.Project
        If InStr(oLink.From.Project, "\") > 0 Then
          strProject = Replace(strProject, ".mpp", "")
          strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
        End If
        lngFactor = oSubMap(strProject)
        lngLinkUID = (lngFactor * 4194304) + lngLinkUID
      Else
        If blnSubprojects Then
          lngFactor = Round(oTask / 4194304, 0)
          lngLinkUID = (lngFactor * 4194304) + oLink.From.UniqueID
        Else
          lngLinkUID = oLink.From.UniqueID
        End If
      End If
      With cptNetworkBrowser_frm.lboPredecessors
        .AddItem
        .Column(0, .ListCount - 1) = lngLinkUID
        .Column(1, .ListCount - 1) = lngLinkUID Mod 4194304 & IIf(oLink.From.ExternalTask, "x", "")
        'todo: sort this out - what do you want to show for external links in a standalone project file?
        If blnSubprojects And oLink.From.ExternalTask Then
          .Column(2, .ListCount - 1) = ActiveProject.Tasks.UniqueID(lngLinkUID).ID
          .Column(7, .ListCount - 1) = IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.From.Name) > 65, Left(oLink.From.Name, 65) & "... ", oLink.From.Name)
        ElseIf Not blnSubprojects And oLink.From.ExternalTask Then
          .Column(2, .ListCount - 1) = oLink.From.ID
          .Column(7, .ListCount - 1) = IIf(Len(oLink.From.Name) > 65, Left(oLink.From.Name, 65) & "... ", oLink.From.Name)
        Else
          .Column(2, .ListCount - 1) = oLink.From.ID
          .Column(7, .ListCount - 1) = IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.From.Name) > 65, Left(oLink.From.Name, 65) & "... ", oLink.From.Name)
        End If
        .Column(3, .ListCount - 1) = Choose(oLink.Type + 1, "FF", "FS", "SF", "SS") & IIf(oLink.Type <> pjFinishToStart, "*", "")
        .Column(4, .ListCount - 1) = Round(oLink.Lag / (ActiveProject.HoursPerDay * 60), 2) & "d"
        Select Case oLink.From.ConstraintType
          Case pjFNET
            If oLink.From.Finish > oLink.From.ConstraintDate Then
              .Column(5, .ListCount - 1) = Format(oLink.From.Finish, "mm/dd/yy")
            Else
              .Column(5, .ListCount - 1) = "<" & Format(oLink.From.Finish, "mm/dd/yy")
            End If
          Case pjFNLT
            If oLink.From.Finish < oLink.From.ConstraintDate Then
              .Column(5, .ListCount - 1) = Format(oLink.From.Finish, "mm/dd/yy")
            Else
              .Column(5, .ListCount - 1) = ">" & Format(oLink.From.Finish, "mm/dd/yy")
            End If
          Case pjMFO
            If oLink.From.Finish = oLink.From.ConstraintDate Then
              .Column(5, .ListCount - 1) = "=" & Format(oLink.From.Finish, "mm/dd/yy")
            Else
              .Column(5, .ListCount - 1) = Format(oLink.From.Finish, "mm/dd/yy")
            End If
          Case Else
            .Column(5, .ListCount - 1) = Format(oLink.From.Finish, "mm/dd/yy")
        End Select
        .Column(6, .ListCount - 1) = Round(oLink.From.TotalSlack / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(8, .ListCount - 1) = IIf(oLink.From.Critical, "X", "")
      End With
    ElseIf oLink.To.Guid <> oTask.Guid Then 'it's a successor
      'handle external tasks
      If blnSubprojects And oLink.To.ExternalTask Then
        'fix the returned UID
        lngLinkUID = oLink.To.GetField(185073906) Mod 4194304
        strProject = oLink.To.Project
        If InStr(oLink.To.Project, "\") > 0 Then
          strProject = Replace(strProject, ".mpp", "")
          strProject = Mid(strProject, InStrRev(strProject, "\") + 1)
        End If
        lngFactor = oSubMap(strProject)
        lngLinkUID = (lngFactor * 4194304) + lngLinkUID
      Else
        If blnSubprojects Then
          lngFactor = Round(oTask / 4194304, 0)
          lngLinkUID = (lngFactor * 4194304) + oLink.To.UniqueID
        Else
          lngLinkUID = oLink.To.UniqueID
        End If
      End If
      With cptNetworkBrowser_frm.lboSuccessors
        .AddItem
        .Column(0, .ListCount - 1) = lngLinkUID
        .Column(1, .ListCount - 1) = lngLinkUID Mod 4194304 & IIf(oLink.To.ExternalTask, "x", "")
        If blnSubprojects And oLink.To.ExternalTask Then
          .Column(2, .ListCount - 1) = ActiveProject.Tasks.UniqueID(lngLinkUID).ID
          .Column(7, .ListCount - 1) = IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.To.Name) > 65, Left(oLink.To.Name, 65) & "... ", oLink.To.Name)
        ElseIf Not blnSubprojects And oLink.To.ExternalTask Then
          .Column(2, .ListCount - 1) = oLink.To.ID
          .Column(7, .ListCount - 1) = IIf(Len(oLink.To.Name) > 65, Left(oLink.To.Name, 65) & "... ", oLink.To.Name)
        Else
          .Column(2, .ListCount - 1) = oLink.To.ID
          .Column(7, .ListCount - 1) = IIf(ActiveProject.Tasks.UniqueID(lngLinkUID).Marked, "[m] ", "") & IIf(Len(oLink.To.Name) > 65, Left(oLink.To.Name, 65) & "... ", oLink.To.Name)
        End If
        .Column(3, .ListCount - 1) = Choose(oLink.Type + 1, "FF", "FS", "SF", "SS") & IIf(oLink.Type <> pjFinishToStart, "*", "")
        .Column(4, .ListCount - 1) = Round(oLink.Lag / (ActiveProject.HoursPerDay * 60), 2) & "d"
        Select Case oLink.To.ConstraintType
          Case pjSNET
            If oLink.To.ConstraintDate > oLink.To.Start Then
              .Column(5, .ListCount - 1) = ">" & Format(oLink.To.Start, "mm/dd/yy")
            Else
              .Column(5, .ListCount - 1) = Format(oLink.To.Start, "mm/dd/yy")
            End If
          Case pjSNLT
            If oLink.To.ConstraintDate = oLink.To.Start Then
              .Column(5, .ListCount - 1) = "<" & Format(oLink.To.Start, "mm/dd/yy")
            Else
              .Column(5, .ListCount - 1) = Format(oLink.To.Start, "mm/dd/yy")
            End If
          Case pjMSO
            .Column(5, .ListCount - 1) = "=" & Format(oLink.To.Start, "mm/dd/yy")
          Case Else
            .Column(5, .ListCount - 1) = Format(oLink.To.Start, "mm/dd/yy")
        End Select
        .Column(6, .ListCount - 1) = Round(oLink.To.TotalSlack / (ActiveProject.HoursPerDay * 60), 2) & "d"
        .Column(8, .ListCount - 1) = IIf(oLink.To.Critical, "X", "")
      End With
    End If
  Next oLink
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Set oTaskDependencies = Nothing
  Set oSubproject = Nothing
  Set oLink = Nothing
  Set oTask = Nothing
  Exit Sub
err_here:
  If Err.Number <> 424 Then Call cptHandleErr("cptNetworkBrowser_bas", "cptShowPreds", Err, Erl)
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
    'todo: "Highlight Marked tasks in the current view?"
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
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    oTask.Marked = False
next_task:
  Next oTask
  ActiveProject.Tasks.UniqueID(0).Marked = False
  'todo: fix this
  If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
    cptSpeed True
    If Edition = pjEditionProfessional Then
      If Not cptFilterExists("Active Tasks") Then
        FilterEdit Name:="Active Tasks", TaskFilter:=True, create:=True, OverwriteExisting:=False, FieldName:="Active", Test:="equals", Value:="Yes", ShowInMenu:=True, ShowSummaryTasks:=True
      End If
      FilterApply "Active Tasks"
    ElseIf Edition = pjEditionStandard Then
      FilterApply "All Tasks"
    End If
    FilterApply "Marked"
    cptSpeed False
  Else
    'todo: if lower pane
  End If
  Set oTask = Nothing

End Sub

Sub cptHistoryDoubleClick()
  Dim lngTaskUID As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  lngTaskUID = CLng(cptNetworkBrowser_frm.lboHistory.Value)
  WindowActivate TopPane:=True
  If IsNumeric(lngTaskUID) Then
    On Error Resume Next
    If Not Find("Unique ID", "equals", lngTaskUID) Then
      If ActiveWindow.TopPane.View.Name = "Network Diagram" Then
        ActiveProject.Tasks.UniqueID(lngTaskUID).Marked = True
        FilterApply "Marked"
        GoTo exit_here
      End If
      If MsgBox("Task is hidden - remove filters and show it?", vbQuestion + vbYesNo, "Confirm Apocalypse") = vbYes Then
        FilterClear
        OptionsViewEx DisplaySummaryTasks:=True
        On Error Resume Next
        If Not OutlineShowAllTasks Then
          If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
            Sort "ID", , , , , , False, True
            OutlineShowAllTasks
          Else
            SelectBeginning
            GoTo exit_here
          End If
        End If
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If Not Find("Unique ID", "equals", lngTaskUID) Then
          MsgBox "Unable to find Task UID " & lngTaskUID & "...", vbExclamation + vbOKOnly, "Task Not Found"
        End If
      End If
    End If
  End If
  
exit_here:
  Exit Sub
err_here:
  Call cptHandleErr("cptNetworkBrowser_bas", "cptHistoryDoubleClick", Err, Erl)
  Resume exit_here
End Sub
