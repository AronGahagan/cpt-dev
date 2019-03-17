��Attribute VB_Name = "basAdHoc"
Sub Template86()
Dim Project As Project, Task As Task
Dim vTaskName As Variant

  SpeedON

  For Each Project In Projects
    If Project.ReadOnly Then GoTo next_project
    Project.Activate
    ActiveWindow.TopPane.Activate
    FilterClear
    OutlineShowAllTasks
    For Each Task In Project.Tasks
      If Task.Marked Then Task.Marked = False
      For Each vTaskName In Array("Project Complete", "DC Lease Final Colo Delivery")
        If InStr(Task.Name, vTaskName) > 0 Then
          EditGoTo Task.ID
          Task.Marked = True
          'Task.SetField FieldNameToFieldConstant("Task Owner"), "E2E"
          'Task.SetField FieldNameToFieldConstant("Project Phase"), "Lease.1 E2E Planning"
          'Task.SetField FieldNameToFieldConstant("Change Request Number"), "86"
          'Task.SetField FieldNameToFieldConstant("Change Control Value"), "Template"
          'Task.SetField FieldNameToFieldConstant("Tranche"), "1"
        End If
      Next vTaskName
    Next Task
    FilterClear
    OutlineShowAllTasks
    FilterApply "Marked"
next_project:
  Next Project
  
  Set Project = Nothing
  Set Task = Nothing

  SpeedOFF

End Sub

Sub JeNeSaisQuois()
'unsets and resets the task durations,
'calculating before and afterwards
Dim Task As Task, lgDuration As Long, lgTask As Long, lgTasks As Long

  lgTasks = ActiveProject.Tasks.count

  For Each Task In ActiveProject.Tasks
    If Task.Summary Then GoTo next_task
    lgDuration = Task.Duration
    Application.CalculateProject
    Task.Duration = lgDuration
    Application.CalculateProject
next_task:
    lgTask = lgTask + 1
    Debug.Print lgTask & " / " & lgTasks & " (" & Format(lgTask / lgTasks, "0%")
  Next Task
  
End Sub

Sub FindActiveInactivePreds()
Dim Task As Task, TaskDependency As TaskDependency
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet
Dim lgRow As Long, lgCol As Long
Dim aCol As Variant
  
  On Error GoTo err_here
  
  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  
  Worksheet.Name = "Audit"
  
  aCol = Array("TASK UID", "TASK NAME", "TASK ACTIVE", "PRED UID", "PREDECESSOR NAME", "PRED ACTIVE")
  
  Worksheet.[A1:F1] = aCol

  For Each Task In ActiveProject.Tasks
    If Task Is Nothing Then GoTo next_task
    For Each TaskDependency In Task.TaskDependencies
      If Task = TaskDependency.To Then
        If (Not Task.Active And TaskDependency.From.Active) Or (Task.Active And Not TaskDependency.From.Active) Then
          lgRow = Worksheet.[A1048576].End(xlUp).Row + 1
          Worksheet.Cells(lgRow, 1).Value = Task.UniqueID
          Worksheet.Cells(lgRow, 2).Value = Task.Name
          Worksheet.Cells(lgRow, 3).Value = Task.Active
          Worksheet.Cells(lgRow, 4).Value = TaskDependency.From.UniqueID
          Worksheet.Cells(lgRow, 5).Value = TaskDependency.From.Name
          Worksheet.Cells(lgRow, 6).Value = TaskDependency.From.Active
        End If
      End If
    Next
next_task:
  Next Task

  Worksheet.ListObjects.Add xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), , xlYes
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.Columns.AutoFit
  
  Set rng = Range("Table1[TASK ACTIVE]")
  rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE"
  rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
  With rng.FormatConditions(1).Font
      .Color = -16383844
      .TintAndShade = 0
  End With
  With rng.FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13551615
      .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False
  Set rng = Range("Table1[PRED ACTIVE]")
  rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE"
  rng.FormatConditions(rng.FormatConditions.count).SetFirstPriority
  With rng.FormatConditions(1).Font
      .Color = -16383844
      .TintAndShade = 0
  End With
  With rng.FormatConditions(1).Interior
      .PatternColorIndex = xlAutomatic
      .Color = 13551615
      .TintAndShade = 0
  End With
  rng.FormatConditions(1).StopIfTrue = False

  
exit_here:
  On Error Resume Next
  Set TaskDependency = Nothing
  Set Task = Nothing
  Set rng = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "FindActiveInactivePreds", err)
  Resume exit_here
End Sub

Sub AuditTaskMetaData()
Dim Project As Project, TemplateTask As Task, Task As Task
Dim strProjectName As String
Dim strTemplateTaskName As String, strTaskName As String
Dim lgL2Milestone As Long, strTemplateL2Milestone As String, strL2Milestone As String
Dim lgTaskOwner As Long, strTemplateTaskOwner As String, strTaskOwner As String

  GoTo exit_here

  'LEASE PROJECTS ONLY
  'DC   Tracker Milestones  4-18
  'NID  Milestones          26-36

  For lgTaskId = 3 To 13
    Set Project = Projects("NID Build Template")
    Set TemplateTask = Project.Tasks(lgTaskId)
    strTemplateTaskName = TemplateTask.Name
    strProjectName = RegEx(strTemplateTaskName, "\([A-z0-9]{5}\)\s-\s")
    strTemplateTaskName = Replace(strTemplateTaskName, strProjectName, "")
    Project.Activate
    lgL2Milestone = FieldNameToFieldConstant("L2 Milestone", pjTask)
    strTemplateL2Milestone = TemplateTask.GetField(lgL2Milestone)
    lgTaskOwner = FieldNameToFieldConstant("Task Owner", pjTask)
    strTemplateTaskOwner = TemplateTask.GetField(lgTaskOwner)
    For Each Project In Projects
      If InStr(Project.Name, "Template") = 0 Then
        Project.Activate
        lgL2Milestone = FieldNameToFieldConstant("L2 Milestone", pjTask)
        Set Task = Project.Tasks(lgTaskId)
        strTaskName = Task.Name
        strProjectName = RegEx(strTaskName, "\([A-z0-9]{5}\)\s-\s")
        strTaskName = Replace(strTaskName, strProjectName, "")
        If strTaskName <> strTemplateTaskName Then
          Debug.Print strTemplateTaskName & " <> " & strProjectName & strTaskName
          Task.Name = strProjectName & strTemplateTaskName
        End If
        strL2Milestone = Task.GetField(lgL2Milestone)
        If strTemplateL2Milestone <> strL2Milestone Then
          Debug.Print strTemplateL2Milestone & " <> " & strL2Milestone
          Task.SetField lgL2Milestone, strTemplateL2Milestone
        End If
        strTaskOwner = Task.GetField(lgTaskOwner)
        If strTemplateTaskOwner <> strTaskOwner Then
          Debug.Print strTemplateTaskOwner & " <> " & strTaskOwner
          Task.SetField lgTaskOwner, strTemplateTaskOwner
        End If
      End If
    Next Project
  Next lgTaskId
exit_here:
End Sub

Sub xyz()
Dim lgID As Long

  For lgID = 1 To 868
    If Projects(1).Tasks(lgID).Summary Then GoTo next_task
    If Not Projects(1).Tasks(lgID).Active Then GoTo next_task
    If Projects(1).Tasks(lgID).Start <> Projects(2).Tasks(lgID).Start Then
      Projects(1).Activate
      EditGoTo lgID
      Projects(2).Activate
      EditGoTo lgID
      Debug.Print lgID
    End If
next_task:
  Next lgID

End Sub

Sub ExportCalendarExceptions()
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet
Dim lgRow As Long, lgCol As Long
Dim ce As Exception

  lgCount = ActiveProject.Calendar.Exceptions.count

  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = ActiveProject.Calendar.Name

  Worksheet.[A1:D1] = Array("Exception", "Start", "Finish", "DaysOfWeek")

  For lgException = 1 To ActiveProject.Calendar.Exceptions.count
    Set ce = ActiveProject.Calendar.Exceptions(lgException)
    Worksheet.Cells(lgException + 1, 1).Value = ce.Name
    Worksheet.Cells(lgException + 1, 2).Value = ce.Start
    Worksheet.Cells(lgException + 1, 3).Value = ce.Finish
    Worksheet.Cells(lgException + 1, 4).Value = ce.DaysOfWeek
  Next lgException
  
  Set ce = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  
End Sub

Sub Template91()
Dim Project As Project, Task As Task
Dim lgL2 As Long, strL2 As Long

  SpeedON

  For Each Project In Projects
    Project.Activate
    lgL2 = FieldNameToFieldConstant("L2 Milestone")
    For Each Task In Project.Tasks
      Task.Marked = False
      'check L2 Milestones
      For Each vText In Array("Security Lockdown", "Network Complete", "DC Ready", "Shell Ready", "CTB Complete", "CTD Complete", "CBE Complete", "Ordering Complete", "Schedule Margin", "Site Ready", "RTEG", "RTW")
        If Task.GetField(lgL2) = vText Then Task.Marked = True
      Next vText
      'check task names
      For Each vText In Array("RNG Migration Complete", "HW Orders Complete", "All Day 1", "HLD Complete", "Min Span")
        If InStr(Task.Name, vText) > 0 Then Task.Marked = True
      Next
    Next Task
    FilterApply "Marked"
  Next Project
  
  Set Task = Nothing
  Set Project = Nothing
  
  SpeedOFF
  
End Sub

Sub UpdateColoKey()
Dim Project As Project, Task As Task
Dim lgColoKey As Long, lgSeq As Long
Dim lgL2 As Long

  On Error GoTo err_here

  For Each Project In Projects
    Project.Activate
    lgColoKey = FieldNameToFieldConstant("Colo Key")
    lgL2 = FieldNameToFieldConstant("L2 Milestone")
    lgSeq = 0
    For Each Task In Project.Tasks
      If Task Is Nothing Then GoTo next_task    'groups, empty lines
      If Not Task.Active Then GoTo next_task    'inactive tasks
      If Task.ExternalTask Then GoTo next_task  'cross project links
      If Task.Summary Then GoTo next_task       'summary
      If Task.GetField(lgL2) = "Colo Ready" Then
        lgSeq = lgSeq + 1
        Task.SetField lgColoKey, "DC00" & lgSeq
      End If
next_task:
    Next Task
  Next Project
  
exit_here:
  On Error Resume Next
  Set Task = Nothing
  Set Project = Nothing
  Exit Sub
err_here:
  MsgBox err.Number & ": " & err.Description, vbExclamation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub Template94()
Dim Project As Project, Task As Task
Dim lgField As Long

  On Error GoTo err_here
  
  For Each Project In Projects
    Project.Activate
    lgRFP = 0
    lgSV = 0
    lgKO = 0
    lgAgenda = 0
    lgCSV = 0
    lgUID = 0
    
    ViewApply "_Template Change View"
    FilterClear
    GroupClear
    OutlineShowAllTasks
    lgPhase = FieldNameToFieldConstant("Project Phase")
    lgOwner = FieldNameToFieldConstant("Task Owner")
    lgCN = FieldNameToFieldConstant("Change Request Number")
    lgCV = FieldNameToFieldConstant("Change Control Value")
    lgT = FieldNameToFieldConstant("Tranche")
    For Each Task In Project.Tasks
      If Right(Task.Name, 7) = ") - RFP" And Task.Summary Then
        lgRFP = Task.UniqueID
        Debug.Print Project.Name & " RFP UID = " & lgRFP
      End If
      If Right(Task.Name, 15) = ") - Site Visits" And Task.Summary Then
        lgSV = Task.UniqueID
        Debug.Print Project.Name & " Site Visits UID = " & lgSV
      End If
      If Right(Task.Name, 35) = ") - Conduct Project Kickoff Meeting" Then
        lgKO = Task.UniqueID
        Debug.Print Project.Name & " Kickoff UID = "; lgKO
      End If
      If InStr(Task.Name, ") - Confirm travel agenda") > 0 Then
        lgAgenda = Task.UniqueID
        Debug.Print Project.Name & " Travel Agenda UID = " & lgAgenda
      End If
      If Right(Task.Name, 23) = ") - Conduct Site Visit " Then
        lgCSV = Task.UniqueID
        Debug.Print Project.Name & " Conduct Site Visit UID = " & lgCSV
      End If
    Next
    EditGoTo Project.Tasks.UniqueID(lgRFP).ID
    strProject = RegEx(Project.Tasks.UniqueID(lgRFP).Name, "^\(.*\) - ")
    Set Task = Project.Tasks.Add(strProject & "Create DCO Calendar Placeholder for Site Visits", Project.Tasks.UniqueID(lgRFP).ID)
    Task.Duration = 10 * 480
    Task.SetField lgPhase, "Lease.2 Site Selection"
    Task.SetField lgOwner, "SSA"
    Task.SetField lgCN, "94"
    Task.SetField lgCV, "Template"
    Task.SetField lgT, "1"
    Task.UniqueIDPredecessors = CStr(lgKO)
    Task.UniqueIDSuccessors = CStr(lgAgenda)
    lgUID = Task.UniqueID
    
    EditGoTo Project.Tasks.UniqueID(lgSV).ID
    Set Task = Project.Tasks.Add(strProject & "Update DCO Calendar with Confirmed Team Site Visit Dates", Project.Tasks.UniqueID(lgSV).ID + 2)
    Task.Duration = 10 * 480
    Task.SetField lgPhase, "Lease.2 Site Selection"
    Task.SetField lgOwner, "SSA"
    Task.SetField lgCN, "94"
    Task.SetField lgCV, "Template"
    Task.SetField lgT, "1"
    Task.UniqueIDPredecessors = CStr(lgUID)
    Task.UniqueIDSuccessors = (lgCSV)
    
  Next Project
  
exit_here:
  On Error Resume Next
  Set Project = Nothing
  Set Task = Nothing
  Exit Sub
  
err_here:
  Call HandleErr("basAdHoc", "Template94", err)
  Resume exit_here
  
End Sub

Sub Template95()
Dim Project As Project, Task As Task

  On Error GoTo err_here
  For Each Project In Projects
    Project.Activate
    If Project.Name = "Standard Lease Template_" Then GoTo next_project
    lgPhase = FieldNameToFieldConstant("Project Phase")
    lgOwner = FieldNameToFieldConstant("Task Owner")
    lgCN = FieldNameToFieldConstant("Change Request Number")
    lgCV = FieldNameToFieldConstant("Change Control Value")
    lgT = FieldNameToFieldConstant("Tranche")
    
    ViewApply "_Template Change View"
    FilterClear
    GroupClear
    OutlineShowAllTasks
    blnApply = False
    For Each Task In Project.Tasks
      If InStr(Task.Name, "Class B") > 0 And Task.Active Then
        blnApply = True
      End If
      If InStr(Task.Name, ") - DCPS SOW Document Issuance to DCOPS") > 0 Then
        lgPred1 = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Temporary Secure Storage Available  (1st Tranche)") > 0 Then
        lgPred2 = Task.UniqueID
      ElseIf InStr(Task.Name, ") - SOC Beneficial Use") > 0 Then
        lgPred3 = Task.UniqueID
      ElseIf InStr(Task.Name, ") - DC Ready for First Tranche") > 0 Then
        lgSucc1 = Task.UniqueID
      ElseIf InStr(Task.Name, ") - DC Ready for Second Tranche") > 0 Then
        lgSucc2 = Task.UniqueID
'      ElseIf InStr(Task.Name, "") > 0 Then
'        lgSucc3 = Task.UniqueID
      End If
    Next
    If Not blnApply Then GoTo next_project
    EditGoTo Project.Tasks.UniqueID(lgPred1).ID
    strProject = RegEx(Project.Tasks.UniqueID(lgPred1).Name, "^\(.*\) - ")
    Set Task = Project.Tasks.Add(strProject & "DC Ready Security Requirements", Project.Tasks.UniqueID(lgPred1).ID + 1)
    Task.SetField lgCN, "95"
    Task.SetField lgCV, "Template"
    lgUID = Task.UniqueID
    lgOl = Task.OutlineLevel
    Set Task = Project.Tasks.Add(strProject & "Minimum Security Requirements complete for DC Ready (Class B)", Project.Tasks.UniqueID(lgUID).ID + 1)
    Task.OutlineLevel = lgOl + 1
    Task.Duration = 15 * 480
    Task.SetField lgPhase, "Lease.1 E2E Planning"
    Task.SetField lgOwner, "DCPS"
    Task.SetField lgCN, "95"
    Task.SetField lgCV, "Template"
    Task.SetField lgT, "1"
    Task.UniqueIDPredecessors = CStr(lgPred1 & "," & lgPred2 & "," & lgPred3)
    Task.UniqueIDSuccessors = CStr(lgSucc1 & "," & lgSucc2)
next_project:
  Next Project
  
exit_here:
  On Error Resume Next
  Set Task = Nothing
  Set Project = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "Template95", err)
  Resume exit_here
End Sub

Sub Template96()
Dim Project As Project, Task As Task

  On Error GoTo 0 'err_here
  
  For Each Project In Projects
    Project.Activate
    If InStr(Project.Name, "Standard Lease Template_") > 0 Then GoTo next_project
    lgPhase = FieldNameToFieldConstant("Project Phase")
    lgOwner = FieldNameToFieldConstant("Task Owner")
    lgCN = FieldNameToFieldConstant("Change Request Number")
    lgCV = FieldNameToFieldConstant("Change Control Value")
    lgT = FieldNameToFieldConstant("Tranche")
    
    ViewApply "_Template Change View"
    FilterClear
    GroupClear
    OutlineShowAllTasks

    blnApply = False 'regarding class c/d

    For Each Task In Project.Tasks
      If InStr(Task.Name, ") - DC Ready for First Tranche") > 0 Or InStr(Task.Name, ") - DC Ready - 1st Tranche") > 0 Then
        lgDCFirst = Task.UniqueID
      ElseIf InStr(Task.Name, ") - DC Ready for Second Tranche") > 0 Or InStr(Task.Name, ") - DC Ready - 2nd Tranche") > 0 Then
        lgDCSecond = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Fitout of Secured Storage Space Complete") > 0 Then
        lgFitout = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Network Equipment Lead Time / Docked at site (OOB included)") > 0 Then
        lgNetEQuip = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Metro Fiber Installed (t-2)") > 0 Then
        lgMetro = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Dark Fiber Installed (t-2)") > 0 Then
        lgDark = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Lit waves Circuits installed") > 0 Then
        lgLit = Task.UniqueID
      ElseIf InStr(Task.Name, ") - OOB Circuit Delivered to rack") > 0 Then
        lgOOB = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Fiber Delivery Complete") > 0 Then
        lgFiber = Task.UniqueID
      ElseIf InStr(Task.Name, ") - DCPS Testing & Cx of Security Equipment") > 0 Then
        lgCx = Task.UniqueID
      ElseIf InStr(Task.Name, ") - Datacenter Physical Security (DCPS) - Class C, D") > 0 And Task.Active Then
        blnApply = True
      ElseIf InStr(Task.Name, ") - First Footprint Colo Ready") > 0 Then
        lgFFCR = Task.UniqueID
      End If
    Next
    
    dtBefore = Project.Tasks.UniqueID(lgFFCR).Finish
    
    'add the new task
    If blnApply Then
      EditGoTo Project.Tasks.UniqueID(lgCx).ID
      strProject = RegEx(Project.Tasks.UniqueID(lgCx).Name, "^\(.*\) - ")
      Set Task = Project.Tasks.Add(strProject & "Minimum Security Requirements complete for DC Ready (Class C/D)", Project.Tasks.UniqueID(lgCx).ID + 1)
      Task.Duration = 0
      Task.SetField lgPhase, "Lease.1 E2E Planning"
      Task.SetField lgOwner, "DCPS"
      Task.SetField lgCN, "96"
      Task.SetField lgCV, "Template"
      Task.SetField lgT, "1"
      Task.UniqueIDPredecessors = CStr(lgFitout)
      Task.UniqueIDSuccessors = CStr(lgDCFirst & "," & lgDCSecond)
      lgUID = Task.UniqueID
    End If
    
    'update network equipment
    Set Task = Project.Tasks.UniqueID(lgNetEQuip)
    EditGoTo Task.ID
    Task.UniqueIDPredecessors = Task.UniqueIDPredecessors & "," & CStr(lgDCFirst) & "FF-1 wk"
    Task.SetField lgCN, Task.GetField(lgCN) & ",96"
    
    'update DC Ready
    Set Task = Project.Tasks.UniqueID(lgDCFirst)
    EditGoTo Task.ID
    Task.SetField lgCN, Task.GetField(lgCN) & ",96"
    Task.UniqueIDSuccessors = Task.UniqueIDSuccessors & "," & CStr(lgMetro) & "FF-3 wk"
    Task.UniqueIDSuccessors = Task.UniqueIDSuccessors & "," & CStr(lgDark) & "FF-2 wk"
    Task.UniqueIDSuccessors = Task.UniqueIDSuccessors & "," & CStr(lgLit) & "FF-2 wk"
    Task.UniqueIDSuccessors = Task.UniqueIDSuccessors & "," & CStr(lgOOB)
    
    For Each vTask In Array(lgMetro, lgDark, lgLit, lgOOB)
      Set Task = Project.Tasks.UniqueID(vTask)
      EditGoTo Task.ID
      Task.SetField lgCN, Task.GetField(lgCN) & ",96"
    Next vTask
    
    dtAfter = Project.Tasks.UniqueID(lgFFCR).Finish
    
    If dtAfter <> dtBefore Then
      Debug.Print Project.Name & " FFCR changed from " & FormatDateTime(dtBefore, vbShortDate) & " to " & FormatDateTime(dtAfter, vbShortDate)
    End If
    
next_project:
  Next Project
  
exit_here:
  On Error Resume Next
  Set Task = Nothing
  Set Project = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "Template95", err)
  Resume exit_here
End Sub

Sub AuditTrackerMilestoneLogic()
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet
Dim Project As Project, Task As Task

  On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  
  lgL2 = FieldNameToFieldConstant("L2 Milestone")
  
  For Each Task In ActiveProject.Tasks
    If Not Task.Active Then GoTo next_task
    strMilestone = Task.GetField(lgL2)
    If Len(strMilestone) = 0 Then GoTo next_task
    lgTask = lgTask + 1
    If Workbook.Sheets.count < lgTask Then Set Worksheet = Workbook.Sheets.Add
    If lgTask = 1 Then Set Worksheet = Workbook.Sheets(1)
    Worksheet.Activate
    Worksheet.Name = Replace(Replace(strMilestone, "/", " "), "- ", "")
    Worksheet.[A1].Value = strMilestone & " Predecessors"
    Worksheet.[A1].Font.Bold = True
    Worksheet.[A1].Font.Size = 20
    Worksheet.[A2].Value = "TEMPLATE"
    lgPred = 0
    For Each PredecessorTask In Task.PredecessorTasks
      lgPred = lgPred + 1
      Worksheet.Cells(2 + lgPred, 1).Value = PredecessorTask.Name
    Next
    lgProject = 0
    For Each Project In Projects
      If InStr(Project.Name, "Standard Lease Template_") > 0 Then GoTo next_project
      'Project.Activate
      lgProject = lgProject + 1
      Worksheet.Cells(2, lgProject + 1).Value = Project.Name
      Set pTask = Nothing
      For Each pTask In Project.Tasks
        If pTask.GetField(lgL2) = strMilestone Then Exit For
      Next pTask
      If Not pTask Is Nothing Then
        lgPred = 0
        For Each PredecessorTask In pTask.PredecessorTasks
          lgPred = lgPred + 1
          Worksheet.Cells(2 + lgPred, lgProject + 1).Value = PredecessorTask.Name
        Next PredecessorTask
      End If
next_project:
    Next Project
    xlApp.ActiveWindow.Zoom = 85
    Worksheet.Columns.AutoFit
next_task:
  Next Task
    
exit_here:
  On Error Resume Next
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set Task = Nothing
  Set Project = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "AuditTrackerMilestoneLogic", err)
  Resume exit_here
End Sub

Sub WhatTemplates()
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet
Dim Project As Project, Task As Task

  SpeedON

  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)

  lgCol = 1
  lgRow = 1
  For Each Project In Projects
    Project.Activate
    
    lgCol = lgCol + 1
    
    Worksheet.Cells(1, lgCol).Value = Project.Name
    OptionsViewEx displaysummarytasks:=False

    For lgCN = 46 To 96
      lgRow = lgRow + 1
      SetAutoFilter FieldName:="Change Request Number", FilterType:=pjAutoFilterCustom, Test1:="contains", Criteria1:=CStr(lgCN)
      SelectAll
      On Error Resume Next
      Worksheet.Cells(lgRow, 1) = lgCN
      Worksheet.Cells(lgRow, lgCol).Value = ActiveSelection.Tasks.count
    Next lgCN
    FilterClear
    OptionsViewEx displaysummarytasks:=True
    On Error GoTo err_here
    lgRow = 1
  Next Project
  
exit_here:
  On Error Resume Next
  SpeedOFF
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set xlApp = Nothing
  Set Project = Nothing
  Set Task = Nothing
  Exit Sub
err_here:
  Call HandleErr("module", "procedure", err)
  Resume exit_here
End Sub

Sub ShowTemplate()
Dim Project As Project
  lgCN = InputBox("Change Control Number:", "Show Template Tasks")
  For Each Project In Projects
    Project.Activate
    SetAutoFilter FieldName:="Change Request Number", FilterType:=pjAutoFilterCustom, Test1:="contains", Criteria1:=CStr(lgCN)
  Next Project
End Sub

Sub Template97()
Dim blnBuild As Boolean
Dim Project As Project, Task As Task

On Error GoTo 0 'err_here

  For Each Project In Projects
    If InStr(Project.Name, "Template") > 0 Then GoTo next_project
    Project.Activate
    OutlineShowAllTasks
    FilterClear
    'determine size
    Select Case Project.Name
      Case "53USW E2E"
        strSize = "M"
      Case "54USW E2E"
        strSize = "M"
      Case "BY3 EX2 E2E"
        strSize = "XS"
      Case Else
        strSize = "L"
    End Select

    strProject = RegEx(Project.Tasks(10).Name, "^\(.*\) - ")
        
    'determine type
    blnBuild = Project.ProjectSummaryTask.GetField(FieldNameToFieldConstant("Program Type", pjProject)) = "Build"
    
    If blnBuild Then
      For Each Task In Project.Tasks
        If Task Is Nothing Then GoTo next_task
        If Task.ExternalTask Then GoTo next_task
        If Task.Summary Then GoTo next_task
        If Not Task.Active Then GoTo next_task
        
        EditGoTo Task.ID
        
        'name updates
        If InStr(Task.Name, "NW Complete") Then
          Task.Name = Replace(Task.Name, "Complete", "Live")
          If InStr(Task.OutlineParent.Name, "First") > 0 Then
            Task.Name = Task.Name & " for First Tranche"
            Project.Tasks.Add strProject & "Existing DC Upgrades Complete", Task.ID
            Project.Tasks.Add strProject & "RNG Fabric Migration Complete", Task.ID
          End If
          GoTo next_task
        End If
        lgTranche = FieldNameToFieldConstant("Tranche")
        'durations
        If InStr(Task.Name, ") - R/S/P") > 0 Then
          Task.Duration = Switch(strSize = "L", "25 edays", strSize = "M", "10 edays")
          If CLng(Task.GetField(lgTranche)) > 1 Then Task.Duration = "15 edays"
          GoTo next_task
        End If
        If InStr(Task.Name, ") - OOB Complete") > 0 Then
          Task.Duration = Switch(strSize = "L", "7 edays", strSize = "M", "5 edays")
          GoTo next_task
        End If
        If InStr(Task.Name, ") - MGFX Complete") > 0 Then
          Task.Duration = Switch(strSize = "L", "18 edays", strSize = "M", "14 edays")
          If CLng(Task.GetField(lgTranche)) > 1 Then
            Task.Duration = "15 edays"
            Project.Tasks.Add strProject & "Network Live for " & Task.GetField(lgTranche) & " Tranche", Task.ID + 1
'            Project.Tasks.Add "(" & strProject & ") - RNG Fabric Migration Complete", Task.ID
            GoTo next_task
          End If
        End If
        If InStr(Task.Name, ") - Optical Complete") > 0 Then
          Task.Duration = Switch(strSize = "L", "14 edays", strSize = "M", "10 edays")
          GoTo next_task
        End If
        If InStr(Task.Name, ") - WAN Complete") > 0 Then
          Task.Duration = Switch(strSize = "L", "14 edays", strSize = "M", "14 edays")
          GoTo next_task
        End If
        If InStr(Task.Name, ") - Fabric Complete") > 0 Then
          If Task.Duration > 0 Then
            Task.Duration = Switch(strSize = "L", "7 edays", strSize = "M", "7 edays")
            GoTo next_task
          End If
        End If
        
next_task:
      Next Task
    ElseIf Not blnBuild Then 'lease
'      For Each Task In Project.Tasks
'        If Task Is Nothing Then GoTo next_task
'        If Task.ExternalTask Then GoTo next_task
'        If Not Task.Active Then GoTo next_task
'
'        'get project name
'        strProject = RegEx(Task.Name, "^\(.*\) - ")
'
'        'adjust durations
'        If InStr(Task.Name, ") - R/S/P") > 0 Then
'          Task.Duration = Switch(strSize = "L", "25 edays", strSize = "M", "10 edays")
'          If CLng(Task.GetField(lgTranche)) > 1 Then Task.Duration = "15 edays"
'        End If
'        If InStr(Task.Name, ") - OOB Complete") > 0 Then
'          Task.Duration = Switch(strSize = "L", "7 edays", strSize = "M", "5 edays")
'        End If
'        If InStr(Task.Name, ") - MGFX Complete") > 0 Then
'          Task.Duration = Switch(strSize = "L", "18 edays", strSize = "M", "14 edays")
'          If CLng(Task.GetField(lgTranche)) > 1 Then
'            Task.Duration = "15 edays"
'        End If
'        If InStr(Task.Name, ") - Optical Complete") > 0 Then
'          Task.Duration = Switch(strSize = "L", "14 edays", strSize = "M", "10 edays")
'        End If
'        If InStr(Task.Name, ") - Wan Complete") > 0 Then
'          Task.Duration = Switch(strSize = "L", "14 edays", strSize = "M", "14 edays")
'        End If
'        If InStr(Task.Name, ") - Fabric Complete") > 0 Then
'          Task.Duration = Switch(strSize = "L", "7 edays", strSize = "M", "7 edays")
'        End If
'
'        'convert to workdays
'        lgDuration = Application.DateDifference(Task.Start, Task.Finish)
'        Task.Duration = lgDuration & " days"
'
'        'revise task name
'        If InStr(Task.Name, ") - NW Complete") > 0 Then
'          Task.Name = Replace(Task.Name, "Complete", "Live")
'          lgNWLive = Task.UniqueID
'        End If
'
'      Next Task
    End If
    
next_project:
  Next Project

exit_here:
  On Error Resume Next
  Set Project = Nothing
  Set Task = Nothing
  Set Resource = Nothing
  Set Assignment = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "Template97", err)
  Resume exit_here

End Sub

Sub ListURL()
Dim Project As Project
  For Each Project In Projects
    Debug.Print Project.Name & " - https://microsoft.sharepoint.com/teams/mcio_eppm_prod/project%20detail%20pages/schedule.aspx?projuid=" & Replace(Replace(Project.GetServerProjectGuid, "{", ""), "}", "") & "&ret=0"
  Next Project
  Set Project = Nothing
End Sub

Sub CollectTaskOwners()
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Range
Dim Project As Project, Task As Task

  On Error GoTo err_here

  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  
  xlApp.ScreenUpdating = False
  xlApp.Calculation = xlCalculationManual
  
  Worksheet.Name = "Task Owners"

  Worksheet.[A1:D1] = Array("METRO", "PROJECT", "OWNER", "NAME", "EMAIL")
  lgRow = 1
  
  For Each Project In Projects
    strProject = Project.Name
    strMetro = Project.ProjectSummaryTask.GetField(FieldNameToFieldConstant("Metro", pjProject))
    lgTaskOwner = FieldNameToFieldConstant("Task Owner")
    For Each Task In Project.Tasks
      If Not Task Is Nothing Then
        If Not Task.Active Then GoTo next_task
        If Task.Summary Then GoTo next_task
        If Task.ExternalTask Then GoTo next_task
        If Not IsDate(Task.ActualFinish) Then
          lgRow = lgRow + 1
          Worksheet.Cells(lgRow, 1) = strMetro
          Worksheet.Cells(lgRow, 2) = Project.Name
          Worksheet.Cells(lgRow, 3) = Task.GetField(lgTaskOwner)
        End If
      End If
next_task:
    Next Task
  Next Project

  xlApp.ActiveWindow.Zoom = 85
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  Worksheet.Columns.AutoFit
  Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)).RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
  
  MsgBox "Exported.", vbOKOnly, "Complete"
  
exit_here:
  On Error Resume Next
  xlApp.ScreenUpdating = True
  xlApp.Calculation = xlCalculationAutomatic
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set xlApp = Nothing
  Set Project = Nothing
  Set Task = Nothing
  Exit Sub
err_here:
  Call HandleErr("module", "procedure", err)
  Resume exit_here
End Sub
'
'Sub adhoc_nid()
'Dim Project As Project, Task As Task
'
'  For Each Project In Projects
'    Project.Activate
'    ActiveWindow.TopPane.Activate
'    FilterClear
'    GroupClear
'    OutlineShowAllTasks
'    blnDone = False
'    For Each Task In Project.Tasks
'      If Not Task Is Nothing Then
'        If Task.Summary Then GoTo next_task
'        If Not Task.Active Then GoTo next_task
'        If Task.ExternalTask Then GoTo next_task
'        If IsDate(Task.ActualStart) Then GoTo next_task
'        If Task.GetField(FieldNameToFieldConstant("Task Owner")) = "NID" Then
'          For Each strTask In Array("Network PO's Created", "Network Equipment PO's approved", "Network Equipment Lead Time / Docked at site (OOB included)")
'            If InStr(Task.Name, strTask) > 0 Then
'              EditGoTo Task.ID
'              If InStr(strTask, "Created") > 0 Then
'                Task.Duration = "5d"
'                Exit For
'              End If
'              If InStr(strTask, "approved") > 0 Then
'                Task.Duration = "10d"
'                Exit For
'              End If
'              If InStr(strTask, "Lead") > 0 Then
'                Task.Duration = "75edays"
'                Task.Duration = (Application.DateDifference(Task.Start, Task.Finish) / (60 * 8)) & "d"
'                blnDone = True
'                Exit For
'              End If
'            End If
'          Next strTask
'        End If
'        If blnDone Then GoTo next_project
'      End If
'next_task:
'    Next Task
'next_project:
'  Next Project
'
'End Sub

Sub CHECK_NID_SLA()
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Range
Dim Project As Project, Task As Task
Dim ListObject As ListObject

  On Error GoTo err_here

  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = xlCalculationManual
  Set Worksheet = Workbook.Sheets(1)
  
  lgRow = 1
  
  Worksheet.[A1:G1] = Array("METRO", "PROJECT", "DC READY", "CTB COMPLETE", "LATEST", "COLO READY", "VARIANCE")
  lgMetro = 1
  lgProject = 2
  lgDC = 3
  lgCTB = 4
  lgMAX = 5
  lgCR = 6
  lgVAR = 7
  
  For Each Project In Projects
    Project.Activate
    lgL2 = FieldNameToFieldConstant("L2 Milestone")
    lgRow = lgRow + 1
    Worksheet.Cells(lgRow, lgMetro) = Project.ProjectSummaryTask.GetField(FieldNameToFieldConstant("Metro"))
    Worksheet.Cells(lgRow, lgProject) = Project.Name
    For Each Task In Project.Tasks
      If Task Is Nothing Then GoTo next_task
      If Task.Summary Then GoTo next_task
      If Task.ExternalTask Then GoTo next_task
      If Not Task.Active Then GoTo next_task
      If Task.GetField(FieldNameToFieldConstant("Tranche")) <> "1" Then GoTo next_task
      
      If Task.GetField(lgL2) = "DC Ready" Then
        Worksheet.Cells(lgRow, lgDC) = Task.Finish
        GoTo next_task
      ElseIf Task.GetField(lgL2) = "CTB Complete" Then
        Worksheet.Cells(lgRow, lgCTB) = Task.Finish
        GoTo next_task
      ElseIf Task.GetField(lgL2) = "CTD Complete" Then
        Worksheet.Cells(lgRow, lgCR) = Task.Finish
        GoTo next_task
      ElseIf Task.GetField(lgL2) = "Network Live" Then
        Worksheet.Cells(lgRow, lgNL) = Task.Finish
        GoTo next_task
      End If
      
next_task:
    Next Task
  Next Project

  'make it a table
  Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), , xlYes)
  ListObject.Name = "SLA_VARIANCE"
  'format it
  Worksheet.Range("SLA_VARIANCE[[DC READY]:[COLO READY]]").NumberFormat = "m/d/yyyy"
  Worksheet.Range("SLA_VARIANCE[[DC READY]:[COLO READY]]").HorizontalAlignment = xlCenter
  'add formulae
  ListObject.ListColumns(lgMAX).DataBodyRange.FormulaR1C1 = "=MAX(SLA_VARIANCE[@[DC READY]:[CTB COMPLETE]])"
  ListObject.ListColumns(lgVAR).DataBodyRange.FormulaR1C1 = "=DAYS(SLA_VARIANCE[@[COLO READY]],SLA_VARIANCE[@LATEST])"
  'autofit columns
  ListObject.Range.Columns.AutoFit
  'sort by metro, project
  ListObject.Sort.SortFields.Clear
  ListObject.Sort.SortFields.Add2 key:=Worksheet.Range("SLA_VARIANCE[METRO]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  ListObject.Sort.SortFields.Add2 key:=Worksheet.Range("SLA_VARIANCE[PROJECT]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ListObject.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
  
  xlApp.Calculation = xlCalculationAutomatic
  xlApp.Visible = True
  
exit_here:
  On Error Resume Next
  Set ListObject = Nothing
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set rng = Nothing
  Set xlApp = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "CHECK_NID_SLA", err)
  Resume exit_here
End Sub

Sub SortByCritical()
  Sort Key1:="Finish", Key2:="Duration", Ascending2:=False, Renumber:=False
End Sub

Sub CountDays()
Dim dtStart As Date, dtFinish As Date

  If ActiveSelection.Tasks.count <> 2 Then Exit Sub
  dtStart = WorksheetFunction.Min(ActiveSelection.Tasks(1).Finish, ActiveSelection.Tasks(2).Finish)
  dtFinish = WorksheetFunction.Max(ActiveSelection.Tasks(1).Finish, ActiveSelection.Tasks(2).Finish)
  lgCalDays = WorksheetFunction.Days(dtFinish, dtStart)
  lgWorkDays = WorksheetFunction.NetworkDays(dtStart, dtFinish)
  MsgBox "From " & FormatDateTime(dtStart, vbShortDate) & " to " & FormatDateTime(dtFinish, vbShortDate) & ":" & vbCrLf & vbCrLf & lgCalDays & " calendar days" & vbCrLf & vbCrLf & lgWorkDays & " workdays", vbInformation + vbOKOnly, "Count Days"

End Sub

Sub BulkAppend()
Dim Task As Task
  strAppend = InputBox("Append what text to selected tasks?", "Bulk Append")
  For Each Task In ActiveSelection.Tasks
    Task.Name = Trim(Task.Name) & " " & Trim(strAppend)
  Next Task
  Set Task = Nothing
End Sub

Sub ReplaceInOutlineChildren()
Dim Task As Task, ChildTask As Task
Dim lgParentLevel As Long
  
  SpeedON
  
  Set Task = ActiveSelection.Tasks(1)
  OutlineShowSubTasks
  lgParentLevel = Task.OutlineLevel
  lgParentId = Task.ID
  strReplace = Right(Task.Name, 5)
  For lgID = lgParentId + 1 To ActiveProject.Tasks.count
    EditGoTo lgID
    Set Task = ActiveSelection.Tasks(1)
    If Task.OutlineLevel > lgParentLevel Then
      Task.Marked = True
      Task.Name = Replace(Task.Name, "XXXXX", strReplace)
    Else
      Exit For
    End If
  Next lgID
  EditGoTo lgParentId
  OutlineHideSubTasks
  Set Task = Nothing
  SpeedOFF
End Sub

Sub ApplyViewAll()
Dim Project As Project
  strView = InputBox("Apply which view?", "Apply View All")
  For Each Project In Projects
    If InStr(Project.Name, "AMERS") = 0 Then
      Project.Activate
      ViewApply strView
      GroupClear
      FilterClear
    End If
  Next Project
  Set Project = Nothing
End Sub

Sub AlignRES()
Dim Tasks As Tasks, Task As Task, dtDeadline As Date
Dim strL2 As String, strProject As String, strTranche As String

  On Error GoTo err_here
  SpeedON
  Projects("RES-AMERSWest.mpp").Activate
  Set Tasks = ActiveSelection.Tasks
  lgL2 = FieldNameToFieldConstant("L2 Milestone")
  lgTranche = FieldNameToFieldConstant("Tranche")
  For Each Task In Tasks
    If Not Task.Summary And Task.Active And Not Task.ExternalTask Then
      If Left(Task.Name, 1) <> "(" Then GoTo next_task
      strProject = Mid(Task.Name, 2, 5) & " E2E"
      If Len(strProject) = 0 Then GoTo exit_here
      strL2 = Task.GetField(lgL2)
      If Len(strL2) = 0 Then GoTo next_task
      strTranche = Task.GetField(lgTranche)
      dtDeadline = GetFinishDate(strProject, strL2, strTranche)
      If dtDeadline > 0 Then
        Task.Deadline = dtDeadline
      Else
        Debug.Print strProject & " / " & strL2 & " / " & "Tranche " & strTranche & " not found!", vbExclamation + vbOKOnly
      End If
      dtDeadline = 0
    End If
next_task:
  Next Task
  SpeedOFF
  MsgBox "Aligned", vbExclamation + vbOKOnly
  
exit_here:
  SpeedOFF
  On Error Resume Next
  Set Project = Nothing
  Set Tasks = Nothing
  Set Task = Nothing
  Exit Sub
err_here:
  Call HandleErr("basAdHoc", "AlignRES", err)
  Resume exit_here
End Sub

Function GetFinishDate(strProject As String, strL2 As String, strTranche As String) As Date
Dim Task As Task, lgL2 As Long

  On Error Resume Next
  If strProject = "BY3X2 E2E" Then strProject = "BY3 EX2 E2E"
  Set Project = Projects(strProject)
  On Error GoTo 0
  If Project Is Nothing Then
    strProject = InputBox("Align which project?", "Align RES")
    Set Project = Projects(strProject)
  End If

  lgL2 = FieldNameToFieldConstant("L2 Milestone")
  lgTranche = FieldNameToFieldConstant("Tranche")
  
  For Each Task In Project.Tasks
    If Not Task.Summary And Task.Active And Not Task.ExternalTask Then
      If Task.GetField(lgL2) = strL2 And Task.GetField(lgTranche) = strTranche Then
        'EditGoTo Task.ID
        GetFinishDate = Task.Finish
        Exit Function
      End If
    End If
  Next Task
  
End Function

Sub CountTasks()
Dim lgProject As Long, lgTasks As Long

  For lgProject = 1 To ActiveProject.Subprojects.count
    Debug.Print ActiveProject.Subprojects(lgProject).SourceProject.Name & " = " & ActiveProject.Subprojects(lgProject).SourceProject.Tasks.count & " tasks"
    lgTasks = lgTasks + ActiveProject.Subprojects(lgProject).SourceProject.Tasks.count
  Next lgProject
  
  lgTasks = lgTasks + ActiveProject.Tasks.count
  
  Debug.Print "Total Tasks = " & Format(lgTasks, "#,##0")
  
End Sub

Sub ExportCustomFields()
Dim lgField As Long, aFields As Object

  Set aFields = CreateObject("System.Collections.ArrayList")

  On Error Resume Next
  For lgField = 188776000 To 188778000
    If Application.FieldConstantToFieldName(lgField) <> "<Unavailable>" Then
      Debug.Print lgField & ": " & Application.FieldConstantToFieldName(lgField)
    End If
  Next lgField
  
End Sub
