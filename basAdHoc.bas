Attribute VB_Name = "basAdHoc"
'<cpt_version>1.0</cpt_version>
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

  lgTasks = ActiveProject.Tasks.Count

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

