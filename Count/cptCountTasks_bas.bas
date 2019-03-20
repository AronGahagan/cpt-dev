Attribute VB_Name = "cptCountTasks_bas"
'<cpt_version>v1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub CountTasks(strScope As String)
Dim Task As Task, Tasks As Tasks
Dim lgTasks As Long, lgSummary As Long, lgInactive As Long
Dim strMsg As String

  On Error GoTo err_here

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  Select Case strScope
    Case "All"
      On Error Resume Next
      Set Tasks = ActiveProject.Tasks
      If Tasks Is Nothing Or Tasks.count = 0 Then
        MsgBox "There are no tasks in this project.", vbInformation + vbOKOnly, "Task Counter"
        GoTo exit_here
      End If
    Case "Selected"
      On Error Resume Next
      Set Tasks = ActiveSelection.Tasks
      If Tasks Is Nothing Then
        MsgBox "There are no selected tasks.", vbInformation + vbOKOnly, "Task Counter"
        GoTo exit_here
      End If
    Case "Visible"
      SelectAll
      On Error Resume Next
      Set Tasks = ActiveSelection.Tasks
      If Tasks Is Nothing Then
        MsgBox "There are no visible tasks.", vbInformation + vbOKOnly, "Task Counter"
        GoTo exit_here
      End If
  End Select

  For Each Task In Tasks
    If Not Task Is Nothing Then
      If Task.Summary Then
        lgSummary = lgSummary + 1
        If Not Task.Active Then
          lgInactive = lgInactive + 1
          lgSummary = lgSummary - 1
        End If
      Else
        lgTasks = lgTasks + 1
        If Not Task.Active Then
          lgInactive = lgInactive + 1
          lgTasks = lgTasks - 1
        End If
      End If
    End If
  Next Task
  
  strMsg = strScope & " task(s):" & vbCrLf
  strMsg = strMsg & Format(lgSummary, "#,##0") & " summary task(s)" & vbCrLf
  strMsg = strMsg & Format(lgTasks, "#,##0") & " subtask(s)" & vbCrLf
  strMsg = strMsg & Format(lgSummary + lgTasks, "#,##0") & " total task(s)" & vbCrLf
  If lgInactive > 0 Then
    strMsg = strMsg & "(" & Format(lgInactive, "#,##0") & " inactive task(s) not included in total.)"
  End If
  
  MsgBox strMsg, vbInformation + vbOKOnly, "Task Counter"
  
exit_here:
  On Error Resume Next
  Set Tasks = Nothing
  Set Task = Nothing
  Exit Sub
err_here:
  Call HandleErr("cptCountTasks_bas", "CountTasks", err)
  Resume exit_here
End Sub

Sub CountTasksAll()
  Call CountTasks("All")
End Sub

Sub CountTasksSelected()
  Call CountTasks("Selected")
End Sub

Sub CountTasksVisible()
  Call CountTasks("Visible")
End Sub
