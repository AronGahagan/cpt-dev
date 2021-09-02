Attribute VB_Name = "cptCountTasks_bas"
'<cpt_version>v1.0.3</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptCountTasks(strScope As String)
Dim oTask As Task, oTasks As Tasks
Dim lngTasks As Long, lngSummary As Long, lngInactive As Long, lngActive As Long
Dim strMsg As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Task Counter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the oTask table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  Select Case strScope
    Case "All"
      On Error Resume Next
      Set oTasks = ActiveProject.Tasks
      If oTasks Is Nothing Or oTasks.Count = 0 Then
        MsgBox "There are no Tasks in this project.", vbInformation + vbOKOnly, "Task Counter"
        GoTo exit_here
      End If
    Case "Selected"
      On Error Resume Next
      Set oTasks = ActiveSelection.Tasks
      If oTasks Is Nothing Then
        MsgBox "There are no selected Tasks.", vbInformation + vbOKOnly, "Task Counter"
        GoTo exit_here
      End If
    Case "Visible"
      SelectAll
      On Error Resume Next
      Set oTasks = ActiveSelection.Tasks
      If oTasks Is Nothing Then
        MsgBox "There are no visible Tasks.", vbInformation + vbOKOnly, "Task Counter"
        GoTo exit_here
      End If
  End Select
  
  If Edition = pjEditionProfessional Then
    lngActive = FieldNameToFieldConstant("Active")
  Else
    lngActive = 0
  End If
  
  For Each oTask In oTasks
    If Not oTask Is Nothing Then
      If oTask.Summary Then
        lngSummary = lngSummary + 1
        If lngActive > 0 Then
          If oTask.GetField(lngActive) = "No" Then
            lngInactive = lngInactive + 1
            lngSummary = lngSummary - 1
          End If
        End If
      Else
        lngTasks = lngTasks + 1
        If lngActive > 0 Then
          If oTask.GetField(lngActive) = "No" Then
            lngInactive = lngInactive + 1
            lngTasks = lngTasks - 1
          End If
        End If
      End If
    End If
  Next oTask
  
  strMsg = strScope & " Task(s):" & vbCrLf
  strMsg = strMsg & Format(lngSummary, "#,##0") & " summary Task(s)" & vbCrLf
  strMsg = strMsg & Format(lngTasks, "#,##0") & " subTask(s)" & vbCrLf
  strMsg = strMsg & Format(lngSummary + lngTasks, "#,##0") & " total Task(s)" & vbCrLf
  If lngInactive > 0 Then
    strMsg = strMsg & "(" & Format(lngInactive, "#,##0") & " inactive Task(s) not included in total.)"
  End If
  
  MsgBox strMsg, vbInformation + vbOKOnly, "Task Counter"
  
exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Set oTask = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCountTasks_bas", "cptCountTasks", Err, Erl)
  Resume exit_here
End Sub

Sub cptCountTasksAll()
  Call cptCountTasks("All")
End Sub

Sub cptCountTasksSelected()
  Call cptCountTasks("Selected")
End Sub

Sub cptCountTasksVisible()
  Call cptCountTasks("Visible")
End Sub
