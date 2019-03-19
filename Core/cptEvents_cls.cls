VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cptEvents_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<cpt_version>v0.1</cpt_version>

Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private WithEvents MyMSPApplication As Application
Attribute MyMSPApplication.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set MyMSPApplication = Application

End Sub

Private Sub MyMSPApplication_WindowSelectionChange(ByVal Window As Window, ByVal sel As Selection, ByVal selType As Variant)
Dim Tasks As Tasks, Task As Task, lgItem As Long, lgTasks As Long
Dim frmText As Object

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If ModuleExists("cptTextTools_frm") Then
    If Not IsError(ActiveSelection.Tasks) Then
      Set frmText = ThisProject.VBProject.VBComponents("cptTextTools_frm")
      If Not frmText Is Nothing And frmText.Visible Then
        frmText.lboOutput.Clear
        For Each Task In ActiveSelection.Tasks
          frmText.lboOutput.AddItem
          frmText.lboOutput.List(lgItem, 0) = Task.UniqueID
          frmText.lboOutput.List(lgItem, 1) = Task.Name
          lgItem = lgItem + 1
        Next Task
        Call ThisProject.VBProject.VBComponents("cptTextTools_bas").UpdatePreview
      End If
    End If
  End If
  
  On Error Resume Next
  Debug.Print ActiveWindow.ActivePane.View.Type
  Select Case ActiveWindow.ActivePane.View.Type
    Case 0 'Task Table
    Case 1
  End Select
  Set Tasks = ActiveSelection.Tasks
  If Not Tasks Is Nothing Then
    lgTasks = ActiveSelection.Tasks.count
    Application.StatusBar = Format(lgTasks, "#,##0") & " task" & IIf(lgTasks = 1, "", "s") & " selected"
  Else
    Application.StatusBar = "0 tasks selected"
  End If
  
exit_here:
  On Error Resume Next
  Set frmText = Nothing
  Set Tasks = Nothing
  Exit Sub
err_here:
  Call HandleErr("clsChange", "MyMSApplication_WindowSelectionChange", err)
  Resume exit_here
  
End Sub
