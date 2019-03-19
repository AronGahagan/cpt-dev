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

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptTextTools_frm.Visible Then
    If Not IsError(ActiveSelection.Tasks) Then
      cptTextTools_frm.lboOutput.Clear
      For Each Task In ActiveSelection.Tasks
        cptTextTools_frm.lboOutput.AddItem
        cptTextTools_frm.lboOutput.List(lgItem, 0) = Task.UniqueID
        cptTextTools_frm.lboOutput.List(lgItem, 1) = Task.Name
        lgItem = lgItem + 1
      Next Task
      Call cptTextTools_bas.UpdatePreview
    End If
  End If
  
'  On Error Resume Next
'  Set Tasks = ActiveSelection.Tasks
'  If Not Tasks Is Nothing Then
' ' If Not IsError(ActiveSelection.Tasks) Then
'    lgTasks = ActiveSelection.Tasks.count
'    Application.StatusBar = Format(lgTasks, "#,##0") & " task" & IIf(lgTasks = 1, "", "s") & " selected"
'  End If
  
exit_here:
  On Error Resume Next
  Set Tasks = Nothing
  Exit Sub
err_here:
  Call HandleErr("clsChange", "MyMSApplication_WindowSelectionChange", err)
  Resume exit_here
  
End Sub
