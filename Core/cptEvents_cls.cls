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
'objects
Dim Tasks As Tasks, Task As Task, Resources As Resources, Resource As Resource
'strings
'longs
Dim lngItem As Long, lngItems As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  On Error Resume Next
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type = pjTaskItem Then
    If ModuleExists("cptText_frm") Then
      If Not IsError(ActiveSelection.Tasks) Then
        If cptText_frm.Visible Then
          cptText_frm.lboOutput.Clear
          For Each Task In ActiveSelection.Tasks
            cptText_frm.lboOutput.AddItem
            cptText_frm.lboOutput.List(lngItem, 0) = Task.UniqueID
            cptText_frm.lboOutput.List(lngItem, 1) = Task.Name
            lngItem = lngItem + 1
          Next Task
          Call ThisProject.VBProject.VBComponents("cptText_bas").UpdatePreview
        End If
      End If
    End If
  End If
  
  On Error Resume Next
  
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type = pjTaskItem Then
    Set Tasks = ActiveSelection.Tasks
    If Not Tasks Is Nothing Then
      lngItems = ActiveSelection.Tasks.count
      Application.StatusBar = Format(lngItems, "#,##0") & " task" & IIf(lngItems = 1, "", "s") & " selected"
    Else
      Application.StatusBar = "0 tasks selected"
    End If
  ElseIf ActiveProject.Application.ActiveWindow.ActivePane.View.Type = pjResourceItem Then
    Set Resources = ActiveSelection.Resources
    If Not Resources Is Nothing Then
      lngItems = ActiveSelection.Tasks.count
      Application.StatusBar = Format(lngItems, "#,##0") & " resources" & IIf(lngItems = 1, "", "s") & " selected"
    Else
      Application.StatusBar = "0 resources selected"
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set Tasks = Nothing
  Set Resources = Nothing
  Exit Sub
err_here:
  Call HandleErr("clsChange_cls", "MyMSApplication_WindowSelectionChange", err)
  Resume exit_here
  
End Sub
