VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cptEvents_cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<cpt_version>v1.3.0</cpt_version>
Option Explicit
Private WithEvents MyMSPApplication As Application
Attribute MyMSPApplication.VB_VarHelpID = -1

Private Sub Class_Initialize()

    Set MyMSPApplication = Application

End Sub

Private Sub MyMSPApplication_WindowSelectionChange(ByVal Window As Window, ByVal sel As Selection, ByVal selType As Variant)
'objects
Dim ctl As Object
Dim frmText As Object 'UserForm
Dim oUserForm As Object 'UserForm
Dim Tasks As Tasks, Task As Task, Resources As Resources
'strings
'longs
Dim lngItem As Long, lngItems As Long
'integers
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  On Error Resume Next
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type = pjTaskItem Then
    If Not IsError(ActiveSelection.Tasks) Then
      Set frmText = cptGetUserForm("cptText_frm") 'returns Nothing if form is not loaded
      Set ctl = cptGetControl(frmText, "lboOutput")
      If Not frmText Is Nothing Then
        frmText.Controls("lboOutput").Clear
        For Each Task In ActiveSelection.Tasks
          frmText.Controls("lboOutput").AddItem
          frmText.Controls("lboOutput").List(lngItem, 0) = Task.UniqueID
          frmText.Controls("lboOutput").List(lngItem, 1) = Task.Name
          lngItem = lngItem + 1
        Next Task
        Application.Run "cptUpdatePreview"
      End If
    End If
  End If
  
  On Error Resume Next
  
  'Validate users selected view type
  If cptGetShowStatusBarCountFirstRun Then
    If ActiveProject.Application.ActiveWindow.ActivePane.View.Type = pjTaskItem Then
      Set Tasks = ActiveSelection.Tasks
      If Not Tasks Is Nothing Then
        lngItems = ActiveSelection.Tasks.Count
        Application.StatusBar = Format(lngItems, "#,##0") & " task" & IIf(lngItems = 1, "", "s") & " selected"
      Else
        Application.StatusBar = "0 tasks selected"
      End If
    ElseIf ActiveProject.Application.ActiveWindow.ActivePane.View.Type = pjResourceItem Then
      Set Resources = ActiveSelection.Resources
      If Not Resources Is Nothing Then
        lngItems = ActiveSelection.Resources.Count
        Application.StatusBar = Format(lngItems, "#,##0") & " resource" & IIf(lngItems = 1, "", "s") & " selected"
      Else
        Application.StatusBar = "0 resources selected"
      End If
    End If
  End If
  
  Set oUserForm = cptGetUserForm("cptNetworkBrowser_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptShowPreds"
 
  Set oUserForm = cptGetUserForm("cptTaskHistory_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptUpdateTaskHistory"
 
  Set oUserForm = cptGetUserForm("cptSmartDuration_frm") 'returns Nothing if form is not loaded
  If Not oUserForm Is Nothing Then Application.Run "cptUpdateSmartDurationForm"
  
exit_here:
  On Error Resume Next
  Set oUserForm = Nothing
  Set ctl = Nothing
  Set frmText = Nothing
  Set Tasks = Nothing
  Set Resources = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("clsChange_cls", "MyMSApplication_WindowSelectionChange", Err, Erl)
  Resume exit_here
  
End Sub
