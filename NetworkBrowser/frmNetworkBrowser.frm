VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNetworkBrowser 
   Caption         =   "Dependency Browser"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   OleObjectBlob   =   "frmNetworkBrowser.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNetworkBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdBack_Click()

  On Error GoTo err_here
  
  Me.lboHistory.SetFocus
  
  If IsNull(Me.lboHistory.Value) Then Me.lboHistory.ListIndex = -1

  If Me.lboHistory.ListCount > 0 Then
    Me.lboHistory.ListIndex = Me.lboHistory.ListIndex + 1
    Call HistoryDoubleClick
  End If

exit_here:
  Exit Sub
err_here:
  If err.Number = 380 Then MsgBox "No more history.", vbInformation, "The End"
  Resume exit_here
  
End Sub

Private Sub cmdClearHistory_Click()
  If Me.lboHistory.ListCount > 0 Then
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "Confirm") = vbYes Then Me.lboHistory.Clear
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFwd_Click()

  On Error GoTo err_here
  
  Me.lboHistory.SetFocus
  
  If IsNull(Me.lboHistory.Value) Then Me.lboHistory.ListIndex = 0

  If Me.lboHistory.ListCount > 0 Then
    Me.lboHistory.ListIndex = Me.lboHistory.ListIndex - 1
    Call HistoryDoubleClick
  End If

exit_here:
  Exit Sub
err_here:
  Resume exit_here

End Sub

Private Sub cmdMark_Click()
Dim lngID As Long

  Application.Calculation = pjManual
  If ActiveSelection.Tasks.count = 1 Then
    lngID = ActiveSelection.Tasks(1).ID
    If Not ActiveSelection.Tasks(1).Marked Then ActiveSelection.Tasks(1).Marked = True
    For i = 0 To Me.lboPredecessors.ListCount - 1
      If Me.lboPredecessors.Selected(i) Then
        If Me.lboPredecessors.Column(0, i) = "ID" Then GoTo exit_here
        ActiveProject.Tasks(CLng(Me.lboPredecessors.Column(0, i))).Marked = True
      End If
    Next i
    For i = 0 To Me.lboSuccessors.ListCount - 1
      If Me.lboSuccessors.Selected(i) Then
        If Me.lboSuccessors.Column(0, i) = "ID" Then GoTo exit_here
        ActiveProject.Tasks(CLng(Me.lboSuccessors.Column(0, i))).Marked = True
      End If
    Next i
  Else
    MsgBox "Please select only one task.", vbInformation + vbOKOnly, "Error"
    Exit Sub
  End If
  ActiveWindow.TopPane.Activate
  FilterApply "Marked"
  'Sort "Total Slack", False, "Start", True
  Sort "Start", True, "Duration", True
  SelectAll
  'GoTo exit_here
  SelectAll
  If ActiveWindow.BottomPane Is Nothing Then
    Application.DetailsPaneToggle False
  End If
  
  ActiveWindow.BottomPane.Activate
  If ActiveWindow.BottomPane.View.Name <> "Network Diagram" Then ViewApply "Network Diagram"
  EditGoTo lngID
exit_here:
  Application.Calculation = pjAutomatic
End Sub

Private Sub cmdRefresh_Click()
  Call ShowPreds
End Sub

Private Sub cmdUnmark_Click()
Dim lngID As Long

  Application.Calculation = pjManual
  If ActiveSelection.Tasks.count = 1 Then
    lngID = ActiveSelection.Tasks(1).ID
    For i = 0 To Me.lboPredecessors.ListCount - 1
      If Me.lboPredecessors.Selected(i) Then
        ActiveProject.Tasks(CLng(Me.lboPredecessors.Column(0, i))).Marked = False
      End If
    Next i
    For i = 0 To Me.lboSuccessors.ListCount - 1
      If Me.lboSuccessors.Selected(i) Then
        ActiveProject.Tasks(CLng(Me.lboSuccessors.Column(0, i))).Marked = False
      End If
    Next i
  Else
    MsgBox "Please select only one task.", vbInformation + vbOKOnly, "Error"
    Exit Sub
  End If
  ActiveWindow.TopPane.Activate
  FilterApply "Marked"
  Sort "Start", True, "Duration", True
  SelectAll
  SelectAll
  ActiveWindow.BottomPane.Activate
  If ActiveWindow.BottomPane Is Nothing Then
    Application.DetailsPaneToggle False
  End If
  ActiveWindow.BottomPane.Activate
  If ActiveWindow.BottomPane.View.Name <> "Network Diagram" Then ViewApply "Network Diagram"
  EditGoTo lngID
exit_here:
  Application.Calculation = pjAutomatic
End Sub

Private Sub cmdUnmarkAll_Click()
Dim Task As Task

  cptSpeed True
  ActiveWindow.BottomPane.Activate
  On Error Resume Next
  If ActiveSelection.Tasks.count = 0 Then Exit Sub
  For Each Task In ActiveSelection.Tasks
    Task.Marked = False
  Next Task
  ActiveWindow.TopPane.Activate
  FilterApply "Marked"
  Sort "Start", True, "Duration", True
  SelectAll
  ActiveWindow.BottomPane.Activate
  cptSpeed False
  
End Sub

Sub lboHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call HistoryDoubleClick
End Sub

Sub lboPredecessors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim lngTaskID As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With Me.lboHistory
    .AddItem ActiveSelection.Tasks.Item(1).ID, 0
  End With
  lngTaskID = CLng(Me.lboPredecessors.Column(0))
  If lngTaskID > 0 Then
    WindowActivate TopPane:=True
    EditGoTo lngTaskID, ActiveProject.Tasks(lngTaskID).Start
    Me.lboHistory.AddItem lngTaskID, 0
    Me.lboHistory.ListIndex = Me.lboHistory.TopIndex
    Call ShowPreds
  End If
  
exit_here:
  Exit Sub
err_here:
  If err.Number = 1101 Then
    Call RemoveFilters(lngTaskID)
    Resume exit_here
  End If
  Call cptHandleErr("cptNetworkBrowser_frm", "lboPredecesors_DblClick", err, Erl)
  Resume exit_here
End Sub

Private Sub lboSuccessors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim lngTaskID As Long, Task As Task

  On Error Resume Next
  Set Task = ActiveSelection.Tasks(1)

  On Error GoTo err_here
  
  With Me.lboHistory
    If Not Task Is Nothing Then .AddItem Task.ID, 0
  End With
  lngTaskID = CLng(Me.lboSuccessors.Column(0))
  If lngTaskID > 0 Then
    WindowActivate TopPane:=True
    EditGoTo lngTaskID, ActiveProject.Tasks(lngTaskID).Start
    Me.lboHistory.AddItem lngTaskID, 0
    Me.lboHistory.ListIndex = Me.lboHistory.TopIndex
    Call ShowPreds
  End If
  
exit_here:
  Exit Sub
err_here:
  If err.Number = 1101 Then
    Call RemoveFilters(lngTaskID)
    Resume exit_here
  End If
  Call cptHandleErr("cptNetworkBrowser_frm", "lboSuccessors_DblClick", err, Erl)
  Resume exit_here
End Sub

Private Sub tglTrace_Click()
  If Not Me.tglTrace Then
    Me.tglTrace.Caption = "Jump"
    Me.cmdMark.Enabled = False
    Me.cmdUnmark.Enabled = False
    Me.lboPredecessors.MultiSelect = fmMultiSelectSingle
    Me.lboSuccessors.MultiSelect = fmMultiSelectSingle
  Else
    Me.tglTrace.Caption = "Trace"
    Me.cmdMark.Enabled = True
    Me.cmdUnmark.Enabled = True
    Me.lboPredecessors.MultiSelect = fmMultiSelectMulti
    Me.lboSuccessors.MultiSelect = fmMultiSelectMulti
  End If
End Sub

Private Sub UserForm_Initialize()

End Sub
