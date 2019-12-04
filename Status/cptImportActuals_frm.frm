VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptImportActuals_frm 
   Caption         =   "Import Actuals"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   OleObjectBlob   =   "cptImportActuals_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptImportActuals_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboTask_Change()
  If Not IsNull(Me.cboTask.Value) Then
    Me.txtUID.Value = Me.cboTask.Value
  End If
End Sub

Private Sub cmdAssignTask_Click()
'objects
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ensure a task is selected
  If IsNull(Me.cboTask) Then
    MsgBox "Please select a valid task.", vbExclamation + vbOKOnly, "Import Actuals"
    Me.cboTask.SetFocus
    Me.cboTask.DropDown
    GoTo exit_here
  End If
  
  'capture the task uid
  For lngItem = 0 To Me.lboMap.ListCount - 1
    If Me.lboMap.Selected(lngItem) Then
      Me.lboMap.List(lngItem, 2) = Me.txtUID
      Me.lboMap.List(lngItem, 3) = Me.cboTask.List(, 1)
    End If
  Next lngItem

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_frm", "cmdAssignTask_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdCancel_Click()
  'todo: kill ini and csv in temp directory
  'todo: kill the task-map adtg
  Unload Me
End Sub

Private Sub cmdClearAll_Click()
  Call cptClearMappedTasks
End Sub

Private Sub cmdClearSelected_Click()
  Call cptClearMappedTasks(False)
End Sub

Private Sub optExistingTasks_Click()
  Me.lboMap.Enabled = Me.optExistingTasks
  Me.txtUID.Enabled = Me.optExistingTasks
  Me.txtSearch.Enabled = Me.optExistingTasks
  Me.cboTask.Enabled = Me.optExistingTasks
  Me.cmdClearAll.Enabled = Me.optExistingTasks
  Me.cmdClearSelected.Enabled = Me.optExistingTasks
  Me.cmdAssignTask.Enabled = Me.optExistingTasks
End Sub

Private Sub optNewTasks_Click()
  Me.lboMap.Enabled = Not Me.optNewTasks
  Me.txtUID.Enabled = Not Me.optNewTasks
  Me.txtSearch.Enabled = Not Me.optNewTasks
  Me.cboTask.Enabled = Not Me.optNewTasks
  Me.cmdClearAll.Enabled = Not Me.optNewTasks
  Me.cmdClearSelected.Enabled = Not Me.optNewTasks
  Me.cmdAssignTask.Enabled = Not Me.optNewTasks
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
  Cancel = True
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  Call cptListWPCN(Node)
End Sub

Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, _
                                  Effect As Long, Button As Integer, _
                                  Shift As Integer, x As Single, y As Single)
  Call cptAddFilesActuals(Data)
End Sub

Private Sub txtSearch_Change()
  If Me.ActiveControl <> Me.txtSearch Then Exit Sub
  If Len(Me.txtSearch.Text) > 0 Then
    Call cptUpdateTaskMapList(Me.txtSearch.Text)
  Else
    Call cptUpdateTaskMapList
  End If
End Sub

Private Sub txtSearch_Enter()
  Call cptUpdateTaskMapList
End Sub

Private Sub txtUID_Change()
'objects
Dim Task As Task
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Me.ActiveControl <> Me.txtUID Then GoTo exit_here

  If Len(Me.txtUID.Text) > 0 Then
    On Error Resume Next
    Set Task = ActiveProject.Tasks.UniqueID(Me.txtUID.Text)
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Not Task Is Nothing Then
      Me.cboTask.Value = Me.txtUID.Text
    End If
  Else
    Me.cboTask.Value = Null
  End If

exit_here:
  On Error Resume Next
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_frm", "txtUID_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtUID_Enter()
  Me.txtSearch = ""
  Call cptUpdateTaskMapList
End Sub

Private Sub UserForm_Initialize()
  Me.TreeView1.OLEDropMode = ccOLEDropManual
End Sub

Private Sub UserForm_Terminate()
  If Dir(cptDir & "\settings\cpt-actuals-map.adtg") <> vbNullString Then
    Kill cptDir & "\settings\cpt-actuals-map.adtg"
  End If
  If Dir(Environ("temp") & "\Schema.ini") <> vbNullString Then
    Kill Environ("temp") & "\Schema.ini"
  End If
  'todo: kill temp csv files by name - when? on form close_click
End Sub
