VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptImportActuals_frm 
   Caption         =   "Import Actuals"
   ClientHeight    =   6810
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
  If Not IsNull(Me.cboTask.Value) Then Me.txtUID.Text = Me.cboTask.Value
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

Private Sub optExistingTasks_Click()
  Me.lboMap.Enabled = Me.optExistingTasks
  Me.txtUID.Enabled = Me.optExistingTasks
  Me.txtSearch.Enabled = Me.optExistingTasks
  Me.cboTask.Enabled = Me.optExistingTasks
End Sub

Private Sub optNewTasks_Click()
  Me.lboMap.Enabled = Not Me.optNewTasks
  Me.txtUID.Enabled = Not Me.optNewTasks
  Me.txtSearch.Enabled = Not Me.optNewTasks
  Me.cboTask.Enabled = Not Me.optNewTasks
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
  If Len(Me.txtSearch.Text) > 0 Then
    Call cptUpdateTaskMapList(Me.txtSearch.Text)
  Else
    Call cptUpdateTaskMapList
  End If
End Sub

Private Sub UserForm_Initialize()
  Me.TreeView1.OLEDropMode = ccOLEDropManual
End Sub

Private Sub UserForm_Terminate()
  If Dir(cptDir & "\settings\cpt-actuals-map.adtg") <> vbNullString Then
    Kill cptDir & "\settings\cpt-actuals-map.adtg"
  End If
  'todo: kill ini
  'todo: kill temp csv files by name
End Sub
