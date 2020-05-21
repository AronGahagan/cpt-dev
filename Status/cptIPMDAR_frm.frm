VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptIPMDAR_frm 
   Caption         =   "Create IPMDAR Schedule Performance Dataset (SPD)"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11310
   OleObjectBlob   =   "cptIPMDAR_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptIPMDAR_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboContract_Change()
'objects
Dim oSubDir As Object
Dim oRootDir As Object
Dim oFSO As Object
Dim aSubmittals As Object
'strings
'longs
Dim lngPeriod As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'load cboPrevDirectories
  'get list of directories
  Set aSubmittals = CreateObject("System.Collections.SortedList")
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oRootDir = oFSO.GetFolder(Environ("USERPROFILE") & "\IPMDAR\" & Me.cboContract.Value)
  For Each oSubDir In oRootDir.SubFolders
    'todo: sort prev dir by datecreated or alphabetically? user chooses?
    If Mid(oSubDir.Path, InStrRev(oSubDir.Path, "\") + 1) <> Format(ActiveProject.StatusDate, "yyyy-mm-dd") Then  'todo: exclude current period dir
      aSubmittals.Add oSubDir.DateCreated, Mid(oSubDir.Path, InStrRev(oSubDir.Path, "\") + 1)
    End If
    'todo: what if future periods exist for project? A: let the user be smart about this
  Next
  'list the previous periods in descending order
  cptIPMDAR_frm.cboPrevDir.Clear
  For lngPeriod = aSubmittals.Count To 1 Step -1
    cptIPMDAR_frm.cboPrevDir.AddItem aSubmittals.getByIndex(lngPeriod - 1)
  Next lngPeriod
  If aSubmittals.Count > 0 Then 'default to most previous
    cptIPMDAR_frm.cboPrevDir.Value = aSubmittals.getByIndex(aSubmittals.Count - 1)
  Else
    'todo: no prior periods exist yet
  End If
  Me.cboPrevDir.ControlTipText = "Matches subdirectories in C:\Users\[username]\IPMDAR\" & Me.cboContract.Value & "\"
exit_here:
  On Error Resume Next
  Set oSubDir = Nothing
  Set oRootDir = Nothing
  Set oFSO = Nothing
  Set aSubmittals = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_frm", "cboContract_Change", Err, Erl)
  
  Resume exit_here
End Sub

Private Sub cboPhysicalPercentComplete_Change()
  Me.txtA_PhysicalPercentComplete.Value = Me.cboPhysicalPercentComplete.Value
End Sub

Private Sub cboResourceID_Change()
  Me.txtA_ResourceID.Value = "[Resource]" & Me.cboResourceID.Value
End Sub

Private Sub cboTaskID_Change()
  Me.txtA_TaskID.Value = "[Task]" & Me.cboTaskID.Value
End Sub

Private Sub cmdCreate_Click()
  'todo: parse all files, regex for forbidden values, trim, report
End Sub

Private Sub cmdLoad_Click()
  Call cptLoadCOBRAData
End Sub

Private Sub cmdRequestCOBRAData_Click()
  Call cptRequestCOBRAData
End Sub

Private Sub cmdReset_Click()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.txtSchema.Value = "IPMDAR_SCHEDULE_PERFORMANCE_DATASET/1.0"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_frm", "cmdReset_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboCalendars_AfterUpdate()
  Me.txtCalendarComments = Me.lboCalendars.List(Me.lboCalendars.ListIndex, 3)
End Sub

Private Sub lboFiles_AfterUpdate()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
Dim vCal As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.mpOptions.Value = Me.lboFiles.ListIndex
  Select Case Me.lboFiles.Value
    Case "Calendars.json"
      Call cptLoadCalendars
      
    Case Else
    
  End Select

  Me.lboFiles.SetFocus

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_frm", "lboFiles_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub optOutlineCode_Click()
  Me.cboOutlineCode.Enabled = Me.optOutlineCode
  Me.cboOutlineCode.SetFocus
  Me.cboOutlineCode.DropDown
End Sub

Private Sub optSummaryTasks_Click()
  Me.cboOutlineCode.Enabled = Not Me.optSummaryTasks
End Sub

Private Sub txtCalendarComments_Change()
  Me.lboCalendars.List(Me.lboCalendars.ListIndex, 3) = Me.txtCalendarComments
  If Len(Me.txtCalendarComments) > 0 Then
    Me.lboCalendars.List(Me.lboCalendars.ListIndex, 2) = "+"
  Else
    Me.lboCalendars.List(Me.lboCalendars.ListIndex, 2) = "x"
  End If
  'todo: need to save calendar comments somewhere as edited somewhere too -- automatically edit the json?
End Sub

Private Sub UserForm_Initialize()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
Dim vPage As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For Each vPage In Array(1, 9)
    Me.mpOptions.Pages(vPage).ScrollBars = fmScrollBarsVertical
    Me.mpOptions.Pages(vPage).KeepScrollBarsVisible = fmScrollBarsVertical
    If vPage = 1 Then
      Me.mpOptions.Pages(vPage).ScrollHeight = 2.2 * Me.mpOptions.Height
    ElseIf vPage = 9 Then
      Me.mpOptions.Pages(vPage).ScrollHeight = 1.5 * Me.mpOptions.Height
    End If
  Next

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_frm", "UserForm_Initialize", Err, Erl)
  Resume exit_here
End Sub
