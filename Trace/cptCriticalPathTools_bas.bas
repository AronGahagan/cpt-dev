Attribute VB_Name = "cptCriticalPathTools_bas"
'<cpt_version>v1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
'revised 2017-12-14

Sub ExportCriticalPath(ByRef Project As Project, Optional blnSendEmail = False, Optional blnKeepOpen = False, Optional ByRef TargetTask As Task)
'objects
Dim Task As Task, Tasks As Tasks ', TargetTask As Task
Dim pptApp As PowerPoint.Application, Presentation As PowerPoint.Presentation, Slide As PowerPoint.Slide
Dim Shape As PowerPoint.Shape
Dim ShapeRange As PowerPoint.ShapeRange
'strings
Dim strFileName As String, strMsg As String, strProjectName As String, strDir As String
'longs
Dim lgT1Milestone As Long, lgDrivingPath As Long, lgL2Milestone As Long
Dim lgTask As Long, lgTasks As Long, lgSlide As Long
'dates
Dim dtOriginalDeadline As Date, dtFrom As Date, dtTo As Date
'boolean
Dim blnFoundIt As Boolean
'variants
Dim vPath As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not ModuleExists("ClearPlan_CritPathModule") Then
    MsgBox "Please install the ClearPlan Critical Path Module.", vbCritical + vbOKOnly, "CP Toolbar"
    GoTo exit_here
  End If
  
  Call DrivingPaths
    
  Set pptApp = CreateObject("PowerPoint.Application")
  pptApp.Visible = True
  Set Presentation = pptApp.Presentations.Add(msoCTrue)
  
  'ensure directory
  strDir = Environ("USERPROFILE") & "\Desktop\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'make a project subdirectory
  'strDir = strDir & strProjectName & "\"
  'If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'make a date subdirectory
  'strDir = strDir & Format(Now(), "yyyy-mm-dd") & "\"
  'If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir

  strFileName = strDir & Replace(Replace(Project.Name, " ", "-"), ".mpp", "") & "-CriticalPathAnalysis-" & Format(Now, "yyyy-mm-dd") & ".pptx"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  Presentation.SaveAs strFileName
  'make a title slide
  Set Slide = Presentation.Slides.Add(1, ppLayoutCustom)
  Slide.Layout = ppLayoutTitle
  Slide.Shapes(1).TextFrame.TextRange.Text = strProjectName & vbCrLf & "Critical Path Analysis"
  Slide.Shapes(2).TextFrame.TextRange.Text = GetUserFullName & vbCrLf & Format(Now, "mm/dd/yyyy") 'Project.ProjectSummaryTask.GetField(FieldNameToFieldConstant("E2E Scheduler"))
  'for each primary,secondary,tertiary > make a slide
  For Each vPath In Array("1", "2", "3")
    'copy the picture
    SetAutoFilter FieldName:="CP Driving Paths", FilterType:=pjAutoFilterCustom, Test1:="contains", Criteria1:=CStr(vPath)
    Sort Key1:="Finish", Key2:="Duration", Ascending2:=False, Renumber:=False
    'SelectAll
    'ZoomTimescale Selection:=True
    'TimescaleEdit MajorUnits:=0, MajorLabel:=0, MajorCount:=1, MinorCount:=3, MinorTicks:=True, Separator:=True, TierCount:=2
    TimescaleEdit MajorUnits:=0, MinorUnits:=2, MajorLabel:=0, MinorLabel:=10, MinorTicks:=True, Separator:=True, TierCount:=2
    SelectBeginning
    If Not IsDate(Project.StatusDate) Then
      dtFrom = DateAdd("d", -14, Project.ProjectStart)
    Else
      dtFrom = DateAdd("d", -14, Project.StatusDate)
    End If
    dtTo = DateAdd("d", 30, TargetTask.Finish)
    Debug.Print vPath & ": " & FormatDateTime(dtFrom, vbShortDate) & " - " & FormatDateTime(dtTo, vbShortDate)
    SelectAll
    'account for when a path is somehow not found
    On Error Resume Next
    Set Tasks = ActiveSelection.Tasks
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Tasks Is Nothing Then GoTo next_path
    'account for when task count exceeds easily visible range
    'on powerpoint slide
    lgTasks = Tasks.count
    lgSlide = 0
    lgTask = 0
    Do While lgTask <= lgTasks
      lgTask = lgTask + 20
      lgSlide = lgSlide + 1
      SelectBeginning
      SelectTaskField Row:=lgTask - 20, Column:="Name", Height:=20, Extend:=False
      EditCopyPicture Object:=False, ForPrinter:=0, SelectedRows:=1, FromDate:=Format(dtFrom, "mm/dd/yy hh:nn AMPM"), ToDate:=Format(dtTo, "m/d/yy hh:mm ampm"), ScaleOption:=pjCopyPictureShowOptions, MaxImageHeight:=-1#, MaxImageWidth:=-1#, MeasurementUnits:=2
      'paste the picture
      Presentation.Slides.Add Presentation.Slides.count + 1, ppLayoutCustom
      Set Slide = Presentation.Slides(Presentation.Slides.count)
      Slide.Layout = ppLayoutChart
      Slide.Shapes(1).TextFrame.TextRange.Text = Choose(vPath, "Primary", "Secondary", "Tertiary") & " Critical Path" & IIf(lgSlide > 1, " (cont'd)", "")
      Slide.Shapes(2).Delete
      Slide.Shapes.Paste
      Slide.Shapes(Slide.Shapes.count).Width = Slide.Master.Width * 0.9
      Slide.Shapes(Slide.Shapes.count).Left = (Slide.Master.Width / 2) - (Slide.Shapes(Slide.Shapes.count).Width / 2)
      If Slide.Shapes(Slide.Shapes.count).Top <> 108 Then Slide.Shapes(Slide.Shapes.count).Top = 108
    Loop
    Presentation.Save
next_path:
    Set Tasks = Nothing
  Next vPath
  SetAutoFilter "CP Driving Paths"
  SelectBeginning
  If Not Presentation.Saved Then Presentation.Save
  If Not blnKeepOpen Then
    Presentation.Close
    pptApp.Quit
  End If
  
exit_here:
  On Error Resume Next
  Calculation = pjAutomatic
  Set TargetTask = Nothing
  Set Task = Nothing
  Set pptApp = Nothing
  Set Presentation = Nothing
  Set Slide = Nothing
  Exit Sub
  
err_here:
  Call HandleErr("basCriticalPathTools", "ExportCriticalPath", err)
  Resume exit_here
End Sub

Sub ExportCriticalPathAll()
Dim Project As Project
  
  For Each Project In Projects
    If InStr(Project.Name, "E2E") > 0 Or InStr(Project.Name, "RNG") > 0 Then Call ExportCriticalPath(Project)
  Next Project
  
  Set Project = Nothing
  
  MsgBox "Complete.", vbInformation + vbOKOnly, "Critical Path Analysis"
  
End Sub

Sub ExportCriticalPathActive()
  Call ExportCriticalPath(ActiveProject, blnKeepOpen:=True)
End Sub

Sub ExportCriticalPathSelected()
Dim TargetTask As Task

  On Error GoTo err_here

  Set TargetTask = ActiveCell.Task

  Call ExportCriticalPath(ActiveProject, blnKeepOpen:=True, TargetTask:=ActiveSelection.Tasks(1))
  
exit_here:
  On Error Resume Next
  Set TargetTask = Nothing
  Exit Sub
err_here:
  If err.Number = 1101 Then
    MsgBox "Please a a single (non-summary, active, and incomplete) 'Target' task.", vbExclamation + vbOKOnly, "Trace Tools - Error"
  Else
    Call HandleErr("basCriticalPathTools", "ExportCriticalPathSelected", err)
  End If
  Resume exit_here
End Sub

Sub InstantCriticalPath()
Dim dtFinish As Date, dtDeadline As Date
Dim Task As Task

  On Error GoTo err_here
  If ActiveSelection.Tasks.count > 1 Then GoTo exit_here
  SpeedON
  Set Task = ActiveSelection.Tasks(1)
  If Task.PercentComplete = 100 Then GoTo exit_here
  If IsDate(Task.Deadline) Then
    InputBox "Deadline already exists. Copy this to paste back in afterward:", "Instant Critical Path", Task.Deadline
  End If
  FilterClear
  Task.Deadline = DateAdd("yyyy", -2, Task.Finish)
  CalculateProject
  OptionsViewEx projectsummary:=False, displaynameindent:=False, displaysummarytasks:=False
  SetAutoFilter "Total Slack", FilterType:=pjAutoFilterIn, Criteria1:=Task.GetField(pjTaskTotalSlack)
  Sort "Finish", True, "Duration", False
  EditGoTo Task.ID
  
exit_here:
  On Error Resume Next
  Set Task = Nothing
  SpeedOFF
  Exit Sub
err_here:
  Call HandleErr("basCriticalPathTools", "InstantCriticalPath", err)
  Resume exit_here
End Sub

Sub InstantCriticalPathClear()
  SetTaskField Field:="Deadline", Value:=""
  FilterClear
End Sub

Sub ResetView()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.ScreenUpdating = False
  ViewApply "Gantt Chart"
  ActiveWindow.TopPane.Activate
  FilterClear
  GroupClear
  OptionsViewEx displayoutlinesymbols:=True, displaynameindent:=True, displaysummarytasks:=True
  OutlineShowAllTasks
  Application.ScreenUpdating = True

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("basCriticalPathTools", "ResetView", err)
  Resume exit_here
End Sub
