Attribute VB_Name = "cptCriticalPathTools_bas"
'<cpt_version>v1.0.4</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptExportCriticalPath(ByRef Project As Project, Optional blnSendEmail As Boolean = False, Optional blnKeepOpen As Boolean = False, Optional ByRef TargetTask As Task)
'objects
Dim oShell As Object
Dim pptExists As PowerPoint.Presentation
Dim Task As Task, Tasks As Tasks
Dim pptApp As PowerPoint.Application, Presentation As PowerPoint.Presentation, Slide As PowerPoint.Slide
'Dim Shape As PowerPoint.Shape
'Dim ShapeRange As PowerPoint.ShapeRange
'strings
Dim strFileName As String, strProjectName As String, strDir As String
'longs
Dim lgTask As Long, lgTasks As Long, lgSlide As Long
'dates
Dim dtFrom As Date, dtTo As Date
'boolean
'variants
Dim vPath As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptModuleExists("cptCriticalPath_bas") Then
    MsgBox "Please install the ClearPlan Critical Path Module.", vbCritical + vbOKOnly, "CP Toolbar"
    GoTo exit_here
  End If
  
  cptSpeed True
  
  export_to_PPT = True
  Call DrivingPaths
  export_to_PPT = False
  
  If Not IsDate(Project.StatusDate) Then
    dtFrom = DateAdd("d", -14, Project.ProjectStart)
  Else
    dtFrom = DateAdd("d", -14, Project.StatusDate)
  End If
  dtTo = DateAdd("d", 30, TargetTask.Finish)

  EditGoTo Date:=dtFrom
  
  Set pptApp = CreateObject("PowerPoint.Application")
  pptApp.Visible = True
  Set Presentation = pptApp.Presentations.Add(msoCTrue)
  
  'ensure directory
  Set oShell = CreateObject("WScript.Shell")
  strDir = oShell.SpecialFolders("Desktop") & "\"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'build filename
  strFileName = strDir & Replace(Replace(Project.Name, " ", "-"), ".mpp", "") & "-CriticalPathAnalysis-" & Format(Now, "yyyy-mm-dd") & ".pptx"
  On Error Resume Next
  Set pptExists = pptApp.Presentations(strFileName)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not pptExists Is Nothing Then 'add timestamp to this file
    pptExists.Save
    pptExists.Close
  End If
  'might exist but be closed
  If Dir(strFileName) <> vbNullString Then
    If MsgBox(strFileName & " exists. Overwrite?", vbExclamation + vbYesNo, "File Exists") = vbYes Then
      Kill strFileName
    Else
      MsgBox "The presentation you are creating will have a time stamp in the filename to prevent overwriting.", vbInformation + vbOKOnly, "File Name Changed"
      strFileName = Replace(strFileName, ".mpp", "-" & Format(Now, "hh-nn-ss") & ".mpp")
    End If
  Else
    
  End If
  Presentation.SaveAs strFileName
  'make a title slide
  Set Slide = Presentation.Slides.Add(1, ppLayoutCustom)
  Slide.Layout = ppLayoutTitle
  strProjectName = Replace(ActiveProject.Name, ".mpp", "")
  Slide.Shapes(1).TextFrame.TextRange.Text = strProjectName & vbCrLf & "Critical Path Analysis"
  Slide.Shapes(2).TextFrame.TextRange.Text = cptGetUserFullName & vbCrLf & Format(Now, "mm/dd/yyyy") 'Project.ProjectSummaryTask.GetField(FieldNameToFieldConstant("E2E Scheduler"))
  
  'for each primary,secondary,tertiary > make a slide
  For Each vPath In Array("1", "2", "3")
    'copy the picture
    'SetAutoFilter FieldName:="CP Driving Paths", FilterType:=pjAutoFilterCustom, Test1:="contains", Criteria1:=CStr(vPath)
    SetAutoFilter FieldName:="CP Driving Path Group ID", FilterType:=pjAutoFilterIn, Criteria1:=CStr(vPath)

    Sort key1:="Finish", Key2:="Duration", Ascending2:=False, Renumber:=False
    TimescaleEdit MajorUnits:=0, MinorUnits:=2, MajorLabel:=0, MinorLabel:=10, MinorTicks:=True, Separator:=True, TierCount:=2
    SelectBeginning
    Debug.Print vPath & ": " & FormatDateTime(dtFrom, vbShortDate) & " - " & FormatDateTime(dtTo, vbShortDate)
    SelectAll
    'account for when a path is somehow not found
    On Error Resume Next
    Set Tasks = ActiveSelection.Tasks
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Tasks Is Nothing Then GoTo next_path
    'account for when task count exceeds easily visible range
    'on powerpoint slide
    lgTasks = Tasks.Count
    lgSlide = 0
    lgTask = 0
    Do While lgTask <= lgTasks
      lgTask = lgTask + 20
      lgSlide = lgSlide + 1
      SelectBeginning
      SelectTaskField Row:=lgTask - 20, Column:="Name", Height:=20, Extend:=False
      EditCopyPicture Object:=False, ForPrinter:=0, SelectedRows:=1, FromDate:=Format(dtFrom, "mm/dd/yy hh:nn AMPM"), ToDate:=Format(dtTo, "m/d/yy hh:mm ampm"), ScaleOption:=pjCopyPictureTimescale, MaxImageHeight:=-1#, MaxImageWidth:=-1#, MeasurementUnits:=2  'pjCopyPictureShowOptions
      'paste the picture
      Presentation.Slides.Add Presentation.Slides.Count + 1, ppLayoutCustom
      Set Slide = Presentation.Slides(Presentation.Slides.Count)
      Slide.Layout = ppLayoutChart
      Slide.Shapes(1).TextFrame.TextRange.Text = Choose(vPath, "Primary", "Secondary", "Tertiary") & " Critical Path" & IIf(lgSlide > 1, " (cont'd)", "")
      Slide.Shapes(2).Delete
      Slide.Shapes.Paste
      Slide.Shapes(Slide.Shapes.Count).Width = Slide.Master.Width * 0.9
      Slide.Shapes(Slide.Shapes.Count).Left = (Slide.Master.Width / 2) - (Slide.Shapes(Slide.Shapes.Count).Width / 2)
      If Slide.Shapes(Slide.Shapes.Count).Top <> 108 Then Slide.Shapes(Slide.Shapes.Count).Top = 108
    Loop
    Presentation.Save
next_path:
    Set Tasks = Nothing
  Next vPath
  SetAutoFilter "CP Driving Path Group ID"
  SelectBeginning
  If Not Presentation.Saved Then Presentation.Save
  
  MsgBox "Critical Path slides created.", vbInformation + vbOKOnly, "Complete"
  
  pptApp.Activate
  
exit_here:
  On Error Resume Next
  Set oShell = Nothing
  cptSpeed False
  Set pptExists = Nothing
  Set TargetTask = Nothing
  Set Task = Nothing
  Set pptApp = Nothing
  Set Presentation = Nothing
  Set Slide = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptCriticalPathTools_bas", "cptExportCriticalPath", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCriticalPathSelected()
'objects
Dim TargetTask As Task

  On Error GoTo err_here

  Set TargetTask = ActiveCell.Task

  Call cptExportCriticalPath(ActiveProject, blnKeepOpen:=True, TargetTask:=ActiveSelection.Tasks(1))
  
exit_here:
  On Error Resume Next
  Set TargetTask = Nothing
  Exit Sub
err_here:
  If Err.Number = 1101 Then
    MsgBox "Please a a single (non-summary, active, and incomplete) 'Target' task.", vbExclamation + vbOKOnly, "Trace Tools - Error"
  Else
    Call cptHandleErr("cptCriticalPathTools_bas", "cptExportCriticalPathSelected", Err, Erl)
  End If
  Resume exit_here
End Sub

Sub cptDrivingPath()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  singlePath = True
  Call DrivingPaths

exit_here:
  On Error Resume Next
  singlePath = False
  Exit Sub
err_here:
  Call cptHandleErr("cptCriticalPathTools_bas", "cptDrivingPath", Err, Erl)
  Resume exit_here
End Sub
