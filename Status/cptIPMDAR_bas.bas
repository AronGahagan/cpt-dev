Attribute VB_Name = "cptIPMDAR_bas"
'<cpt_version>0.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
'for x = 0 to cptIPMDAR_frm.lboFiles.ListCount-1 : debug.Print (x+1) & " - " & cptIPMDAR_frm.lboFiles.List(x,0) : next x

Sub cptJSON_Main()
'objects
Dim aRecords As Object
Dim oStream As Stream
Dim oFSO As Scripting.FileSystemObject
Dim oSubProject As MSProject.SubProject
Dim oSourceProject As MSProject.Project
Dim aProjects As Object
Dim oProject As MSProject.Project
'strings
Dim strBuffer As String
Dim strTemp As String
Dim strErr As String
Dim strDir As String
'longs
Dim lngTemp As Long
Dim lngProject As Long
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
Dim vRecord As Variant
Dim vFile As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'USER ACTIONS BEFORE RUNNING:
  '-- Identify Critical Path
  '-- Identify Driving Path(s)
  '-- Add custom field descriptions using Data Dictionary feature
  
  'USER ACTIONS AFTER RUNNING:
  '-- Use IPMDAR_DATA_REVIEW TO:
  '--- ensure unique task names
  '--- ensure unique calendars (CalendarID)
  '--- ensure unique workshifts (CalendarID + Ordinal)
  '--- save offline copy of schedule
  '--- zip files

  Set oProject = ActiveProject
      
  'ensure status date
  If Not IsDate(oProject.StatusDate) Then
    MsgBox "Please provide a status date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    GoTo exit_here
  End If
  
  'confirm root IPMDAR directory
  strDir = Environ("USERPROFILE") & "\IPMDAR"
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'confirm contract directory
  strDir = strDir & "\" & cptIPMDAR_frm.cboContract.Value
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
  'confirm period directory
  strDir = strDir & "\" & Format(oProject.StatusDate, "yyyy-mm-dd")
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir

  'create the FileType.txt
  lngFile = FreeFile
  Open strDir & "\FileType.txt" For Output As #lngFile
  Print #lngFile, "IPMDAR_SCHEDULE_PERFORMANCE_DATASET/1.0"
  Close #lngFile
  
  'create the exports

  'DatasetMetadata.json is created on COBRA data load
  If Dir(strDir & "\DatasetMetadata.json") = vbNullString Then
    strErr = "DatasetMetadata.json" & vbCrLf
  End If

  'SourceSoftwareMetadata.json
  If Not cptJSON_SourceSoftwareMetadata(strDir) Then
    strErr = strErr & "SourceSoftwareMetadata.json" & vbCrLf
  End If
  
  'todo: ProjectCustomFieldDefinitions.json
  'todo: what if these don't match between subprojects? only grab this from the master project?
  'ProjectScheduleData.json
  If Not cptJSON_ProjectScheduleData(oProject, strDir) Then 'includes ProjectCustomFieldValues
    strErr = strErr & "ProjectScheduleData.json" & vbCrLf
  End If
  'todo: ProjectCustomFieldValues.json
  'todo: TaskCustomFieldDefinitions.json - should come from a master project
  'todo: ResourceCustomFieldDefinitions.json - should come from a master project
  
  'issue calls for code effeciency
  Set aProjects = CreateObject("System.Collections.ArrayList")
  If oProject.Subprojects.Count > 0 Then
    'ensure task view
    If ActiveWindow.TopPane.View.Type <> pjTaskItem Then
      ActiveWindow.TopPane.Activate
      ViewApply "Gantt Chart"
    End If
    SelectAll
    OutlineShowAllTasks
    aProjects.Add oProject
    For Each oSubProject In oProject.Subprojects
      aProjects.Add oSubProject.SourceProject
    Next 'oSourceProject
  Else
    aProjects.Add oProject
  End If
  
  'Tasks.json (includes TaskScheduleData, TaskCustomFieldValues, TaskConstraints, TaskRelationships,TaskOutlineStructure)
  If Not cptJSON_Tasks(oProject, strDir) Then
    strErr = strErr & "Tasks.json" & vbCrLf
  End If
  'todo: TaskCustomFieldValues
  
  'overwrite existing files by default
  'do this before looping through subprojects
  If Dir(strDir & "\Calendars.json") <> vbNullString Then Kill strDir & "\Calendars.json"
  If Dir(strDir & "\CalendarWorkshifts.json") <> vbNullString Then Kill strDir & "\CalendarWorkshifts.json"
  If Dir(strDir & "\CalendarExceptions.json") <> vbNullString Then Kill strDir & "\CalendarExceptions.json"
  If Dir(strDir & "\Resources.json") <> vbNullString Then Kill strDir & "\Resources.json"
  If Dir(strDir & "\ResourceAssignments.json") <> vbNullString Then Kill strDir & "\ResourceAssignments.json"
  For lngProject = 0 To aProjects.Count - 1
    Set oSourceProject = aProjects.Item(lngProject)
    
    'Calendars.json
    If Not cptJSON_Calendars(oSourceProject, strDir) Then 'includes CalendarWorkshifts, CalendarExceptions
      strErr = strErr & "Calendars.json" & vbCrLf
    End If
    
    'Resources.json (includes ResourceCustomFieldValues, ResourceAssignments)
    If Not cptJSON_Resources(oSourceProject, strDir) Then
      strErr = "Resources.json" & vbCrLf
    End If
    'todo: ResourceCustomFieldValues
  
  Next lngProject
  
  'clean-up the files
  For Each vFile In Array("Calendars", "CalendarWorkshifts", "CalendarExceptions", "Resources", "ResourceAssignments")
    lngTemp = FreeFile
    strTemp = ""
    Open strDir & "\" & vFile & ".json" For Input As #lngTemp
    Do While Not EOF(lngTemp)
      Line Input #lngTemp, strBuffer
      strTemp = strTemp & strBuffer
    Loop
    Close #lngTemp
    strTemp = Replace(strTemp, "][", ",")
    strTemp = Replace(Replace(strTemp, "[", ""), "]", "")
    Set aRecords = CreateObject("System.Collections.ArrayList")
    For Each vRecord In Split(strTemp & ",", "},")
      If Not aRecords.Contains(vRecord & "}") And Len(vRecord) > 0 Then aRecords.Add (vRecord & "}")
    Next vRecord
    lngTemp = FreeFile
    Open strDir & "\" & vFile & ".json" For Output As #lngTemp
    strTemp = Join(aRecords.ToArray, ",")
    'strTemp = Left(strTemp, Len(strTemp) - 1)
    Print #lngTemp, "[" & strTemp & "]"
    Close #lngTemp
  Next
  
  'todo: scrub for character limitations SPD FFS 2.1.6
    'cycle through the json files
    'if tripped then
      'split into array
      'check each line and highlight it somehow
      'or simply remove it
    'end if
  'todo: ensure all string values are trimmed
  'todo: remove all instances of more than one space
  'todo: remove tab characters
  'todo: create schedule narrative template containing:
  'todo: -- section headers [created with the ClearPlan toolbar
  'todo: -- placeholders for all explanations for leads, lags, constraints
  'todo: -- placeholders for CWBS, SOW if exist in Outline Code
  'todo: -- export IMS Data Dictionary and paste into narrative
  'todo: prompt to save server file as .mpp and consolidate?
  
  'prompt user
  If MsgBox("Create IPMDAR Data Review workbook?", vbQuestion + vbYesNo) = vbYes Then
    Call cptCreateIPMDARWorkbook(strDir)
  End If
    
  If Len(strErr) > 0 Then
    MsgBox "The following exports were not created successfully:" & vbCrLf & strErr, vbExclamation + vbOKOnly, "Incomplete"
  Else
    MsgBox "Schedule Performance Data exported correctly.", vbInformation + vbOKOnly, "SPD"
  End If

exit_here:
  On Error Resume Next
  Set aRecords = Nothing
  Set oStream = Nothing
  Set oFSO = Nothing
  Set oSubProject = Nothing
  Set oSourceProject = Nothing
  Set aProjects = Nothing
  Set oProject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSAON_Main", Err, Erl)
  Resume exit_here
End Sub

Function cptJSON_DatasetMetadata(ByRef oProject As Project, strDir As String) As Boolean
'objects
'strings
Dim strEOC As String
Dim strJSON As String
Dim strFile As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  'this is loaded directly from COBRA
  GoTo exit_here

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    
  strFile = strDir & "\DatasetMetadata.json"
  lngFile = FreeFile
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngFile

  strJSON = "[{"
'  SecurityMarking
'  DistributionStatement
'  ReportingPeriodEndDate =
  strJSON = strJSON & Chr(34) & "ReportingPeriodEndDate" & Chr(34) & ": " & Chr(34) & Format(oProject.StatusDate, "yyyy-mm-dd") & Chr(34) & ","
'  ContractorName
'  ContractorIDCodeTypeID
'  ContractorIDCode
'  ContractorAddress_Street
'  ContractorAddress_City
'  ContractorAddress_State
'  ContractorAddress_Country
'  ContractorAddress_ZipCode
'  PointOfContactName
'  PointOfContactTitle
'  PointOfContactTelephone
'  PointOfContactEmail
'  ContractName
'  ContractNumber
'  ContractType
'  ContractTaskOrEffortName
'  ProgramName
'  ProgramPhase
'  EVMSAccepted
'  EVMSAcceptanceDate
  
  strJSON = strJSON & "}]"
  
  Print #lngFile, strJSON
  cptJSON_DatasetMetadata = True

exit_here:
  On Error Resume Next
  Close #lngFile
  Exit Function
err_here:
  'Call HandleErr("cptIPMDAR", "cptJSON_DatasetMetadata", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  cptJSON_DatasetMetadata = False
  Resume exit_here
End Function

Function cptJSON_SourceSoftwareMetadata(strDir As String) As Boolean
'objects
'strings
Dim strVersion As String
Dim strJSON As String
Dim strFile As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strFile = strDir & "\SourceSoftwareMetadata.json"
  lngFile = FreeFile
  Open strFile For Output As #lngFile
  
  strJSON = "[{"
  strJSON = strJSON & Chr(34) & "Data_SoftwareName" & Chr(34) & ": " & Chr(34) & Application.Name & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "Data_SoftwareVersion" & Chr(34) & ": " & Chr(34) & Application.Version & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "Data_SoftwareCompanyName" & Chr(34) & ": " & Chr(34) & "Microsoft Corporation" & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "Data_SoftwareComments" & Chr(34) & ": " & Chr(34) & "Build " & Application.Build & " running on " & Application.OperatingSystem & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "Export_SoftwareName" & Chr(34) & ": " & Chr(34) & "ClearPlan Toolbar" & Chr(34) & ","
  'get current version
  strVersion = Replace(Replace(cptRegEx(ThisProject.VBProject.VBComponents("cptIPMDAR").CodeModule.Lines(1, 3), "<.*>"), "<cpt_version>", ""), "</cpt_version>", "")
  strJSON = strJSON & Chr(34) & "Export_SoftwareVersion" & Chr(34) & ": " & Chr(34) & "cptIPMDAR v" & strVersion & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "Export_SoftwareCompanyName" & Chr(34) & ": " & Chr(34) & "ClearPlan Consulting, LLC" & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "Export_SoftwareComments" & Chr(34) & ": null"
  strJSON = strJSON & "}]"

  Print #lngFile, strJSON
  cptJSON_SourceSoftwareMetadata = True

exit_here:
  On Error Resume Next
  Close #lngFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_SourceSoftwareMetadata", Err, Erl)
  cptJSON_SourceSoftwareMetadata = False
  Resume exit_here
End Function

Function cptJSON_ProjectScheduleData(ByRef oProject As Project, strDir As String) As Boolean
'objects
'strings
Dim strJSON As String
Dim strFile As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ProjectScheduleData.json
  strFile = strDir & "\ProjectScheduleData.json"
  lngFile = FreeFile
  Open strFile For Output As #lngFile

  'todo: ProjectCustomFieldValues.json

  strJSON = "[{"
  strJSON = strJSON & Chr(34) & "StatusDate" & Chr(34) & ": " & Chr(34) & Format(oProject.StatusDate, "yyyy-mm-dd") & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "CurrentStartDate" & Chr(34) & ": " & Chr(34) & Format(oProject.ProjectSummaryTask.Start, "yyyy-mm-dd") & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "CurrentFinishDate" & Chr(34) & ": " & Chr(34) & Format(oProject.ProjectSummaryTask.Finish, "yyyy-mm-dd") & Chr(34) & ","
  If IsDate(oProject.BaselineSavedDate(pjBaseline)) Then
    strJSON = strJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": " & Chr(34) & Format(oProject.ProjectSummaryTask.BaselineStart, "yyyy-mm-dd") & Chr(34) & ","
    strJSON = strJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": " & Chr(34) & Format(oProject.ProjectSummaryTask.BaselineFinish, "yyyy-mm-dd") & Chr(34) & ","
  Else
    strJSON = strJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": null,"
    strJSON = strJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": null,"
  End If
  If IsDate(oProject.ProjectSummaryTask.ActualStart) Then
    strJSON = strJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": " & Chr(34) & Format(oProject.ProjectSummaryTask.ActualStart, "yyyy-mm-dd") & Chr(34) & ","
  Else
    strJSON = strJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": null,"
  End If
  If IsDate(oProject.ProjectSummaryTask.ActualFinish) Then
    strJSON = strJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": " & Chr(34) & Format(oProject.ProjectSummaryTask.ActualFinish, "yyyy-mm-dd") & Chr(34) & ","
  Else
    strJSON = strJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": null,"
  End If
  strJSON = strJSON & Chr(34) & "DurationUnitsID" & Chr(34) & ": " & Chr(34) & "DAYS" & Chr(34) 'can be DAYS or HOURS
  strJSON = strJSON & "}]"

  Print #lngFile, strJSON
  cptJSON_ProjectScheduleData = True

exit_here:
  On Error Resume Next
  Close #lngFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_ProjectScheduleData", Err, Erl)
  cptJSON_ProjectScheduleData = False
  Resume exit_here
End Function

Function cptJSON_ProjectCustomFieldDefinitions(ByRef oProject As Project, strDir As String) As Boolean

End Function

Function cptJSON_Calendars(ByRef oProject As Project, strDir As String) As Boolean
'objects
Dim aWorkHours As Object
Dim oCalendarException As MSProject.Exception
Dim oWorkshift As MSProject.WorkWeek
Dim oCalendar As MSProject.Calendar
'strings
Dim strCalendarExceptionsJSON As String
Dim strCalendarExceptionsFile As String
Dim strCalendarWorkshiftsJSON As String
Dim strCalendarWorkshiftsFile As String
Dim strCalendarsJSON As String
Dim strCalendarsFile As String
'longs
Dim lngWorkHours As Long
Dim lngDayOfWeek As Long
Dim lngCalendarExceptionsFile As Long
Dim lngCalendarWorkshiftsFile As Long
Dim lngCalendarsFile As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtException As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'Calendars.json
  strCalendarsFile = strDir & "\Calendars.json"
  lngCalendarsFile = FreeFile
  Open strCalendarsFile For Append As #lngCalendarsFile
  
  'CalendarWorkshifts.json
  strCalendarWorkshiftsFile = strDir & "\CalendarWorkshifts.json"
  lngCalendarWorkshiftsFile = FreeFile
  Open strCalendarWorkshiftsFile For Append As #lngCalendarWorkshiftsFile
  
  'CalendarExceptions.json
  strCalendarExceptionsFile = strDir & "\CalendarExceptions.json"
  lngCalendarExceptionsFile = FreeFile
  Open strCalendarExceptionsFile For Append As #lngCalendarExceptionsFile
  
  'todo: calendar loop needs to match cptLoadCalendars()
  For Each oCalendar In oProject.BaseCalendars 'only used calendars
    'todo: what about resource calendars? all exceptions or only exceptions to base calendar?
  
    'Calendars.json
    strCalendarsJSON = strCalendarsJSON & "{"
    strCalendarsJSON = strCalendarsJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oCalendar.Guid & Chr(34) & ","
    strCalendarsJSON = strCalendarsJSON & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & oCalendar.Name & Chr(34) & ","
    strCalendarsJSON = strCalendarsJSON & Chr(34) & "Comments" & Chr(34) & ": null},"
    
    'CalendarWorkshifts.json
    For Each oWorkshift In oCalendar.WorkWeeks
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & "{"
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oCalendar.Guid & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "Ordinal" & Chr(34) & ": " & Chr(34) & oWorkshift.Index & Chr(34) & ","
      Set aWorkHours = CreateObject("System.Collections.SortedList")
      For lngDayOfWeek = 1 To 7
        With oWorkshift.WeekDays(lngDayOfWeek)
          lngWorkHours = DateDiff("n", CDate(.Shift1.Start), CDate(.Shift1.Finish), vbSunday) / 60
          lngWorkHours = lngWorkHours + DateDiff("n", CDate(.Shift2.Start), CDate(.Shift2.Finish), vbSunday) / 60
          lngWorkHours = lngWorkHours + DateDiff("n", CDate(.Shift3.Start), CDate(.Shift3.Finish), vbSunday) / 60
          lngWorkHours = lngWorkHours + DateDiff("n", CDate(.Shift4.Start), CDate(.Shift4.Finish), vbSunday) / 60
          lngWorkHours = lngWorkHours + DateDiff("n", CDate(.Shift5.Start), CDate(.Shift5.Finish), vbSunday) / 60
          aWorkHours.Add lngDayOfWeek, lngWorkHours
        End With
      Next 'lngDayOfWeek
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "SundayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(0) & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "MondayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(1) & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "TuesdayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(2) & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "WednesdayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(3) & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "ThursdayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(4) & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "FridayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(5) & Chr(34) & ","
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "SaturdayWorkHours" & Chr(34) & ": " & Chr(34) & aWorkHours.getByIndex(6) & Chr(34)
      strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & "},"
    Next 'oWorkshift
    
    'CalendarExceptions.json
    For Each oCalendarException In oCalendar.Exceptions
      dtException = oCalendarException.Start
      Do While dtException <= oCalendarException.Finish
        If Weekday(dtException) > 1 And Weekday(dtException) < 7 Then
          strCalendarExceptionsJSON = strCalendarExceptionsJSON & "{"
          strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oCalendar.Guid & Chr(34) & ","
          strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "ExceptionDate" & Chr(34) & ": " & Chr(34) & Format(dtException, "yyyy-mm-dd") & Chr(34) & ","
          'todo: what about WorkHours?
          strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "WorkHours" & Chr(34) & ": null"
          strCalendarExceptionsJSON = strCalendarExceptionsJSON & "},"
        End If
        dtException = DateAdd("d", 1, dtException)
      Loop
    Next 'oCalendarException
    
  Next 'oCalendar
  
  'Calendars.json
  strCalendarsJSON = "[" & Left(strCalendarsJSON, Len(strCalendarsJSON) - 1) & "]"
  'todo: add both brackets at this step if len(x)>0 otherwise don't create it
  Print #lngCalendarsFile, strCalendarsJSON
  
  'CalendarWorkshifts.json
  If Len(strCalendarWorkshiftsJSON) > 0 Then
    strCalendarWorkshiftsJSON = "[" & Left(strCalendarWorkshiftsJSON, Len(strCalendarWorkshiftsJSON) - 1) & "]"
    Print #lngCalendarWorkshiftsFile, strCalendarWorkshiftsJSON
  End If
  
  'CalendarExceptions.json
  If Len(strCalendarExceptionsJSON) > 0 Then
    strCalendarExceptionsJSON = "[" & Left(strCalendarExceptionsJSON, Len(strCalendarExceptionsJSON) - 1) & "]"
    Print #lngCalendarExceptionsFile, strCalendarExceptionsJSON
  End If
  
  cptJSON_Calendars = True

exit_here:
  On Error Resume Next
  Set aWorkHours = Nothing
  Set oCalendarException = Nothing
  Set oWorkshift = Nothing
  Set oCalendar = Nothing
  Close #lngCalendarsFile
  Close #lngCalendarWorkshiftsFile
  Close #lngCalendarExceptionsFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_Calendars", Err, Erl)
  cptJSON_Calendars = False
  Resume exit_here
End Function

Function cptJSON_Tasks(ByRef oProject As Project, strDir As String) As Boolean
'objects
Dim oTaskDependency As MSProject.TaskDependency
Dim oTask As MSProject.Task
'strings
Dim strTaskType As String
Dim strTasksJSON As String
Dim strTasksFile As String
Dim strTaskOutlineStructureJSON As String
Dim strTaskOutlineStructureFile As String
Dim strTaskScheduleDataJSON As String
Dim strTaskScheduleDataFile As String
Dim strTaskConstraintsJSON As String
Dim strTaskConstraintsFile As String
Dim strTaskRelationshipsFile As String
Dim strTaskRelationshipsJSON As String
'longs
Dim lngOutlineOffset As Long
Dim lngTasksFile As Long
Dim lngTaskOutlineStructureFile As Long
Dim lngTaskScheduleDataFile As Long
Dim lngTaskConstraintsFile As Long
Dim lngTaskRelationshipsFile As Long
'integers
'doubles
'booleans
Dim blnDisplayProjectSummary As Boolean
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: identify driving path(s) based on previous month's uid/guid, to either a) next contractor identified event; or b) gov selected event
  'todo: identify critical path(s)
  
  'TaskOutlineStructure.json
  strTaskOutlineStructureFile = strDir & "\TaskOutlineStructure.json"
  lngTaskOutlineStructureFile = FreeFile
  If Dir(strTaskOutlineStructureFile) <> vbNullString Then Kill strTaskOutlineStructureFile
  Open strTaskOutlineStructureFile For Output As #lngTaskOutlineStructureFile
  
  'Tasks.json
  strTasksFile = strDir & "\Tasks.json"
  lngTasksFile = FreeFile
  If Dir(strTasksFile) <> vbNullString Then Kill strTasksFile
  Open strTasksFile For Output As #lngTasksFile
  
  'TaskScheduleData.json
  strTaskScheduleDataFile = strDir & "\TaskScheduleData.json"
  lngTaskScheduleDataFile = FreeFile
  If Dir(strTaskScheduleDataFile) <> vbNullString Then Kill strTaskScheduleDataFile
  Open strTaskScheduleDataFile For Output As #lngTaskScheduleDataFile
  
  'todo: TaskCustomFieldValues.json
  
  'TaskConstraints.json
  strTaskConstraintsFile = strDir & "\TaskConstraints.json"
  lngTaskConstraintsFile = FreeFile
  If Dir(strTaskConstraintsFile) <> vbNullString Then Kill strTaskConstraintsFile
  Open strTaskConstraintsFile For Output As #lngTaskConstraintsFile
  
  'todo: TaskRelationships.json
  strTaskRelationshipsFile = strDir & "\TaskRelationships.json"
  lngTaskRelationshipsFile = FreeFile
  If Dir(strTaskRelationshipsFile) <> vbNullString Then Kill strTaskRelationshipsFile
  Open strTaskRelationshipsFile For Output As #lngTaskRelationshipsFile
   
  'todo: use a single text field to house string array of identification codes?
  
  'todo: account for elapsed durations
  
  For Each oTask In oProject.Tasks
    If Not oTask Is Nothing Then
              
      'build TaskOutlineStructure.json
      strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "{"
      strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "Level" & Chr(34) & ": " & Chr(34) & CLng(oTask.OutlineLevel) & Chr(34) & ","
      strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
      If oTask.OutlineLevel > 1 Then
        strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "ParentTaskID" & Chr(34) & ": " & Chr(34) & oTask.OutlineParent.Guid & Chr(34) & ","
      Else
        strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "ParentTaskID" & Chr(34) & ": null,"
      End If
      strTaskOutlineStructureJSON = Left(strTaskOutlineStructureJSON, Len(strTaskOutlineStructureJSON) - 1)
      strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "},"
            
      'build Tasks.json
      strTasksJSON = strTasksJSON & "{"
      strTasksJSON = strTasksJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
      strTasksJSON = strTasksJSON & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & Trim(oTask.Name) & Chr(34) & ","
      'TaskTypeID: ACTIVITY,MILESTONE,SUMMARY,HAMMOCK
      If oTask.Summary Then
        strTaskType = "SUMMARY"
      ElseIf oTask.Milestone Then
        strTaskType = "MILESTONE"
      ElseIf oTask.Duration > 0 Then
        strTaskType = "ACTIVITY"
      End If
      strTasksJSON = strTasksJSON & Chr(34) & "TaskTypeID" & Chr(34) & ": " & Chr(34) & strTaskType & Chr(34) & ","
      'TaskSubtypeID: RISK_MITIGATION_TASK,SCHEDULE_VISIBILITY_TASK,SCHEDULE_MARGIN,CONTRACTUAL_MILESTONE
      If strTaskType = "ACTIVITY" Then
        strTasksJSON = strTasksJSON & Chr(34) & "TaskPlanningLevelID" & Chr(34) & ": " & Chr(34) & strTaskType & Chr(34) & ","
        'TaskPlanningLevelID: SUMMARY_LEVEL_PLANNING_PACKAGE,CONTROL_ACCOUNT,PLANNING_PACKAGE,WORK_PACKAGE,ACTIVITY
      End If 'strTaskType = "ACTIVITY"
      'WBSElementID:
      'OBSElementID:
      'ControlAccountID:
      'WorkPackageID:
      'IMPElementID:
      'SOWReference:
      'SubcontractorReference:
      'EarnedValueTechniqueID: APPORTIONED_EFFORT,LEVEL_OF_EFFORT,MILESTONE,FIXED_0_100,FIXED_100_0,FIXED_X_Y,PERCENT_COMPLETE,STANDARDS,UNITS,OTHER_DISCRETE
      'OtherEarnedValueTechnique: (must be null unless EarnedValueTechniqueID = OTHER_DISCRETE or FIXED_X_Y)
      'SourceSubprojectReference:
      strTasksJSON = strTasksJSON & Chr(34) & "SourceSubprojectReference" & Chr(34) & ": " & Chr(34) & oTask.Project & Chr(34) & ","
      'SourceTaskReference:
      strTasksJSON = strTasksJSON & Chr(34) & "SourceTaskReference" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
      'Comments:
      strTasksJSON = Left(strTasksJSON, Len(strTasksJSON) - 1)
      strTasksJSON = strTasksJSON & "},"
      
      'build TaskScheduleData.json
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & "{"
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
      If oTask.Calendar = "None" Then
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oProject.Calendar.Guid & Chr(34) & ","
      Else
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oTask.CalendarObject.Guid & Chr(34) & ","
      End If
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentDuration" & Chr(34) & ": " & Chr(34) & oTask.Duration / (oProject.HoursPerDay * 60) & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyStart, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyFinish, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateStart, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateFinish, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FreeFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.FreeSlack / (oProject.HoursPerDay * 60) & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "TotalFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.TotalSlack / (oProject.HoursPerDay * 60) & Chr(34) & ","
      'todo: need flags for on critical path
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnCriticalPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
      'todo: need flags for on driving path
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnDrivingPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineDuration" & Chr(34) & ": " & Chr(34) & oTask.BaselineDuration / (oProject.HoursPerDay * 60) & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineStart, "yyyy-mm-dd") & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineFinish, "yyyy-mm-dd") & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "StartVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.StartVariance / (oProject.HoursPerDay * 60) & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FinishVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.FinishVariance / (oProject.HoursPerDay * 60) & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalculatedPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PercentComplete & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "PhysicalPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PhysicalPercentComplete & Chr(34) & ","
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "RemainingDuration" & Chr(34) & ": " & Chr(34) & oTask.RemainingDuration / (oProject.HoursPerDay * 60) & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualStart, "yyyy-mm-dd") & Chr(34) & ","
      'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualFinish, "yyyy-mm-dd") & Chr(34) & ","
      strTaskScheduleDataJSON = Left(strTaskScheduleDataJSON, Len(strTaskScheduleDataJSON) - 1)
      strTaskScheduleDataJSON = strTaskScheduleDataJSON & "},"
      
      'build TaskConstraints.json
      If oTask.ConstraintType <> pjASAP Then
        strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
        
        Select Case oTask.ConstraintType

          Case pjALAP
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "AS_LATE_AS_POSSIBLE" & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": null"

          Case pjMSO
            If oProject.HonorConstraints Then
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "MUST_START_ON" & Chr(34) & ","
            Else
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "SHOULD_START_ON" & Chr(34) & ","
            End If
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)

          Case pjMFO
            If oProject.HonorConstraints Then
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "MUST_FINISH_ON" & Chr(34) & ","
            Else
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "SHOULD_FINISH_ON" & Chr(34) & ","
            End If
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)

          Case pjSNET
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "START_NO_EARLIER_THAN" & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)

          Case pjSNLT
            If oProject.HonorConstraints Then
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "START_NO_LATER_THAN" & Chr(34) & ","
            Else
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "SHOULD_START_NO_LATER_THAN" & Chr(34) & ","
            End If
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)

          Case pjFNET
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "FINISH_NO_EARLER_THAN" & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)

          Case pjFNLT
            If oProject.HonorConstraints Then
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "FINISH_NO_LATER_THAN" & Chr(34) & ","
            Else
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "SHOULD_FINISH_NO_LATER_THAN" & Chr(34) & ","
            End If
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)
          
        End Select

        strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
      
      End If 'oTask.ConstraintType <> pjASAP
      
      'TaskConstraints.json - resource leveling delay
      If oTask.LevelingDelay > 0 Then 'new record
        
        'resource leveling start delay
        strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "RESOURCE_LEVELING_START_DELAY" & Chr(34) & ","
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34)
        strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
        
        'resource leveling finish delay
        strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "RESOURCE_LEVELING_FINISH_DELAY" & Chr(34) & ","
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34)
        strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
        
      End If 'oTask.LevelingDelay
      
      'TaskConstraints.json - deadline
      If IsDate(oTask.Deadline) Then 'new record
        strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.Guid & Chr(34) & ","
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "DEADLINE" & Chr(34) & ","
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
        strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Deadline, "yyyy-mm-dd") & Chr(34)
        strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
      End If 'IsDate(oTask.Deadline)
      
      'build TaskRelationships.json
      For Each oTaskDependency In oTask.TaskDependencies
        strTaskRelationshipsJSON = strTaskRelationshipsJSON & "{"
        strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "PredecessorTaskID" & Chr(34) & ": " & Chr(34) & oTaskDependency.From.Guid & Chr(34) & ","
        strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "SuccessorTaskID" & Chr(34) & ": " & Chr(34) & oTaskDependency.To.Guid & Chr(34) & ","
        Select Case oTaskDependency.Type
          Case pjFinishToFinish
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "RelationshipTypeID" & Chr(34) & ": " & Chr(34) & "FINISH_TO_FINISH" & Chr(34) & ","
          Case pjFinishToStart
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "RelationshipTypeID" & Chr(34) & ": " & Chr(34) & "FINISH_TO_START" & Chr(34) & ","
          Case pjStartToFinish
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "RelationshipTypeID" & Chr(34) & ": " & Chr(34) & "START_TO_FINISH" & Chr(34) & ","
          Case pjStartToStart
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "RelationshipTypeID" & Chr(34) & ": " & Chr(34) & "START_TO_START" & Chr(34) & ","
        End Select
        strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagDuration" & Chr(34) & ": " & Chr(34) & oTaskDependency.Lag / (oProject.HoursPerDay * 60) & Chr(34) & ","
        If oTaskDependency.To.Calendar <> "None" Then
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagCalendarID" & Chr(34) & ": " & Chr(34) & oTaskDependency.To.CalendarObject.Guid & Chr(34)
        Else
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagCalendarID" & Chr(34) & ": " & Chr(34) & oProject.Calendar.Guid & Chr(34)
        End If
        strTaskRelationshipsJSON = strTaskRelationshipsJSON & "},"
      Next 'oTaskDependency
      
    End If 'Not oTask Is Nothing Then
next_task:
  Next 'oTask
    
  'create Tasks.json
  strTasksJSON = "[" & Left(strTasksJSON, Len(strTasksJSON) - 1) & "]"
  Print #lngTasksFile, strTasksJSON
  
  'create TaskOutlineStructure.json
  strTaskOutlineStructureJSON = "[" & Left(strTaskOutlineStructureJSON, Len(strTaskOutlineStructureJSON) - 1) & "]"
  Print #lngTaskOutlineStructureFile, strTaskOutlineStructureJSON
  
  'create TaskScheduleData.json
  strTaskScheduleDataJSON = "[" & Left(strTaskScheduleDataJSON, Len(strTaskScheduleDataJSON) - 1) & "]"
  Print #lngTaskScheduleDataFile, strTaskScheduleDataJSON
  
  'create TaskConstraints.json
  strTaskConstraintsJSON = "[" & Left(strTaskConstraintsJSON, Len(strTaskConstraintsJSON) - 1) & "]"
  Print #lngTaskConstraintsFile, strTaskConstraintsJSON
  
  'create TaskRelationships.json
  strTaskRelationshipsJSON = "[" & Left(strTaskRelationshipsJSON, Len(strTaskRelationshipsJSON) - 1) & "]"
  Print #lngTaskRelationshipsFile, strTaskRelationshipsJSON
  
  cptJSON_Tasks = True

exit_here:
  On Error Resume Next
  Set oTaskDependency = Nothing
  Set oTask = Nothing
  Close #lngTasksFile
  Close #lngTaskOutlineStructureFile
  Close #lngTaskScheduleDataFile
  Close #lngTaskConstraintsFile
  Close #lngTaskRelationshipsFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_Tasks", Err, Erl)
  cptJSON_Tasks = False
  Resume exit_here
End Function

Function cptJSON_Resources(ByRef oProject As Project, strDir As String) As Boolean
'objects
Dim oAssignment As MSProject.Assignment
Dim oResource As MSProject.Resource
'strings
Dim strResourceAssignmentsJSON As String
Dim strResourceAssignmentsFile As String
Dim strEOC As String
Dim strResourcesJSON As String
Dim strResourcesFile As String
'longs
Dim lngResourceAssignmentsFile As Long
Dim lngResourcesFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
   
  'Resources.json
  strResourcesFile = strDir & "\Resources.json"
  lngResourcesFile = FreeFile
  Open strResourcesFile For Append As #lngResourcesFile
  
  'ResourceAssignments.json
  strResourceAssignmentsFile = strDir & "\ResourceAssignments.json"
  lngResourceAssignmentsFile = FreeFile
  Open strResourceAssignmentsFile For Append As #lngResourceAssignmentsFile
  
  'todo: ResourceCustomFieldValues.json
  
  For Each oResource In oProject.Resources
  
    'Resources.json
    strResourcesJSON = strResourcesJSON & "{"
    strResourcesJSON = strResourcesJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oResource.Guid & Chr(34) & ","
    strResourcesJSON = strResourcesJSON & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & Trim(oResource.Name) & Chr(34) & ","
    Select Case oResource.Type
      Case pjResourceTypeWork
        strEOC = "LABOR"
      Case Else
        'todo: need pick list to identify non-labor into these subcategories: MATERIAL; OTHER_DIRECT_COSTS; SUBCONTRACT
        strEOC = "MATERIAL"
    End Select
    strResourcesJSON = strResourcesJSON & Chr(34) & "ElementOfCostId" & Chr(34) & ": " & Chr(34) & strEOC & Chr(34)
    strResourcesJSON = strResourcesJSON & "},"
    
    'ResourceAssignments.json
    For Each oAssignment In oResource.Assignments
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & "{"
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "ResourceID" & Chr(34) & ": " & Chr(34) & oResource.Guid & Chr(34) & ","
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oAssignment.TaskGuid & Chr(34) & ","
      If oResource.Type = pjResourceTypeWork Then
        'Budget_AtCompletion_Dollars
        'Budget_AtCompletion_Hours
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Budget_AtCompletion_Hours" & Chr(34) & ": " & Chr(34) & oAssignment.BaselineWork / 60 & Chr(34) & ","
        'Budget_AtCompletion_Hours
        'Estimate_ToComplete_Dollars
        'Estimate_ToComplete_Hours
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Estimate_ToComplete_Hours" & Chr(34) & ": " & Chr(34) & oAssignment.RemainingWork / 60 & Chr(34) & ","
        'Estimate_ToComplete_Hours
        'todo: Actual_ToDate_Dollars
        'todo: Actual_ToDate_Hours
      Else
        'Budget_AtCompletion_Dollars
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Budget_AtCompletion_Dollars" & Chr(34) & ": " & Chr(34) & oAssignment.BaselineCost & Chr(34) & ","
        'Budget_AtCompletion_Hours
        'Estimate_ToComplete_Dollars
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Estimate_ToComplete_Dollars" & Chr(34) & ": " & Chr(34) & oAssignment.RemainingCost & Chr(34) & ","
        'Estimate_ToComplete_Hours
        'todo: Actual_ToDate_Dollars
        'todo: Actual_ToDate_Hours
      End If
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "PhysicalPercentComplete" & Chr(34) & ": " & Chr(34) & oAssignment.Task.PhysicalPercentComplete & Chr(34)
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & "},"
    Next 'oAssignment
    
  Next 'oResource
    
  'create Resources.json
  If Len(strResourcesJSON) > 0 Then
    strResourcesJSON = "[" & Left(strResourcesJSON, Len(strResourcesJSON) - 1) & "]"
    Print #lngResourcesFile, strResourcesJSON
  End If
  
  'create ResourceAssignments.json
  If Len(strResourceAssignmentsJSON) > 0 Then
    strResourceAssignmentsJSON = "[" & Left(strResourceAssignmentsJSON, Len(strResourceAssignmentsJSON) - 1) & "]"
    Print #lngResourceAssignmentsFile, strResourceAssignmentsJSON
  End If
  
  cptJSON_Resources = True
  
exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Set oResource = Nothing
  Close #lngResourcesFile
  Close #lngResourceAssignmentsFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_Resources", Err, Erl)
  cptJSON_Resources = False
  Resume exit_here
End Function

Sub cptShowFrmIPMDAR()
'objects
Dim oFile As Scripting.File
Dim aContracts As Object
Dim oRootDir As Object
Dim oSubDir As Object
Dim oFSO As Scripting.FileSystemObject 'Object
Dim aSubmittals As Object
Dim oSubProject As Object
Dim oResource As Resource
Dim oTask As Task
'strings
Dim strBuffer As String
Dim strPeriod As String
Dim strContract As String
Dim strFile As String
Dim strCalendarComments As String
Dim strDir As String
Dim strExisting As String
'longs
Dim lngFile As Long
Dim lngContracts As Long
Dim lngCalendar As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
Dim vbResponse As Variant
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Please provide a Project Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'confirm IPMDAR directory exists
  strDir = Environ("USERPROFILE") & "\IPMDAR\"
  If Dir(strDir, vbDirectory) = vbNullString Then
    MkDir strDir
  End If
  
  'load contracts in directory
  Set aContracts = CreateObject("System.Collections.SortedList")
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oRootDir = oFSO.GetFolder(strDir)
  For Each oSubDir In oRootDir.SubFolders
    aContracts.Add Replace(oSubDir.Path, strDir, ""), Replace(oSubDir.Path, strDir, "")
    'todo: automatically link to the directory with matching guid.txt
    Set oFile = oFSO.GetFile(oSubDir.Path & "\guid.txt")
    'todo: if directory name doesn't match contract name then sync them up
    'todo: prompt 'json assets associated with this project are in directory xxx. -- or simply link them to the directory, and next COBRA load it will be aligned
  Next
  'list the contracts
  cptIPMDAR_frm.cboContract.Clear
  For lngContracts = aContracts.Count To 1 Step -1
    cptIPMDAR_frm.cboContract.AddItem aContracts.getByIndex(lngContracts - 1)
  Next lngContracts
  
  'todo: add 'new contract' button? notify location of Contract Name? Reset?
  'todo: cboContract_Change >> update docprops >> refresh cboPeriods >> [lots of fields updated]
  'todo: save ActiveProject.GetServerProjectGuid to txt and validate
  'todo: - if not found, find it in a subdirectory // if contract names match then...what?
  
  'confirm contract
  strDir = ""
  On Error Resume Next
  strDir = ActiveProject.CustomDocumentProperties("cptSPD_DIR").Value 'trips error if it doesn't exist
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  'if stored contract no longer matches one of the directories then prompt to reset
  If Dir(strDir, vbDirectory) = vbNullString Then strDir = ""
  If Len(strDir) = 0 Then
    vbResponse = InputBox("(We recommend matching the contract name in COBRA)." & vbCrLf & "Contract Name:", "Please provide a Contract Name:")
    If StrPtr(vbResponse) = 0 Then 'user hit cancel
      GoTo exit_here
    ElseIf Len(vbResponse) = 0 Then 'user entered zero-length string
      GoTo exit_here
    Else
      strContract = CStr(vbResponse)
    End If
    strDir = Environ("USERPROFILE") & "\IPMDAR\" & strContract
    If Dir(strDir, vbDirectory) = vbNullString Then
      MkDir strDir
    End If
    On Error Resume Next
    ActiveProject.CustomDocumentProperties.Add Name:="cptSPD_DIR", LinkToContent:=False, Type:=4, Value:=strDir
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    ActiveProject.CustomDocumentProperties("cptSPD_DIR").Value = strDir
    lngFile = FreeFile
    Open strDir & "\guid.txt" For Output As #lngFile
    Print #lngFile, ActiveProject.GetServerProjectGuid
    Close #lngFile
    'make this file read-only and archive
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFSO.GetFile(strDir & "\guid.txt")
    oFile.Attributes = ReadOnly
    oFile.Attributes = Hidden
  Else
    'todo: validate guid matches
    If Dir(strDir & "\guid.txt", vbHidden + vbReadOnly) <> vbNullString Then
      lngFile = FreeFile
      Open strDir & "\guid.txt" For Input As #lngFile
      Do While Not EOF(lngFile)
        Line Input #lngFile, strBuffer
      Loop
      Close #lngFile
      If strBuffer <> ActiveProject.GetServerProjectGuid Then
        If MsgBox("The associated directory is linked to a different source file." & vbCrLf & vbCrLf & "Proceed anyway?", vbExclamation + vbYesNo, "GUID Mismatch") = vbNo Then
          GoTo exit_here
        Else
          'todo: handle mismatched guid; store source project file name as second value; use Line Input #lngFile,strGUID,strProjectName
        End If
      End If
    End If
  End If
  On Error Resume Next
  cptIPMDAR_frm.cboContract.Value = Mid(strDir, InStrRev(strDir, "\") + 1)
  If Err.Number = 380 Then 'item not yet in list, so add it
    cptIPMDAR_frm.cboContract.AddItem strContract
    cptIPMDAR_frm.cboContract.Value = Mid(strDir, InStrRev(strDir, "\") + 1)
    Err.Clear
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  'generate a directory with current status date
  'todo: should this really be automatic on form load?
  strPeriod = strDir & "\" & Format(dtStatus, "yyyy-mm-dd")
  If Dir(strPeriod, vbDirectory) = vbNullString Then
    MkDir strPeriod
  End If
  
  'load listbox
  With cptIPMDAR_frm.lboFiles
    .Clear
    .AddItem "FileType.txt"
    .AddItem "DatasetMetadata.json"
    .AddItem "SourceSoftwareMetadata.json"
    .AddItem "ProjectScheduleData.json"
    .AddItem "ProjectCustomFieldDefinitions.json"
    .AddItem "ProjectCustomFieldValues.json"
    .AddItem "Calendars.json"
    .AddItem "CalendarWorkshifts.json"
    .AddItem "CalendarExceptions.json"
    .AddItem "Tasks.json"
    .AddItem "TaskScheduleData.json"
    .AddItem "TaskCustomFieldDefinitions.json"
    .AddItem "TaskCustomFieldValues.json"
    .AddItem "TaskConstraints.json"
    .AddItem "TaskRelationships.json"
    .AddItem "TaskOutlineStructure.json"
    .AddItem "Resources.json"
    .AddItem "ResourceCustomFieldDefinitions.json"
    .AddItem "ResourceCustomFieldValues.json"
    .AddItem "ResourceAssignments.json"
  End With
    
  'File.txt
  cptIPMDAR_frm.txtSchema = "IPMDAR_SCHEDULE_PERFORMANCE_DATASET/1.0"
  cptIPMDAR_frm.lboFiles.ListIndex = 0
  
  'DataSetMetadata
  With cptIPMDAR_frm
    .txtReportingPeriodEndDate = Format(dtStatus, "yyyy-mm-dd")
    With .cboContractorIDCodeTypeID
      .Clear
      .AddItem "DUNS"
      .AddItem "DUNS_PLUS_4"
      .AddItem "CAGE"
    End With
    .optEVFalse = True
  End With
  
  'SourceSoftwareMetadata
  With cptIPMDAR_frm
    .txtData_SoftwareName = Application.Name
    .txtData_SoftwareVersion = Application.Version
    .txtData_SoftwareCompanyName = "Microsoft Corporation"
    .txtData_SoftwareComments = Application.Name & " " & Application.Build & " on " & Application.OperatingSystem
    .txtExport_SoftwareName = "ClearPlan Toolbar"
    .txtExport_SoftwareVersion = Replace(Replace(cptRegEx(ThisProject.VBProject.VBComponents("cptIPMDAR_bas").CodeModule.Lines(1, 3), "<cpt_version>.*<\/cpt_version>"), "<cpt_version>", ""), "</cpt_version>", "")
    .txtExport_SoftwareCompanyName = "ClearPlan Consulting, LLC"
    .txtExport_SoftwareComments = "www.ClearPlanConsulting.com"
  End With
  
  'ProjectScheduleData
  With cptIPMDAR_frm
    .txtStatusDate = Format(dtStatus, "yyyy-mm-dd")
    .txtCurrentStartDate = Format(ActiveProject.ProjectStart, "yyyy-mm-dd")
    .txtCurrentFinishDate = Format(ActiveProject.ProjectFinish, "yyyy-mm-dd")
    If IsDate(ActiveProject.ProjectSummaryTask.BaselineStart) Then
      .txtBaselineStartDate = Format(ActiveProject.ProjectSummaryTask.BaselineStart, "yyyy-mm-dd")
    End If
    If IsDate(ActiveProject.ProjectSummaryTask.BaselineFinish) Then
      .txtBaselineFinishDate = Format(ActiveProject.ProjectSummaryTask.BaselineFinish, "yyyy-mm-dd")
    End If
    If IsDate(ActiveProject.ProjectSummaryTask.ActualStart) Then
      .txtActualStartDate = Format(ActiveProject.ProjectSummaryTask.ActualStart, "yyyy-mm-dd")
    End If
    If IsDate(ActiveProject.ProjectSummaryTask.ActualFinish) Then
      .txtActualFinishDate = Format(ActiveProject.ProjectSummaryTask.ActualFinish, "yyyy-mm-dd")
    End If
    .cboDurationUnitsID.Clear
    .cboDurationUnitsID.AddItem "DAYS"
    .cboDurationUnitsID.AddItem "HOURS"
    .cboDurationUnitsID.Value = "DAYS"
  End With
  
  'Calendars - loads on ListBox select
  
  'Tasks
  With cptIPMDAR_frm
    .cboTaskID.Clear
    .cboTaskID.AddItem "Unique ID"
    .cboTaskID.AddItem "GUID"
    .cboTaskID.AddItem "<< others >>"
    .txtName = "[Task]Name"
    With .cboSourceSubprojectReference
      .Clear
      .AddItem "[Task]Project"
      .Value = "[Task]Project"
    End With
    With .cboSourceTaskReference
      .Clear
      .AddItem "[Task]UniqueID"
      .Value = "[Task]UniqueID"
    End With
    With .cboComments
      .Clear
      .AddItem "[Task]Notes"
      For lngItem = 1 To 30
        .AddItem "Text" & lngItem
      Next lngItem
    End With
  End With
  
  'TaskScheduleData
  With cptIPMDAR_frm
    .cboPhysicalPercentComplete.Clear
    .cboPhysicalPercentComplete.AddItem "Physical % Complete" 'todo: lngField | name
    'todo: add only numbers...what about enterprise?
  End With
  
  'TaskOutlineStructure
  With cptIPMDAR_frm
    .optSummaryTasks.Value = True
    'todo: load project-specific saved settings
    .cboOutlineCode.Clear
    For lngItem = 1 To 10
      .cboOutlineCode.AddItem "Outline Code" & lngItem 'todo: make 2 cols, lngField, strName
    Next lngItem
    'todo: add ECFs?
  End With
  
  'Resources
  With cptIPMDAR_frm
    .cboResourceID.Clear
    .cboResourceID.AddItem "Unique ID"
    .cboResourceID.AddItem "GUID"
    .cboResourceID.AddItem "<< others >>" 'todo: other fields?
    'todo: load project-specific saved settings
  End With
  
  'ResourceAssignments
  With cptIPMDAR_frm
    'A_ResourceID = cboResourceID
    'A_TaskID = cboTaskID
    .txtA_Budget_AtCompletion_Dollars.Value = "[Resource]BaselineCost"
    .txtA_Budget_AtCompletion_Hours.Value = "[Resource]BaselineWork"
    .txtA_Estimate_ToComplete_Dollars.Value = "[Resource]RemainingCost"
    .txtA_Estimate_ToComplete_Hours.Value = "[Resource]RemainingWork"
    .txtA_Actual_ToDate_Dollars.Value = "null"
    .txtA_Actual_ToDate_Hours.Value = "null"
    'A_PhysicalPercentComplete = Task Physical % Complete
  End With
  
  'todo: load existing settings
  
  'set to first page and show it
  cptIPMDAR_frm.mpOptions.Value = 0
  cptIPMDAR_frm.Show False
  cptIPMDAR_frm.lboFiles.SetFocus
  
  'todo: for enumerated values, show list of required values
  'todo: for enumerated values, provide auto-generation
  'todo: provide peek for selected project/task? e.g.: source | sample | json
  'todo: add null as a value in all nullable fields?
  'todo: tab order
  
exit_here:
  On Error Resume Next
  Set oFile = Nothing
  Set aContracts = Nothing
  Set oRootDir = Nothing
  Set oSubDir = Nothing
  Set oFSO = Nothing
  Set aSubmittals = Nothing
  Set oSubProject = Nothing
  Set oResource = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_bas", "cptShowFrmCptIPMDAR", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowFrmCptTaskTypeMap()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  With cptTaskTypeMapping_frm
    With .cboFieldToMap
      .Clear
      .AddItem "* TaskTypeID"
      .AddItem "TaskSubTypeID"
    End With
    With .cboWhereField
      .Clear
      .AddItem
      .List(.ListCount - 1, 0) = FieldNameToFieldConstant("Name", pjTask)
      .List(.ListCount - 1, 1) = "Task Name"
    End With
    With .cboOperator
      .Clear
      .AddItem "equals"
      .AddItem "contains"
      .AddItem "begins with"
      .AddItem "does not equal"
      .AddItem "does not contain"
    End With
    With .cboTaskType
      .Clear
      .AddItem "ACTIVITY"
      .AddItem "MILESTONE"
      .AddItem "SUMMARY"
      .AddItem "HAMMOCK"
    End With
    
    .Show False
    
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_Bas", "cptShowFrmCptTaskTypeMap", Err, Erl)
  Resume exit_here
End Sub

Function CHARW(CharCode As Variant, Optional Exact_functionality As Boolean = False) As String
'Use a Leading "U" or "u" to indicate Unicode values
'Exact_functionality returns the Unicode characters for Ascii(128) to Ascii(159) rather than
'the Windows characters

   If UCase(Left$(CharCode, 1)) = "U" Then CharCode = Replace(CharCode, "U", "&H", 1, 1, vbTextCompare)
   CharCode = CLng(CharCode)
   If CharCode < 256 Then
      If Exact_functionality Then
         CHARW = ChrW(CharCode)
      Else
         CHARW = Chr(CharCode)
      End If
   Else
      CHARW = ChrW(CharCode)
   End If
End Function

Sub cptCreateIPMDARWorkbook(strDir As String)
'objects
Dim oFSO As FileSystemObject
Dim oFolder As Folder
Dim oFile As File
Dim xlApp As Excel.Application
Dim oWorkbook As Workbook
Dim oWorksheet As Worksheet
'strings
Dim strFormula As String
Dim strSource As String
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'set up the workbook
  Set xlApp = CreateObject("Excel.Application")
  Set oWorkbook = xlApp.Workbooks.Add
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.GetFolder(strDir)
  For Each oFile In oFolder.Files
    If LCase(Right(oFile.Path, 5)) = ".json" Then
      
      'get source
      strSource = Replace(oFile.Name, ".json", "")
      
      'get formula for source
      'start with common beginning
      
      strFormula = _
      "let" & Chr(13) & "" & Chr(10) & _
      "    Source = Json.Document(File.Contents(""C:\Users\arong\IPMDAR\2016-08-26\" & strSource & ".json""))," & Chr(13) & "" & Chr(10) & _
      "    #""Converted to Table"" = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error)," & Chr(13) & "" & Chr(10)

      Select Case strSource 'todo: remaining .json files
        Case "CalendarExceptions"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""CalendarID"", ""ExceptionDate"", ""WorkHours""}, {""CalendarID"", ""ExceptionDate"", ""WorkHours""})          "
        
        Case "Calendars"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""ID"", ""Name"", ""Comments""}, {""ID"", ""Name"", ""Comments""})"
        
        Case "CalendarWorkshifts"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""CalendarID"", ""Ordinal"", ""SundayWorkHours"", ""MondayWorkHours"", ""TuesdayWorkHours"", ""WednesdayWorkHours"", ""ThursdayWorkHours"", ""FridayWorkHours"", ""SaturdayWorkHours""}, {""CalendarID"", ""Ordinal"", ""SundayWorkHours"", ""MondayWorkHours"", ""TuesdayWorkHours"", ""WednesdayWorkHours"", ""ThursdayWorkHours"", ""FridayWorkHours"", ""SaturdayWorkHours""})"
        
        Case "ProjectScheduleData"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""StatusDate"", ""CurrentStartDate"", ""CurrentFinishDate"", ""BaselineStartDate"", ""BaselineFinishDate"", ""ActualStartDate"", ""ActualFinishDate"", ""DurationUnitsID""}, {""StatusDate"", ""CurrentStartDate"", ""CurrentFinishDate"", ""BaselineStartDate"", ""BaselineFinishDate"", ""ActualStartDate"", ""ActualFinishDate"", ""DurationUnitsID""})"
          
        Case "ResourceAssignments"
          strFormula = strFormula & _
          "   #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""ResourceID"", ""TaskID"", ""Budget_AtCompletion_Hours"", ""Estimate_ToComplete_Hours"", ""PhysicalPercentComplete""}, {""ResourceID"", ""TaskID"", ""Budget_AtCompletion_Hours"", ""Estimate_ToComplete_Hours"", ""PhysicalPercentComplete""})"
        
        Case "Resources"
          strFormula = strFormula & _
          "   #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""ID"", ""Name"", ""ElementOfCostId""}, {""ID"", ""Name"", ""ElementOfCostId""})"
        
        Case "SourceSoftwareMetadata"
          strFormula = strFormula & _
          "   #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""Data_SoftwareName"", ""Data_SoftwareVersion"", ""Data_SoftwareCompanyName"", ""Data_SoftwareComments"", ""Export_SoftwareName"", ""Export_SoftwareVersion"", ""Export_SoftwareCompanyName"", ""Export_SoftwareComments""}, {""Data_SoftwareName"", ""Data_SoftwareVersion"", ""Data_SoftwareCompanyName"", ""Data_SoftwareComments"", ""Export_SoftwareName"", ""Export_SoftwareVersion"", ""Export_SoftwareCompanyName"", ""Export_SoftwareComments""})"

        Case "TaskConstraints"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""TaskID"", ""ConstraintTypeID"", ""OtherConstraintType"", ""ConstraintDate""}, {""TaskID"", ""ConstraintTypeID"", ""OtherConstraintType"", ""ConstraintDate""})"
          
        Case "TaskOutlineStructure"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""Level"", ""TaskID"", ""ParentTaskID""}, {""Level"", ""TaskID"", ""ParentTaskID""})"
          
        Case "TaskRelationships"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""PredecessorTaskID"", ""SuccessorTaskID"", ""RelationshipTypeID"", ""LagDuration"", ""LagCalendarID""}, {""PredecessorTaskID"", ""SuccessorTaskID"", ""RelationshipTypeID"", ""LagDuration"", ""LagCalendarID""})          "
        
        Case "Tasks"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""ID"", ""Name"", ""TaskTypeID"", ""SourceSubprojectReference"", ""SourceTaskReference""}, {""ID"", ""Name"", ""TaskTypeID"", ""SourceSubprojectReference"", ""SourceTaskReference""})"
          
        Case "TaskScheduleData"
          strFormula = strFormula & _
          "    #""Expanded Column1"" = Table.ExpandRecordColumn(#""Converted to Table"", ""Column1"", {""ID"", ""CalendarID"", ""CurrentDuration"", ""CurrentStartDate"", ""CurrentFinishDate"", ""EarlyStartDate"", ""EarlyFinishDate"", ""LateStartDate"", ""LateFinishDate"", ""FreeFloatDuration"", ""TotalFloatDuration"", ""OnCriticalPath"", ""CalculatedPercentComplete"", ""RemainingDuration""}, {""ID"", ""CalendarID"", ""CurrentDuration"", ""CurrentStartDate"", ""CurrentFinishDate"", ""EarlyStartDate"", ""EarlyFinishDate"", ""LateStartDate"", ""LateFinishDate"", ""FreeFloatDuration"", ""TotalFloatDuration"", ""OnCriticalPath"", ""CalculatedPercentComplete"", ""RemainingDuration""})"
          
      End Select
    
      'formula tail
      strFormula = strFormula & Chr(13) & "" & Chr(10) & _
          "in" & Chr(13) & "" & Chr(10) & _
          "    #""Expanded Column1"""

      'add the queries
      oWorkbook.Queries.Add strSource, strFormula
      Set oWorksheet = oWorkbook.Sheets.Add
        
      With oWorksheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1; Data Source=$Workbook$;Location=" & strSource & ";Extended Properties=""""", Destination:=oWorksheet.Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & strSource & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = strSource
        .Refresh BackgroundQuery:=False
      End With
      oWorksheet.Name = strSource
    End If
  Next 'oFile In oFolder.Files

  'todo: add conditional formatting for duplicate task names
  'todo: ensure all tasks with SUMMARY are included in the OutlineStrucure
  'todo: ensure all outlinelevel=1 have null parent
  'todo: ensure uniqueness based on Primary Keys

  xlApp.Visible = True
  If Dir(strDir & "IPMDAR_DATA_REVIEW.xlsx") <> vbNullString Then
    'todo: provide way to select the reporting period in the workbook - maps to directory structure?
    If MsgBox("IPMDAR Data Review Workbook already exists in this location." & vbCrLf & vbCrLf & "Overwrite?", vbQuestion + vbYesNo, "Confirm Overwrite") = vbYes Then
      oWorkbook.SaveAs strDir & "IPMDAR_DATA_REVIEW.xlsx", 51
    End If
  End If

exit_here:
  On Error Resume Next
  Set oFSO = Nothing
  Set oFolder = Nothing
  Set oFile = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set xlApp = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR", "cptCreateIPMDARWorkbook", Err, Erl)
  Resume exit_here
End Sub

Sub cptRequestCOBRAData()
'objects
Dim olApp As Outlook.Application
Dim oMailItem As MailItem
Dim oDoc As Word.Document
Dim oSel As Word.Selection
'strings
Dim strContract As String
Dim strFile As String
Dim strSQL As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ensure status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Please provide a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If

  'generate *.sql file
  'do not include contract name by design - might change - prompt COBRA Analyst for proper name
  lngFile = FreeFile
  strFile = Environ("TMP") & "\spd-cobra-query.sql"
  Open strFile For Output As #lngFile
  strSQL = "DECLARE @MyProj VARCHAR(MAX) " & vbCrLf
  strSQL = strSQL & "SET @MyProj=inputbox('Project Name:') " & vbCrLf
  strSQL = strSQL & "SELECT CLASSFCN SecurityMarking," & vbCrLf
  strSQL = strSQL & "    CASE" & vbCrLf
  strSQL = strSQL & "        WHEN CONT_STATEMENT IS NULL THEN ''" & vbCrLf
  strSQL = strSQL & "        ELSE CONT_STATEMENT " & vbCrLf
  strSQL = strSQL & "    END DistributionStatement," & vbCrLf
  strSQL = strSQL & "    CONVERT(varchar(10),STATUSDATE,126) ReportingPeriodEnDate," & vbCrLf
  strSQL = strSQL & "    CONT_NAME ContractorName," & vbCrLf
  strSQL = strSQL & "    CONT_IDTYPE ContractorIDCodeTypeID," & vbCrLf
  strSQL = strSQL & "    CONT_IDCODE ContractorIDCode," & vbCrLf
  strSQL = strSQL & "    ADDRESS ContractorAddress_Street," & vbCrLf
  strSQL = strSQL & "    CITY ContractorAddress_City," & vbCrLf
  strSQL = strSQL & "    STATE ContractorAddress_State," & vbCrLf
  strSQL = strSQL & "    COUNTRY ContractorAddress_Country," & vbCrLf
  strSQL = strSQL & "    ZIP ContractorAddress_ZipCode," & vbCrLf
  strSQL = strSQL & "    CONT_REPN PointOfContactName," & vbCrLf
  strSQL = strSQL & "    CONT_REPT PointOfContactTitle," & vbCrLf
  strSQL = strSQL & "    CONT_REPPHONE PointOfContactTelephone," & vbCrLf
  strSQL = strSQL & "    CONT_REPEMAIL PointOfContactEmail," & vbCrLf
  strSQL = strSQL & "    CONTRACT ContractName," & vbCrLf
  strSQL = strSQL & "    CONT_NO ContractNumber," & vbCrLf
  strSQL = strSQL & "    CONT_TYPE ContractType," & vbCrLf
  strSQL = strSQL & "    CONT_TASK ContractTaskOrEffortName," & vbCrLf
  strSQL = strSQL & "    CONT_PROGRAM ProgramName," & vbCrLf
  strSQL = strSQL & "    CONT_PHASE ProgramPhase," & vbCrLf
  strSQL = strSQL & "    CASE " & vbCrLf
  strSQL = strSQL & "        WHEN EVMS_ACC=1 THEN 'TRUE' " & vbCrLf
  strSQL = strSQL & "        Else 'FALSE' " & vbCrLf
  strSQL = strSQL & "    END EVMSAccepted," & vbCrLf
  strSQL = strSQL & "    CONVERT(varchar(10),EVMS_ADATE,126) EVMSAcceptanceDate " & vbCrLf
  strSQL = strSQL & "FROM PROGRAM " & vbCrLf
  strSQL = strSQL & "WHERE [PROGRAM]=@MyProj"
  Print #lngFile, strSQL
  Close #lngFile
  
  'get contract name
  strContract = cptIPMDAR_frm.cboContract.Value
  
  'create email, attach file
  On Error Resume Next
  Set olApp = GetObject(, "Outlook.Application")
  If olApp Is Nothing Then
    Set olApp = CreateObject("Outlook.Application")
  End If
  Set oMailItem = olApp.CreateItem(olMailItem)
  With oMailItem
    .Subject = "REQUEST: COBRA data for " & strContract & " IPMDAR " & Format(dtStatus, "yyyy-mm-dd")
    Set oDoc = .GetInspector.WordEditor
    Set oSel = oDoc.Windows(1).Selection
    oSel.Move wdStory, -1
    oSel.TypeText "Please run the attached query in COBRA's SQL Tool for the '" & strContract & "' contract and provide the resulting csv file at your earliest convenience."
    .Attachments.Add strFile
    .Importance = olImportanceHigh
    .Display False
  End With
    
  'delete tmp file
  If Dir(strFile) <> vbNullString Then Kill strFile
  
exit_here:
  On Error Resume Next
  Set olApp = Nothing
  Set oMailItem = Nothing
  Set oDoc = Nothing
  Set oSel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_bas", "cptRequestCOBRAData", Err, Erl)
  Resume exit_here

End Sub

Sub cptLoadCOBRAData()
'objects
Dim FileDialog As FileDialog
Dim xlApp As Excel.Application
'strings
Dim strGUID As String
Dim strMsg As String
Dim strNewDir As String
Dim strDir As String
Dim strJSON As String
Dim strData As String
Dim strFile As String
'longs
Dim lngFile As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
Dim vItem As Variant
Dim vData As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  Set FileDialog = xlApp.FileDialog(msoFileDialogFilePicker)
try_again:
  strFile = ""
  With FileDialog
    .AllowMultiSelect = False
    .ButtonName = "Load"
    .InitialView = msoFileDialogViewDetails
    .InitialFileName = Environ("USERPROFILE") & "\IPMDAR\" & cptIPMDAR_frm.cboContract.Value & "\" & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "\"
    .Title = "Select COBRA Data File:"
    .Filters.Add "Comma Separated Values (csv)", "*.csv"
    
    If .Show = -1 Then
      If .SelectedItems.Count > 0 Then
        For lngItem = 1 To .SelectedItems.Count
          strFile = .SelectedItems(lngItem)
        Next lngItem
      Else
        GoTo exit_here
      End If
    End If
  End With
  
  'skip if no file selected
  If Len(strFile) = 0 Then GoTo exit_here
  
  'open file and read it into an array
  lngFile = FreeFile
  Open strFile For Input As #lngFile
  Line Input #lngFile, strData
  Close #lngFile
  
  'parse the array
  vData = Split(strData, ",")
  
  'validate period end date against current status date
  If vData(2) <> Format(ActiveProject.StatusDate, "yyyy-mm-dd") Then
    strMsg = "WARNING:" & vbCrLf & vbCrLf
    strMsg = strMsg & "The selected file contains data for period ending " & vData(2) & " which does not match the current status date of " & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "." & vbCrLf & vbCrLf
    strMsg = strMsg & "Are you sure this is the correct file?"
    Select Case MsgBox(strMsg, vbExclamation + vbYesNoCancel, "Misaligned Data Date")
      Case vbCancel
        GoTo exit_here
      Case vbNo
        GoTo try_again
      Case vbYes
        'do nothing
    End Select
  End If
  
  'load the data
  With cptIPMDAR_frm
    .txtSecurityMarking = vData(0)
    strJSON = strJSON & Chr(34) & Replace(.txtSecurityMarking.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(0) & Chr(34) & ","
    
    .txtDistributionStatement = vData(1)
    strJSON = strJSON & Chr(34) & Replace(.txtDistributionStatement.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(1) & Chr(34) & ","
    
    .txtReportingPeriodEndDate = vData(2)
    strJSON = strJSON & Chr(34) & Replace(.txtReportingPeriodEndDate.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(2) & Chr(34) & ","
    
    .txtContractorName = vData(3)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorName.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(3) & Chr(34) & ","
    
    .cboContractorIDCodeTypeID = vData(4) 'COBRA enum matches SPD FFS enum
    strJSON = strJSON & Chr(34) & Replace(.cboContractorIDCodeTypeID.Name, "cbo", "") & Chr(34) & ": " & Chr(34) & vData(4) & Chr(34) & ","
    
    .txtContractorIDCode = vData(5)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorIDCode.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(5) & Chr(34) & ","
    
    .txtContractorAddress_Street = vData(6)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorAddress_Street.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(6) & Chr(34) & ","
    
    .txtContractorAddress_City = vData(7)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorAddress_City.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(7) & Chr(34) & ","
    
    .txtContractorAddress_State = vData(8)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorAddress_State.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(8) & Chr(34) & ","
    
    .txtContractorAddress_Country = vData(9)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorAddress_Country.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(9) & Chr(34) & ","
    
    .txtContractorAddress_ZipCode = vData(10)
    strJSON = strJSON & Chr(34) & Replace(.txtContractorAddress_ZipCode.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(10) & Chr(34) & ","
    
    .txtPointOfContactName = vData(11)
    strJSON = strJSON & Chr(34) & Replace(.txtPointOfContactName.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(11) & Chr(34) & ","
    
    .txtPointOfContactTitle = vData(12)
    strJSON = strJSON & Chr(34) & Replace(.txtPointOfContactTitle.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(12) & Chr(34) & ","
    
    .txtPointOfContactTelephone = vData(13)
    strJSON = strJSON & Chr(34) & Replace(.txtPointOfContactTelephone.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(13) & Chr(34) & ","
    
    .txtPointOfContactEmail = vData(14)
    strJSON = strJSON & Chr(34) & Replace(.txtPointOfContactEmail.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(14) & Chr(34) & ","
    
    .txtContractName = vData(15)
    strJSON = strJSON & Chr(34) & Replace(.txtContractName.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(15) & Chr(34) & ","
    'prompt to align contract name
    If vData(15) <> Dir(ActiveProject.CustomDocumentProperties("cptSPD_DIR").Value, vbDirectory) Then
      If MsgBox("Align directory with COBRA's contract name '" & vData(15) & "'?", vbExclamation + vbYesNo, "Mismatched Contract Name") = vbYes Then
        'rename the directory if it doesn't exist
        strDir = Environ("USERPROFILE") & "\IPMDAR\" & cptIPMDAR_frm.cboContract.Value
        strNewDir = Environ("USERPROFILE") & "\IPMDAR\" & vData(15)
        If Dir(strNewDir, vbDirectory) <> vbNullString Then
          If Dir(strNewDir & "\guid.txt", vbHidden + vbReadOnly) <> vbNullString Then
            lngFile = FreeFile
            Open strNewDir & "\guid.txt" For Input As #lngFile
            Do While Not EOF(lngFile)
              Line Input #lngFile, strGUID 'todo: add source project name
            Loop
            Close #lngFile
            If strGUID <> ActiveProject.GetServerProjectGuid Then
              MsgBox "This directory is linked to a different source file."
              'todo: how to handle this situation
              'todo: also follow this procedure on form show
            End If
          Else
            MsgBox "The directory already exists. Please manually move subdirectories to the existing directory and delete old contract directory.", vbInformation + vbOKOnly, "Contract Directory Exists"
            'todo: handle guid.txt here
          End If
          'todo: prompt to automatically move all files/subdirs to new dir?
        Else
          Name strDir As strNewDir
        End If
        'update doc prop
        ActiveProject.CustomDocumentProperties("cptSPD_DIR").Value = Dir(strNewDir, vbDirectory)
        'update cboContract
        'todo: what if contract/directory is not yet in the list
        cptIPMDAR_frm.cboContract.Value = vData(15)
      End If
    End If
    
    .txtContractNumber = vData(16)
    strJSON = strJSON & Chr(34) & Replace(.txtContractNumber.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(16) & Chr(34) & ","
    
    .txtContractType = vData(17)
    strJSON = strJSON & Chr(34) & Replace(.txtContractType.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(17) & Chr(34) & ","
    
    .txtContractTaskOrEffortName = vData(18)
    strJSON = strJSON & Chr(34) & Replace(.txtContractTaskOrEffortName.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(18) & Chr(34) & ","
    
    .txtProgramName = vData(19)
    strJSON = strJSON & Chr(34) & Replace(.txtProgramName.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(19) & Chr(34) & ","
    
    .txtProgramPhase = vData(20)
    strJSON = strJSON & Chr(34) & Replace(.txtProgramPhase.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(20) & Chr(34) & ","
    
    .optEVTrue = CBool(vData(21))
    strJSON = strJSON & Chr(34) & Replace(.optEVTrue.Name, "opt", "") & Chr(34) & ": " & Chr(34) & vData(21) & Chr(34)
    
    If UBound(vData) > 21 Then
      .txtEVMSAcceptanceDate = vData(22)
      strJSON = strJSON & "," & Chr(34) & Replace(.txtEVMSAcceptanceDate.Name, "txt", "") & Chr(34) & ": " & Chr(34) & vData(22) & Chr(34)
    End If
    
    'confirm root IPMDAR directory
    strDir = Environ("USERPROFILE") & "\IPMDAR"
    If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
    'confirm contract directory
    strDir = strDir & "\" & cptIPMDAR_frm.cboContract.Value
    If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
    'confirm period directory
    strDir = strDir & "\" & .txtStatusDate.Value
    If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
    
  End With
  
  'create DatasetMetadata.json
  'todo: capture file exists
  lngFile = FreeFile
  Open strDir & "\DatasetMetadata.json" For Output As #lngFile
  Print #lngFile, "[{" & strJSON & "}]"
  Close #lngFile
  
exit_here:
  On Error Resume Next
  Set FileDialog = Nothing
  Set xlApp = Nothing
  For lngFile = 1 To FreeFile
    Close #lngFile
  Next lngFile
  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_bas", "cptLoadCOBRAData", Err, Erl)
  Resume exit_here
End Sub

Private Function cptParseCalendarComments(strFile As String, strGUID As String) As String
'objects
'strings
Dim strDir As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  cptParseCalendarComments = ""
  
  'get previous comments
  lngFile = FreeFile
  Open strFile For Input As #lngFile
  Dim strTemp As String
  Dim strBuffer As String
  Do While Not EOF(lngFile)
    Line Input #lngFile, strBuffer
    strTemp = strTemp & strBuffer
  Loop
  Close #lngFile
  
  If InStr(strTemp, strGUID) > 0 Then
    'parse calendar object
    strTemp = cptRegEx(strTemp, strGUID & "[^}]*")
    'parse comment element
    strTemp = cptRegEx(strTemp, "Comments"": .*[^\s]")
    'parse comment
    strTemp = Replace(strTemp, "Comments"": ", "")
    strTemp = Replace(strTemp, Chr(34), "")
    'replace null
    strTemp = Replace(strTemp, "null", "")
  Else
    strTemp = ""
  End If
  
  cptParseCalendarComments = strTemp
    
exit_here:
  On Error Resume Next
  Close #lngFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR_bas", "cptParseCalendarComments", Err, Erl)
  Resume exit_here
End Function

Public Sub cptLoadCalendars()
'objects
Dim oTask As MSProject.Task
Dim oResource As MSProject.Resource
Dim oSubProject As MSProject.SubProject
Dim oProject As MSProject.Project
Dim aProjects As Object
'strings
Dim strExisting As String
Dim strCalendarComments As String
Dim strFile As String
Dim strDir As String
'longs
Dim lngProject As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'set dir and confirm existence
  strDir = Environ("USERPROFILE") & "\IPMDAR"
  If Dir(strDir, vbDirectory) = vbNullString Then
    MsgBox strDir & " not found!", vbCritical + vbOKOnly, "Critical Error"
    GoTo exit_here
  End If
  
  'create list of projects to avoid repetitive subproject code
  Set aProjects = CreateObject("System.Collections.ArrayList")
  aProjects.Add ActiveProject
  For Each oSubProject In ActiveProject.Subprojects
    aProjects.Add oSubProject.SourceProject
  Next oSubProject
  
  'set file path to previous period's Calendar.json
  If Not IsNull(cptIPMDAR_frm.cboPrevDir.Value) Then
    strFile = strDir & "\" & cptIPMDAR_frm.cboContract.Value & "\" & cptIPMDAR_frm.cboPrevDir.Value & "\Calendars.json"
  End If
  
  With cptIPMDAR_frm.lboCalendars
    
    .Clear
    
    strExisting = ""
    For lngProject = 0 To aProjects.Count - 1
      'load project calendar(s)
      Set oProject = aProjects(lngProject)
      If InStr(strExisting, oProject.Calendar.Guid) = 0 Then
        .AddItem
        .List(.ListCount - 1, 0) = oProject.Calendar.Guid
        .List(.ListCount - 1, 1) = oProject.Calendar.Name
        .List(.ListCount - 1, 2) = "x"
        If Dir(strFile) <> vbNullString Then
          strCalendarComments = cptParseCalendarComments(strFile, oProject.Calendar.Guid)
          If Len(strCalendarComments) > 0 Then
            .List(.ListCount - 1, 2) = "+"
            .List(.ListCount - 1, 3) = strCalendarComments
          End If
        End If
        strExisting = strExisting & "[" & oProject.Calendar.Guid & "]"
      End If
    Next lngProject
    
    For lngProject = 0 To aProjects.Count - 1
      'load resource calenders
      Set oProject = aProjects(lngProject)
      'todo: are resource calendars required?
      'todo: what if subproject resource calendars don't match?
      strExisting = ""
      For Each oResource In oProject.Resources
        If Not oResource.Calendar Is Nothing Then
          If oResource.Calendar.Exceptions.Count > oResource.Calendar.BaseCalendar.Exceptions.Count Then
            If InStr(strExisting, oResource.Calendar.Guid) = 0 Then
              .AddItem
              .List(.ListCount - 1, 0) = oResource.CalendarGuid
              .List(.ListCount - 1, 1) = oResource.Calendar
              .List(.ListCount - 1, 2) = "x"
              If Dir(strFile) <> vbNullString Then
                strCalendarComments = cptParseCalendarComments(strFile, oResource.CalendarGuid)
                If Len(strCalendarComments) > 0 Then
                  .List(.ListCount - 1, 2) = "+"
                  .List(.ListCount - 1, 3) = strCalendarComments
                End If
              End If
              strExisting = strExisting & "[" & oResource.Calendar.Guid & "]"
            End If
          End If
        End If
      Next oResource
      
    Next lngProject
    
    'load task calendars
    strExisting = ""
    For Each oTask In ActiveProject.Tasks
      If oTask.Calendar <> ActiveProject.Calendar And oTask.Calendar <> "None" Then
        If InStr(strExisting, oTask.Calendar) = 0 Then
          .AddItem
          .List(.ListCount - 1, 0) = oTask.CalendarGuid
          .List(.ListCount - 1, 1) = oTask.Calendar
          .List(.ListCount - 1, 2) = "x"
          If Not IsNull(cptIPMDAR_frm.cboPrevDir.Value) Then
            strFile = strDir & "\" & cptIPMDAR_frm.cboPrevDir.Value & "\Calendars.json"
            If Dir(strFile) <> vbNullString Then
              strCalendarComments = cptParseCalendarComments(strFile, oTask.CalendarGuid)
              If Len(strCalendarComments) > 0 Then
                .List(.ListCount - 1, 2) = "+"
                .List(.ListCount - 1, 3) = strCalendarComments
              End If
            End If 'Dir(strFile) <> vbNullString
          End If 'Not IsNull...
          strExisting = strExisting & "[" & oTask.Calendar & "]"
        End If 'Instr(strExisting...
      End If 'oTask.Calendar <> Activeproject.Calendar
    Next oTask
    'todo: what if subproject task calendars don't match?
        
    If .ListCount > 0 Then .ListIndex = 0
        
  End With 'cptIPMDAR_frm.lboCalendars

exit_here:
  On Error Resume Next
  Set oTask = Nothing
  Set oResource = Nothing
  Set oSubProject = Nothing
  Set oProject = Nothing
  Set aProjects = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptIPMDAR_bas", "cptLoadCalendars", Err, Erl)
  Resume exit_here
End Sub

Public Sub cptScrubIPMDAR(strDir As String)
'objects
Dim oFile As Scripting.File
Dim oFolder As Scripting.Folder 'Object
Dim oFSO As Scripting.FileSystemObject
'strings
Dim strTemp As String
Dim strBuffer As String
Dim strFile As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  'todo: requirements:
  'todo: - string values must not contain [\u0000-\u0008]|[\u000B-\u000C]|[\u000E-\u000F]|\u007F
  '-- reference: http://www.endmemo.com/unicode/ascii.php
  'todo: - string values used as IDs must be limited to [\u0020-\u007E]
  'todo: trim all text
  'todo: trim all instances of >1 spaces
  'todo: ensure all required fields are not null
  'todo: is JSON well-formed?

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'get directory
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oFolder = oFSO.GetFolder(strDir)
  'loop through each submittal file
  For Each oFile In oFolder.Files
    If InStr(oFile.Name, "json") Then
      Debug.Print "Scrubbing " & oFile.Name & "..." 'todo: put this in lblStatus
      lngFile = FreeFile
      Open oFile.Path For Input As #lngFile
      Do While Not EOF(lngFile)
        Line Input #lngFile, strBuffer
        strTemp = strTemp & strBuffer
      Loop
      If Len(cptRegEx(strBuffer, "[\u0000-\u0008]|[\u000B-\u000C]|[\u000E-\u000F]|\u007F")) > 0 Then
        MsgBox oFile.Name & " failed (but we fixed it for you).", vbInformation + vbOKOnly, "Boom"
        'todo: replace all occurences
        'todo: delete original
        'todo: rewrite the file
      End If
      strBuffer = ""
      Close #lngFile
    End If
  Next
  
  Debug.Print "...complete." 'todo: put this in lblStatus
  
exit_here:
  On Error Resume Next
  Set oFile = Nothing
  Set oFolder = Nothing
  Set oFSO = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptIPMDAR_bas", "cptScrubIPMDAR", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub
