Attribute VB_Name = "cptIPMDAR"
'<cpt_version>0.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptJSON_Main()
'objects
'strings
Dim strErr As String
Dim strDir As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'set up directories
  If Dir(Environ("USERPROFILE") & "\IPMDAR", vbDirectory) = vbNullString Then MkDir Environ("USERPROFILE") & "\IPMDAR\"
  strDir = Environ("USERPROFILE") & "\IPMDAR\" & Format(ActiveProject.StatusDate, "yyyy-mm-dd")
  If Dir(strDir, vbDirectory) = vbNullString Then MkDir Environ("USERPROFILE") & "\IPMDAR\" & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & "\"

  'create the FileType.txt
  lngFile = FreeFile
  Open strDir & "\FileType.txt" For Output As #lngFile
  Print #lngFile, "IPMDAR_SCHEDULE_PERFORMANCE_DATASET/1.0"
  Close #lngFile
  
  'create the exports

  'todo: DatasetMetadata.json
'  If Not cptJSON_DatasetMetadata(strDir) Then
'    strErr = "DatasetMetadata.json" & vbCrLf
'  End If

  'SourceSoftwareMetadata.json
  If Not cptJSON_SourceSoftwareMetadata(strDir) Then
    strErr = strErr & "SourceSoftwareMetadata.json" & vbCrLf
  End If
  
  'ProjectScheduleData.json
  If Not cptJSON_ProjectScheduleData(strDir) Then
    strErr = strErr & "ProjectScheduleData.json" & vbCrLf
  End If
  'todo: ProjectCustomFieldDefinitions.json
  'todo: ProjectCustomFieldValues.json
  
  'Calendars.json (includes CalendarWorkshifts.json and CalendarExceptions.json)
  If Not cptJSON_Calendars(strDir) Then
    strErr = strErr & "Calendars.json" & vbCrLf
  End If
  
  'todo: TaskCustomFieldDefinitions.json
  'Tasks.json (includes TaskScheduleData, TaskCustomFieldValues, TaskConstraints, TaskRelationships,TaskOutlineStructure)
  'todo: ensure task name uniqueness
  If Not cptJSON_Tasks(strDir) Then
    strErr = strErr & "Tasks.json" & vbCrLf
  End If
  'todo: TaskCustomFieldValues
  'todo: TaskRelationships
  
  'todo: ResourceCustomFieldDefinitions.json
  'Resources.json (includes ResourceCustomFieldValues, ResourceAssignments)
  If Not cptJSON_Resources(strDir) Then
    strErr = "Resources.json" & vbCrLf
  End If
  
  'todo: create schedule narrative template containing:
  'todo: create section headers [created with the ClearPlan toolbar
  'todo: -- create placeholders for all explanations for leads, lags, constraints
  'todo: -- create placeholders for CWBS, SOW if exist in Outline Code
  'todo: export IMS Data Dictionary
  'todo: prompt to save server file as .mpp and consolidate?
  'todo: create Validation Workbook w/json queries, highlighted duplicates
  
  'todo: scrub for character limitations SPD FFS 2.1.6
  
  'todo: create zip using DEFLATE method

  If Len(strErr) > 0 Then
    MsgBox "The following exports were not created successfully:" & vbCrLf & strErr, vbExclamation + vbOKOnly, "Incomplete"
  Else
    MsgBox "Schedule Performance Data exported correctly.", vbInformation + vbOKOnly, "SPD"
  End If


exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptIPMDAR", "cptJSAON_Main", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Function cptJSON_DatasetMetadata(strDir As String) As Boolean
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

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: requires user form
  'todo: or automatically setup project custom fields/documentproperties?
  'todo: load previous period's data
  
  strFile = strDir & "\DatasetMetadata.json"
  lngFile = FreeFile
  If Dir(strFile) <> vbNullString Then Kill strFile
  Open strFile For Output As #lngFile

  strJSON = "["
  'todo: get content
  strJSON = Left(strJSON, Len(strJSON) - 1) & "]"
  
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

Function cptJSON_SourceSoftwareMetadata(strDir) As Boolean
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

Function cptJSON_ProjectScheduleData(strDir) As Boolean
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
  strJSON = strJSON & Chr(34) & "StatusDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "CurrentStartDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.ProjectSummaryTask.Start, "yyyy-mm-dd") & Chr(34) & ","
  strJSON = strJSON & Chr(34) & "CurrentFinishDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.ProjectSummaryTask.Finish, "yyyy-mm-dd") & Chr(34) & ","
  If IsDate(ActiveProject.BaselineSavedDate(pjBaseline)) Then
    strJSON = strJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.ProjectSummaryTask.BaselineStart, "yyyy-mm-dd") & Chr(34) & ","
    strJSON = strJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.ProjectSummaryTask.BaselineFinish, "yyyy-mm-dd") & Chr(34) & ","
  Else
    strJSON = strJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": null,"
    strJSON = strJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": null,"
  End If
  If IsDate(ActiveProject.ProjectSummaryTask.ActualStart) Then
    strJSON = strJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.ProjectSummaryTask.ActualStart, "yyyy-mm-dd") & Chr(34) & ","
  Else
    strJSON = strJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": null,"
  End If
  If IsDate(ActiveProject.ProjectSummaryTask.ActualFinish) Then
    strJSON = strJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": " & Chr(34) & Format(ActiveProject.ProjectSummaryTask.ActualFinish, "yyyy-mm-dd") & Chr(34) & ","
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

Function cptJSON_ProjectCustomFieldDefinitions(strDir As String) As Boolean

End Function

Function cptJSON_Calendars(strDir As String) As Boolean
'objects
Dim aWorkHours As Object
Dim oCalendarException As MSProject.Exception
Dim oWorkshift As MSProject.WorkWeek
Dim oCalendar As MSProject.Calendar
Dim oSubProject As MSProject.SubProject
'strings
Dim strCalendarExceptionsJSON As String
Dim strCalendarExceptionsFile As String
Dim strCalendarWorkshiftsJSON As String
Dim strCalendarWorkshiftsFile As String
Dim strCalendarsJSON As String
Dim strCalendarsFile As String
'longs
Dim lngDayOfWeek As Long
Dim lngCalendarExceptionsFile As Long
Dim lngCalendarWorkshiftsFile As Long
Dim lngCalendarsFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'Calendars.json
  strCalendarsFile = strDir & "\Calendars.json"
  lngCalendarsFile = FreeFile
  If Dir(strCalendarsFile) <> vbNullString Then Kill strCalendarsFile
  Open strCalendarsFile For Output As #lngCalendarsFile
  strCalendarsJSON = "["
  
  'CalendarWorkshifts.json
  strCalendarWorkshiftsFile = strDir & "\CalendarWorkshifts.json"
  lngCalendarWorkshiftsFile = FreeFile
  If Dir(strCalendarWorkshiftsFile) <> vbNullString Then Kill strCalendarWorkshiftsFile
  Open strCalendarWorkshiftsFile For Output As #lngCalendarWorkshiftsFile
  strCalendarWorkshiftsJSON = "["
  
  'CalendarExceptions.json
  strCalendarExceptionsFile = strDir & "\CalendarExceptions.json"
  lngCalendarExceptionsFile = FreeFile
  If Dir(strCalendarExceptionsFile) <> vbNullString Then Kill strCalendarExceptionsFile
  Open strCalendarExceptionsFile For Output As #lngCalendarExceptionsFile
  strCalendarExceptionsJSON = "["
  
  If ActiveProject.Subprojects.Count = 0 Then
    
    For Each oCalendar In ActiveProject.BaseCalendars
      'todo: what about resource calendars?
    
      'Calendars.json
      strCalendarsJSON = strCalendarsJSON & "{"
      strCalendarsJSON = strCalendarsJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oCalendar.Index & Chr(34) & ","
      strCalendarsJSON = strCalendarsJSON & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & oCalendar.Name & Chr(34) & ","
      strCalendarsJSON = strCalendarsJSON & Chr(34) & "Comments" & Chr(34) & ": null},"
      
      'CalendarWorkshifts.json
      For Each oWorkshift In oCalendar.WorkWeeks
        strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & "{"
        strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oCalendar.Index & Chr(34) & ","
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
            strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oCalendar.Index & Chr(34) & ","
            strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "ExceptionDate" & Chr(34) & ": " & Chr(34) & Format(dtException, "yyyy-mm-dd") & Chr(34) & ","
            strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "WorkHours" & Chr(34) & ": null"
            strCalendarExceptionsJSON = strCalendarExceptionsJSON & "},"
          End If
          dtException = DateAdd("d", 1, dtException)
        Loop
      Next
      
    Next 'oCalendar
        
  Else
    For Each oSubProject In ActiveProject.Subprojects
      
      For Each oCalendar In oSubProject.SourceProject.BaseCalendars
        'todo: what about resource calendars?
      
        'Calendars.json
        strCalendarsJSON = strCalendarsJSON & "{"
        strCalendarsJSON = strCalendarsJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oCalendar.Index & Chr(34) & ","
        strCalendarsJSON = strCalendarsJSON & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & oCalendar.Name & Chr(34) & ","
        strCalendarsJSON = strCalendarsJSON & Chr(34) & "Comments" & Chr(34) & ": null},"
        
        'CalendarWorkshifts.json
        For Each oWorkshift In oCalendar.WorkWeeks
          strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & "{"
          strCalendarWorkshiftsJSON = strCalendarWorkshiftsJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oCalendar.Index & Chr(34) & ","
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
              strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oCalendar.Index & Chr(34) & ","
              strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "ExceptionDate" & Chr(34) & ": " & Chr(34) & Format(dtException, "yyyy-mm-dd") & Chr(34) & ","
              strCalendarExceptionsJSON = strCalendarExceptionsJSON & Chr(34) & "WorkHours" & Chr(34) & ": null"
              strCalendarExceptionsJSON = strCalendarExceptionsJSON & "},"
            End If
            dtException = DateAdd("d", 1, dtException)
          Loop
        Next
        
      Next 'oCalendar
    Next 'oSubProject
  End If
  
  'Calendars.json
  strCalendarsJSON = Left(strCalendarsJSON, Len(strCalendarsJSON) - 1) & "]"
  'todo: add both brackets at this step if len(x)>0 otherwise don't create it
  Print #lngCalendarsFile, strCalendarsJSON
  
  'CalendarWorkshifts.json
  strCalendarWorkshiftsJSON = Left(strCalendarWorkshiftsJSON, Len(strCalendarWorkshiftsJSON) - 1) & "]"
  Print #lngCalendarWorkshiftsFile, strCalendarWorkshiftsJSON
  
  'CalendarExceptions.json
  strCalendarExceptionsJSON = Left(strCalendarExceptionsJSON, Len(strCalendarExceptionsJSON) - 1) & "]"
  Print #lngCalendarExceptionsFile, strCalendarExceptionsJSON
  
  cptJSON_Calendars = True

exit_here:
  On Error Resume Next
  Set aWorkHours = Nothing
  Set oCalendarException = Nothing
  Set oWorkshift = Nothing
  Set oCalendar = Nothing
  Set oSubProject = Nothing
  Close #lngCalendarsFile
  Close #lngCalendarWorkshiftsFile
  Close #lngCalendarExceptionsFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_Calendars", Err, Erl)
  cptJSON_Calendars = False
  Resume exit_here
End Function

Function cptJSON_Tasks(strDir As String) As Boolean
'objects
Dim oTaskDependency As MSProject.TaskDependency
Dim oSubProject As MSProject.SubProject
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
Dim lngTaskConstratinsFile As Long
Dim lngTaskRelationshipsFile As Long
'integers
'doubles
'booleans
Dim blnDisplayProjectSummary As Boolean
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'todo: ensure task name uniqueness
  'todo: identify driving path(s) based on previous month's uid/guid, to either a) next contractor identified event; or b) gov selected event
  'todo: identify critical path(s)
  
  'TaskOutlineStructure.json
  strTaskOutlineStructureFile = strDir & "\TaskOutlineStructure.json"
  lngTaskOutlineStructureFile = FreeFile
  If Dir(strTaskOutlineStructureFile) <> vbNullString Then Kill strTaskOutlineStructureFile
  Open strTaskOutlineStructureFile For Output As #lngTaskOutlineStructureFile
  strTaskOutlineStructureJSON = "["
  
  'Tasks.json
  strTasksFile = strDir & "\Tasks.json"
  lngTasksFile = FreeFile
  If Dir(strTasksFile) <> vbNullString Then Kill strTasksFile
  Open strTasksFile For Output As #lngTasksFile
  strTasksJSON = "["
  
  'TaskScheduleData.json
  strTaskScheduleDataFile = strDir & "\TaskScheduleData.json"
  lngTaskScheduleDataFile = FreeFile
  If Dir(strTaskScheduleDataFile) <> vbNullString Then Kill strTaskScheduleDataFile
  Open strTaskScheduleDataFile For Output As #lngTaskScheduleDataFile
  strTaskScheduleDataJSON = "["
  
  'todo: TaskCustomFieldValues.json
  
  'TaskConstraints.json
  strTaskConstraintsFile = strDir & "\TaskConstraints.json"
  lngTaskConstraintsFile = FreeFile
  If Dir(strTaskConstraintsFile) <> vbNullString Then Kill strTaskConstraintsFile
  Open strTaskConstraintsFile For Output As #lngTaskConstraintsFile
  strTaskConstraintsJSON = "["
  
  'todo: TaskRelationships.json
  strTaskRelationshipsFile = strDir & "\TaskRelationships.json"
  lngTaskRelationshipsFile = FreeFile
  If Dir(strTaskRelationshipsFile) <> vbNullString Then Kill strTaskRelationshipsFile
  Open strTaskRelationshipsFile For Output As #lngTaskRelationshipsFile
  strTaskRelationshipsJSON = "["
   
  'todo: use a single text field to house string array of identification codes?
  
  'todo: display project summary to force single level 1
  Application.OptionsViewEx projectsummary:=True
  
  'capture project summary task
  Set oTask = ActiveProject.ProjectSummaryTask
  If ActiveProject.Subprojects.Count > 0 Then lngOutlineOffset = 1 Else lngOutlineOffset = 0
  
  'build TaskOutlineStructure.json
  strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "{"
  strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "Level" & Chr(34) & ": " & Chr(34) & CLng(oTask.OutlineLevel + lngOutlineOffset) & Chr(34) & ","
  strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
  strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "ParentTaskID" & Chr(34) & ": null,"
  strTaskOutlineStructureJSON = Left(strTaskOutlineStructureJSON, Len(strTaskOutlineStructureJSON) - 1)
  strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "},"
  
  'build Tasks.json
  strTasksJSON = strTasksJSON & "{"
  strTasksJSON = strTasksJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
  strTasksJSON = strTasksJSON & Chr(34) & "Name" & Chr(34) & ": " & Chr(34) & Trim(oTask.Name) & Chr(34) & ","
  strTasksJSON = strTasksJSON & Chr(34) & "TaskTypeID" & Chr(34) & ": " & Chr(34) & "SUMMARY" & Chr(34) & ","
  strTasksJSON = strTasksJSON & "},"
  
  'build TaskScheduleData.json
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & "{"
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
  If oTask.Calendar = "None" Then
    strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & ActiveProject.Calendar.Index & Chr(34) & ","
  Else
    strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oTask.CalendarObject.Index & Chr(34) & ","
  End If
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentDuration" & Chr(34) & ": " & Chr(34) & oTask.Duration / 60 & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyStart, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyFinish, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateStart, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateFinish, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FreeFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.FreeSlack / 60 & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "TotalFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.TotalSlack / 60 & Chr(34) & ","
  'todo: need flags for on critical path
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnCriticalPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
  'todo: need flags for on driving path
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnDrivingPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineDuration" & Chr(34) & ": " & Chr(34) & oTask.BaselineDuration / 60 & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineStart, "yyyy-mm-dd") & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineFinish, "yyyy-mm-dd") & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "StartVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.StartVariance / 60 & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FinishVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.FinishVariance / 60 & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalculatedPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PercentComplete & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "PhysicalPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PhysicalPercentComplete & Chr(34) & ","
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "RemainingDuration" & Chr(34) & ": " & Chr(34) & oTask.RemainingDuration / 60 & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualStart, "yyyy-mm-dd") & Chr(34) & ","
  'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualFinish, "yyyy-mm-dd") & Chr(34) & ","
  strTaskScheduleDataJSON = Left(strTaskScheduleDataJSON, Len(strTaskScheduleDataJSON) - 1)
  strTaskScheduleDataJSON = strTaskScheduleDataJSON & "},"
  
  'todo: account for elapsed durations
  
  If ActiveProject.Subprojects.Count = 0 Then
    For Each oTask In ActiveProject.Tasks
      If Not oTask Is Nothing Then
                
        'build TaskOutlineStructure.json
        strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "{"
        strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "Level" & Chr(34) & ": " & Chr(34) & CLng(oTask.OutlineLevel + lngOutlineOffset) & Chr(34) & ","
        strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
        If Not oTask.OutlineParent Is Nothing Then
          strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "ParentTaskID" & Chr(34) & ": " & Chr(34) & oTask.OutlineParent.UniqueID & Chr(34) & ","
        End If
        strTaskOutlineStructureJSON = Left(strTaskOutlineStructureJSON, Len(strTaskOutlineStructureJSON) - 1)
        strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "},"
              
        'build Tasks.json
        strTasksJSON = strTasksJSON & "{"
        strTasksJSON = strTasksJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
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
        'TaskPlanningLevelID: SUMMARY_LEVEL_PLANNING_PACKAGE,CONTROL_ACCOUNT,PLANNING_PACKAGE,WORK_PACKAGE,ACTIVITY
        'WBSElementID:
        'OBSElementID:
        'ControlAccountID:
        'WorkPackageID:
        'IMPElementID:
        'SOWReference:
        'SubcontractorReference:
        'EarnedValueTechniqueID: APPORTIONED_EFFORT,LEVEL_OF_EFFORT,MILESTONE,FIXED_0_100,FIXED_100_0,FIXED_X_Y,PERCENT_COMPLETE,STANDARDS,UNITS,OTHER_DISCRETE
        'OtherEarnedValueTechnique:
        'SourceSubprojectReference:
        'SourceTaskReference:
        'Comments:
        strTasksJSON = Left(strTasksJSON, Len(strTasksJSON) - 1)
        strTasksJSON = strTasksJSON & "},"
        
        'build TaskScheduleData.json
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & "{"
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
        If oTask.Calendar = "None" Then
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & ActiveProject.Calendar.Index & Chr(34) & ","
        Else
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oTask.CalendarObject.Index & Chr(34) & ","
        End If
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentDuration" & Chr(34) & ": " & Chr(34) & oTask.Duration / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyStart, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyFinish, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateStart, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateFinish, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FreeFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.FreeSlack / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "TotalFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.TotalSlack / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        'todo: need flags for on critical path
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnCriticalPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
        'todo: need flags for on driving path
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnDrivingPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineDuration" & Chr(34) & ": " & Chr(34) & oTask.BaselineDuration / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineStart, "yyyy-mm-dd") & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineFinish, "yyyy-mm-dd") & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "StartVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.StartVariance / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FinishVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.FinishVariance / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalculatedPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PercentComplete & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "PhysicalPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PhysicalPercentComplete & Chr(34) & ","
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "RemainingDuration" & Chr(34) & ": " & Chr(34) & oTask.RemainingDuration / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualStart, "yyyy-mm-dd") & Chr(34) & ","
        'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualFinish, "yyyy-mm-dd") & Chr(34) & ","
        strTaskScheduleDataJSON = Left(strTaskScheduleDataJSON, Len(strTaskScheduleDataJSON) - 1)
        strTaskScheduleDataJSON = strTaskScheduleDataJSON & "},"
        
        'build TaskConstraints.json
        If oTask.ConstraintType <> pjASAP Then
          strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
          
          Select Case oTask.ConstraintType
  
            Case pjALAP
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "AS_LATE_AS_POSSIBLE" & Chr(34) & ","
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": null"
  
            Case pjMSO
              If ActiveProject.HonorConstraints Then
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "MUST_START_ON" & Chr(34) & ","
              Else
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "SHOULD_START_ON" & Chr(34) & ","
              End If
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
              strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)
  
            Case pjMFO
              If ActiveProject.HonorConstraints Then
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
              If ActiveProject.HonorConstraints Then
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
              If ActiveProject.HonorConstraints Then
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
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "RESOURCE_LEVELING_START_DELAY" & Chr(34) & ","
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34)
          strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
          
          'resource leveling finish delay
          strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "RESOURCE_LEVELING_FINISH_DELAY" & Chr(34) & ","
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34)
          strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
          
        End If 'oTask.LevelingDelay
        
        'TaskConstraints.json - deadline
        If IsDate(oTask.Deadline) Then 'new record
          strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oTask.UniqueID & Chr(34) & ","
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "DEADLINE" & Chr(34) & ","
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
          strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Deadline, "yyyy-mm-dd") & Chr(34)
          strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
        End If 'IsDate(oTask.Deadline)
        
        'build TaskRelationships.json
        For Each oTaskDependency In oTask.TaskDependencies
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & "{"
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "PredecessorTaskID" & Chr(34) & ": " & Chr(34) & oTaskDependency.From.UniqueID & Chr(34) & ","
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "SuccessorTaskID" & Chr(34) & ": " & Chr(34) & oTaskDependency.To.UniqueID & Chr(34) & ","
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
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagDuration" & Chr(34) & ": " & Chr(34) & oTaskDependency.Lag / (ActiveProject.HoursPerDay * 60) & Chr(34) & ","
          If oTaskDependency.To.Calendar <> "None" Then
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagCalendarID" & Chr(34) & ": " & Chr(34) & ActiveProject.Calendar.Index & Chr(34)
          Else
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagCalendarID" & Chr(34) & ": " & Chr(34) & oTaskDependency.To.CalendarObject.Index & Chr(34)
          End If
          strTaskRelationshipsJSON = strTaskRelationshipsJSON & "},"
        Next 'oTaskDependency
        
      End If 'Not oTask Is Nothing Then
next_task:
    Next 'oTask
  
  Else
        
    For Each oSubProject In ActiveProject.Subprojects
      For Each oTask In oSubProject.SourceProject.Tasks 'different
        If Not oTask Is Nothing Then
          
          'build TaskOutlineStructure.json
          strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "{"
          strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "Level" & Chr(34) & ": " & Chr(34) & CLng(oTask.OutlineLevel + lngOutlineOffset) & Chr(34) & ","
          strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & "," 'different
          If Not oTask.OutlineParent Is Nothing Then
            If oTask.OutlineParent.UniqueID = 0 Then
              strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "ParentTaskID" & Chr(34) & ": " & Chr(34) & oTask.OutlineParent.UniqueID & Chr(34) & ","
            Else
              strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & Chr(34) & "ParentTaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.OutlineParent.UniqueID & Chr(34) & ","
            End If
          End If
          strTaskOutlineStructureJSON = Left(strTaskOutlineStructureJSON, Len(strTaskOutlineStructureJSON) - 1)
          strTaskOutlineStructureJSON = strTaskOutlineStructureJSON & "},"
                
          'build Tasks.json
          strTasksJSON = strTasksJSON & "{"
          strTasksJSON = strTasksJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
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
          'TaskPlanningLevelID: SUMMARY_LEVEL_PLANNING_PACKAGE,CONTROL_ACCOUNT,PLANNING_PACKAGE,WORK_PACKAGE,ACTIVITY
          'WBSElementID:
          'OBSElementID:
          'ControlAccountID:
          'WorkPackageID:
          'IMPElementID:
          'SOWReference:
          'SubcontractorReference:
          'EarnedValueTechniqueID: APPORTIONED_EFFORT,LEVEL_OF_EFFORT,MILESTONE,FIXED_0_100,FIXED_100_0,FIXED_X_Y,PERCENT_COMPLETE,STANDARDS,UNITS,OTHER_DISCRETE
          'OtherEarnedValueTechnique:
          'SourceSubprojectReference:
          'SourceTaskReference:
          'Comments:
          strTasksJSON = Left(strTasksJSON, Len(strTasksJSON) - 1)
          strTasksJSON = strTasksJSON & "},"
          
          'build TaskScheduleData.json
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & "{"
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
          'todo: reconcile calendars - perhaps add subproject.index-prefix
          If oTask.Calendar = "None" Then
            strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oSubProject.SourceProject.Calendar.Index & Chr(34) & ","
          Else
            strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalendarID" & Chr(34) & ": " & Chr(34) & oTask.CalendarObject.Index & Chr(34) & ","
          End If
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentDuration" & Chr(34) & ": " & Chr(34) & oTask.Duration / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CurrentFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyStart, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "EarlyFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.EarlyFinish, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateStart, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "LateFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.LateFinish, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FreeFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.FreeSlack / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "TotalFloatDuration" & Chr(34) & ": " & Chr(34) & oTask.TotalSlack / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          'todo: need flags for on critical path
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnCriticalPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
          'todo: need flags for on driving path
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "OnDrivingPath" & Chr(34) & ": " & Chr(34) & oTask.Critical & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineDuration" & Chr(34) & ": " & Chr(34) & oTask.BaselineDuration / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineStart, "yyyy-mm-dd") & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "BaselineFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.BaselineFinish, "yyyy-mm-dd") & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "StartVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.StartVariance / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "FinishVarianceDuration" & Chr(34) & ": " & Chr(34) & oTask.FinishVariance / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "CalculatedPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PercentComplete & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "PhysicalPercentComplete" & Chr(34) & ": " & Chr(34) & oTask.PhysicalPercentComplete & Chr(34) & ","
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "RemainingDuration" & Chr(34) & ": " & Chr(34) & oTask.RemainingDuration / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualStartDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualStart, "yyyy-mm-dd") & Chr(34) & ","
          'strTaskScheduleDataJSON = strTaskScheduleDataJSON & Chr(34) & "ActualFinishDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ActualFinish, "yyyy-mm-dd") & Chr(34) & ","
          strTaskScheduleDataJSON = Left(strTaskScheduleDataJSON, Len(strTaskScheduleDataJSON) - 1)
          strTaskScheduleDataJSON = strTaskScheduleDataJSON & "},"
          
          'TaskConstraints.json
          If oTask.ConstraintType <> pjASAP Then
            strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
            
            Select Case oTask.ConstraintType
    
              Case pjALAP
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "AS_LATE_AS_POSSIBLE" & Chr(34) & ","
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": null"
    
              Case pjMSO
                If oSubProject.SourceProject.HonorConstraints Then
                  strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "MUST_START_ON" & Chr(34) & ","
                Else
                  strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "SHOULD_START_ON" & Chr(34) & ","
                End If
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
                strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.ConstraintDate, "yyyy-mm-dd") & Chr(34)
    
              Case pjMFO
                If oSubProject.SourceProject.HonorConstraints Then
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
                If oSubProject.SourceProject.HonorConstraints Then
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
                If oSubProject.SourceProject.HonorConstraints Then
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
          If oTask.LevelingDelay > 0 Then
            
            'resource leveling start delay
            strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "RESOURCE_LEVELING_START_DELAY" & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Start, "yyyy-mm-dd") & Chr(34)
            strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
            
            'resource leveling finish delay
            strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "RESOURCE_LEVELING_FINISH_DELAY" & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Finish, "yyyy-mm-dd") & Chr(34)
            strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
            
          End If 'oTask.LevelingDelay
          
          'TaskConstraints.json - deadline
          If IsDate(oTask.Deadline) Then 'new record
            strTaskConstraintsJSON = strTaskConstraintsJSON & "{"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTask.UniqueID & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintTypeID" & Chr(34) & ": " & Chr(34) & "DEADLINE" & Chr(34) & ","
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "OtherConstraintType" & Chr(34) & ": null,"
            strTaskConstraintsJSON = strTaskConstraintsJSON & Chr(34) & "ConstraintDate" & Chr(34) & ": " & Chr(34) & Format(oTask.Deadline, "yyyy-mm-dd") & Chr(34)
            strTaskConstraintsJSON = strTaskConstraintsJSON & "},"
          End If 'IsDate(oTask.Deadline)
          
          'build TaskRelationships.json
          For Each oTaskDependency In oTask.TaskDependencies
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & "{"
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "PredecessorTaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTaskDependency.From.UniqueID & Chr(34) & ","
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "SuccessorTaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTaskDependency.To.UniqueID & Chr(34) & ","
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
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagDuration" & Chr(34) & ": " & Chr(34) & oTaskDependency.Lag / (oSubProject.SourceProject.HoursPerDay * 60) & Chr(34) & ","
            If oTaskDependency.To.Calendar <> "None" Then
              strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagCalendarID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oTaskDependency.To.CalendarObject.Index & Chr(34)
            Else
              strTaskRelationshipsJSON = strTaskRelationshipsJSON & Chr(34) & "LagCalendarID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oSubProject.SourceProject.Calendar.Index & Chr(34)
            End If
            strTaskRelationshipsJSON = strTaskRelationshipsJSON & "},"
          Next 'oTaskDependency
          
        End If
      Next 'oTask
    Next 'oSubProject
  End If 'ActiveProject.Subprojects.Count = 0
    
  'create Tasks.json
  strTasksJSON = Left(strTasksJSON, Len(strTasksJSON) - 1) & "]"
  Print #lngTasksFile, strTasksJSON
  
  'create TaskOutlineStructure.json
  strTaskOutlineStructureJSON = Left(strTaskOutlineStructureJSON, Len(strTaskOutlineStructureJSON) - 1) & "]"
  Print #lngTaskOutlineStructureFile, strTaskOutlineStructureJSON
  
  'create TaskScheduleData.json
  strTaskScheduleDataJSON = Left(strTaskScheduleDataJSON, Len(strTaskScheduleDataJSON) - 1) & "]"
  Print #lngTaskScheduleDataFile, strTaskScheduleDataJSON
  
  'create TaskConstraints.json
  strTaskConstraintsJSON = Left(strTaskConstraintsJSON, Len(strTaskConstraintsJSON) - 1) & "]"
  Print #lngTaskConstraintsFile, strTaskConstraintsJSON
  
  'create TaskRelationships.json
  strTaskRelationshipsJSON = Left(strTaskRelationshipsJSON, Len(strTaskRelationshipsJSON) - 1) & "]"
  Print #lngTaskRelationshipsFile, strTaskRelationshipsJSON
  
  cptJSON_Tasks = True

exit_here:
  On Error Resume Next
  Set oTaskDependency = Nothing
  Set oSubProject = Nothing
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

Function cptJSON_Resources(strDir As String) As Boolean
'objects
Dim oAssignment As MSProject.Assignment
Dim oSubProject As MSProject.SubProject
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
  If Dir(strResourcesFile) <> vbNullString Then Kill strResourcesFile
  Open strResourcesFile For Output As #lngResourcesFile
  strResourcesJSON = "["
  
  'ResourceAssignments.json
  strResourceAssignmentsFile = strDir & "\ResourceAssignments.json"
  lngResourceAssignmentsFile = FreeFile
  If Dir(strResourceAssignmentsFile) <> vbNullString Then Kill strResourceAssignmentsFile
  Open strResourceAssignmentsFile For Output As #lngResourceAssignmentsFile
  strResourceAssignmentsJSON = "["
  
  'todo: ResourceCustomFieldValues.json
  
  For Each oResource In ActiveProject.Resources
  
    'Resources.json
    strResourcesJSON = strResourcesJSON & "{"
    strResourcesJSON = strResourcesJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oResource.UniqueID & Chr(34) & ","
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
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "ResourceID" & Chr(34) & ": " & Chr(34) & oResource.UniqueID & Chr(34) & ","
      strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oAssignment.TaskUniqueID & Chr(34) & ","
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
   
  For Each oSubProject In ActiveProject.Subprojects
    For Each oResource In oSubProject.SourceProject.Resources
    
      'Resources.json
      strResourcesJSON = strResourcesJSON & "{"
      strResourcesJSON = strResourcesJSON & Chr(34) & "ID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oResource.UniqueID & Chr(34) & ","
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
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "ResourceID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oResource.UniqueID & Chr(34) & ","
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "TaskID" & Chr(34) & ": " & Chr(34) & oSubProject.Index & "-" & oAssignment.TaskUniqueID & Chr(34) & ","
        If oResource.Type = pjResourceTypeWork Then
          'Budget_AtCompletion_Dollars
          'Budget_AtCompletion_Hours
          strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Budget_AtCompletion_Hours" & Chr(34) & ": " & Chr(34) & oAssignment.BaselineWork / 60 & Chr(34) & ","
          'Budget_AtCompletion_Hours
          'Estimate_ToComplete_Dollars
          'Estimate_ToComplete_Hours
          strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Estimate_ToComplete_Hours" & Chr(34) & ": " & Chr(34) & oAssignment.RemainingWork / 60 & Chr(34) & ","
          'Estimate_ToComplete_Hours
          'Actual_ToDate_Dollars
          'Actual_ToDate_Hours
        Else
          'Budget_AtCompletion_Dollars
          strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Budget_AtCompletion_Dollars" & Chr(34) & ": " & Chr(34) & oAssignment.BaselineCost & Chr(34) & ","
          'Budget_AtCompletion_Hours
          'Estimate_ToComplete_Dollars
          strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "Estimate_ToComplete_Dollars" & Chr(34) & ": " & Chr(34) & oAssignment.RemainingCost & Chr(34) & ","
          'Estimate_ToComplete_Hours
          'Actual_ToDate_Dollars
          'Actual_ToDate_Hours
        End If
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & Chr(34) & "PhysicalPercentComplete" & Chr(34) & ": " & Chr(34) & oAssignment.Task.PhysicalPercentComplete & Chr(34)
        strResourceAssignmentsJSON = strResourceAssignmentsJSON & "},"
      Next 'oAssignment
      
    Next 'oResource
  
  Next 'oSubProject
    
  'create Resources.json
  strResourcesJSON = Left(strResourcesJSON, Len(strResourcesJSON) - 1) & "]"
  Print #lngResourcesFile, strResourcesJSON
  
  'create ResourceAssignments.json
  strResourceAssignmentsJSON = Left(strResourceAssignmentsJSON, Len(strResourceAssignmentsJSON) - 1) & "]"
  Print #lngResourceAssignmentsFile, strResourceAssignmentsJSON
  
  cptJSON_Resources = True
  
exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Set oSubProject = Nothing
  Set oResource = Nothing
  Close #lngResourcesFile
  Close #lngResourceAssignmentsFile
  Exit Function
err_here:
  Call cptHandleErr("cptIPMDAR", "cptJSON_Resources", Err, Erl)
  cptJSON_Resources = False
  Resume exit_here
End Function

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
