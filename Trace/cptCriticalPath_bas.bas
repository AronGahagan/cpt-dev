Attribute VB_Name = "cptCriticalPath_bas"
'<cpt_version>v2.7</cpt_version>

Option Explicit

Private CritField As String 'Stores comma seperated values for each task showing which paths they are a part of
Private GroupField As String 'Stores a single value - used to group/sort tasks in final CP view

'Custom type used to store driving path vars
Type DrivingPaths

    PrimaryFloat As Double 'Stores True Float vlaue for Primary Driving Path
    FindPrimary As Boolean 'Tracks evluation progress by noting Primary Found
    SecondaryFloat As Double 'Stores True Float vlaue for Secondary Driving Path
    FindSecondary As Boolean 'Tracks evluation progress by noting Secondary Found
    TertiaryFloat As Double 'Stores True Float vlaue for Secondary Driving Path
    FindTertiary As Boolean 'Tracks evluation progress by noting Tertiar Found

End Type

Private tDrivingPaths As DrivingPaths 'var to store DrivingPaths type
Private SecondaryDrivers() As String 'Array of Secondary Drivers to be analyzed
Private SecondaryDriverCount As Integer 'Count of secondary Drivers
Private TertiaryDrivers() As String 'Array of tertiary drivers to be analyzed
Private TertiaryDriverCount As Integer 'Count of tertiary drivers
Private AnalyzedTasks As Collection 'Collection of task relationships analyzied (From UID - To UID); unique to each path analysis

'Custom type used to store Driving Task data
Type DrivingTask

    UID As String
    tFloat As Double

End Type

Private DrivingTasks() As DrivingTask 'var to store DrivingTask type
Private drivingTasksCount As Integer 'coung of DrivingTasks
'/ag edit private > public\
Public export_to_PPT As Boolean 'cpt ToolBar controlled var for controlling user notification of completed analysis

Sub DrivingPaths()
'Primary analysis module that controls analysis
'workflow through Primary, Secondary and Tertiary
'driving paths.

    Dim curproj As Project 'Stores active user project - not compatible with Master/Sub Architecture
    Dim t As Task 'Stores initial user selected task
    Dim tdp As TaskDependency
    Dim tdps As TaskDependencies
    Dim i As Integer 'Used to iterate through Primary/Secondary/Tertiary driver arrays
    Dim analysisTaskUID As String 'Stores user selected task for recall and selection after setting final view
    
    'Hardcoded field requirements
    CritField = "Text29"
    GroupField = "Number19"
    
    'Store users active project
    Set curproj = ActiveProject
    
    'used to avoid code break during intial error checks
    On Error Resume Next
    
    'Validate users selected view type
    If curproj.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
        MsgBox "Please select a View with a Task Table."
        curproj = Nothing
        Exit Sub
    End If
    
    'Validate users selected window pane - select the task table if not active
    If curproj.Application.ActiveWindow.ActivePane.Index <> 1 Then
        curproj.Application.ActiveWindow.TopPane.Activate
    End If
    
    'Exit if multiple tasks are selected
    If curproj.Application.ActiveSelection.Tasks.count > 1 Then
        MsgBox "Select a single activity only."
        curproj = Nothing
        Exit Sub
    End If
    
    'store task of activeselection
    Set t = curproj.Application.ActiveCell.Task
    
    'Check for null task rows
    If t Is Nothing Then
        MsgBox "Select a task"
        curproj = Nothing
        Exit Sub
    End If
    
    'Avoid analyzing completed tasks
    If t.PercentComplete = 100 Then
        MsgBox "Select an incomplete task"
        curproj = Nothing
        Exit Sub
    End If
    
    'Avoid analysis on summary rows
    If t.Summary = True Then
        MsgBox "Select a non-summary task"
        curproj = Nothing
        Exit Sub
    End If
    
    'Suspend calculations and screen updating
    curproj.Application.Calculation = pjManual
    curproj.Application.ScreenUpdating = False
    
    On Error GoTo CleanUp
    
    '**********************************************
    'On Error GoTo 0 '*****used for debug only*****
    '**********************************************
    
    'Assign Custom Field names and create lookup table
    SetGroupCPFieldLookupTable GroupField, curproj
    
    'Erase previous Crit and Group field values
    CleanCritFlag curproj
    
    'Erase any previously created/modified view elements
    CleanViews curproj
    
    'Initialize Analyzed Tasks Collection
    Set AnalyzedTasks = New Collection
    
    'Add selected task to Analyzed Tasks collection and store UID for later reference
    AnalyzedTasks.Add t.UniqueID, t.UniqueID & "-" & t.UniqueID
    analysisTaskUID = t.UniqueID

    'Set default Float values
    tDrivingPaths.PrimaryFloat = 0
    tDrivingPaths.SecondaryFloat = 0
    tDrivingPaths.TertiaryFloat = 0
    
    'Now finding Primary Path
    tDrivingPaths.FindPrimary = True
    tDrivingPaths.FindSecondary = False
    tDrivingPaths.FindTertiary = False
    
    'Set default driver counts
    SecondaryDriverCount = 0
    TertiaryDriverCount = 0
    drivingTasksCount = 0
    
    '********************************
    '***Find Primary Driving Paths***
    '********************************
    
    'Store dependencies of user selected task
    Set tdps = t.TaskDependencies
    
    'Note that selected task is visible on paths 1,2,3
    t.SetField FieldNameToFieldConstant(CritField), "1,2,3"
    t.SetField FieldNameToFieldConstant(GroupField), "1"
    
    'Evlauate list of dependencies on selected analysis task
    For Each tdp In tdps
    
        'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
        evaluateTaskDependencies tdp, t, curproj, AnalyzedTasks
        
    Next tdp 'Next user selected analysis task dependency
    
    'Clear variables for re-use in evaluating secondary driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find sendary path driver
    FindNextDriver
    
    '**********************************
    '***Find Secondary Driving Paths***
    '**********************************
    
    'Find Secondary if one exists (driver count is greater than 0)
    If SecondaryDriverCount > 0 Then
    
        'Note that we are now evaluating the secondary driving path
        tDrivingPaths.FindSecondary = True
        
        'iterate through list of secondary drivers
        For i = 1 To SecondaryDriverCount
        
            'avoid evaluating outside of the array bounds
            If i > SecondaryDriverCount Then Exit For
            
            'store the current driving task
            Set t = curproj.Tasks.UniqueID(SecondaryDrivers(i))
            
            'add the driving task to the analyzed tasks collection
            AnalyzedTasks.Add t.UniqueID & "-" & t.UniqueID
            
            'If the task has not already been analyzed during previous path analysis,
            'set the Crit and Group Field values
            If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                With t
                    .SetField FieldNameToFieldConstant(CritField), "2"
                    .SetField FieldNameToFieldConstant(GroupField), "2"
                End With
            Else
            
                'If the task has already been analyzed during the previous path analysis,
                'append path value to the Crit and Group Fields
                If InStr(t.GetField(FieldNameToFieldConstant(CritField)), "2") = 0 Then
                    t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & ",2"
                End If
                
            End If
            
            'Store secondary driving task dependencies
            Set tdps = t.TaskDependencies
            
            'Evlauate list of dependencies on secondary driving task
            For Each tdp In tdps
            
                'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                evaluateTaskDependencies tdp, t, curproj, AnalyzedTasks
                
            Next tdp 'Next secondary driver dependency
            
        Next i 'next Secondary Path Driver
        
    End If
    
    'Clear variables for re-use in evaluating secondary driver
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Set AnalyzedTasks = New Collection
    
    'Iterate through drivingtasks array to find sendary path driver
    FindNextDriver
    
    '*********************************
    '***Find Tertiary Driving Paths***
    '*********************************
    
    'Find tertiary if one exists (driver count is greater than 0)
    If TertiaryDriverCount > 0 Then
    
        'Note that we are now evaluating the tertiary driving path
        tDrivingPaths.FindTertiary = True
        
        'iterate through list of tertiary drivers
        For i = 1 To TertiaryDriverCount
        
            'avoid evaluating outside of the array bounds
            If i > TertiaryDriverCount Then Exit For
            
            'store the current driving task
            Set t = curproj.Tasks.UniqueID(TertiaryDrivers(i))
            
            'add the driving task to the analyzed tasks collection
            AnalyzedTasks.Add t.UniqueID & "-" & t.UniqueID
            
            'If the task has not already been analyzed during previous path analysis,
            'set the Crit and Group Field values
            If t.GetField(FieldNameToFieldConstant(CritField)) = vbNullString Then
                With t
                    .SetField FieldNameToFieldConstant(CritField), "3"
                    .SetField FieldNameToFieldConstant(GroupField), "3"
                End With
            Else
            
                'If the task has already been analyzed during the previous path analysis,
                'append path value to the Crit and Group Fields
                If InStr(t.GetField(FieldNameToFieldConstant(CritField)), "3") = 0 Then
                    t.SetField FieldNameToFieldConstant(CritField), t.GetField(FieldNameToFieldConstant(CritField)) & ",3"
                End If
                
            End If
            
            'Store tertiary driving task dependencies
            Set tdps = t.TaskDependencies
            
            'Evlauate list of dependencies on tertiary driving task
            For Each tdp In tdps
            
                'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
                evaluateTaskDependencies tdp, t, curproj, AnalyzedTasks
                
            Next tdp 'Next tertiary driver dependency
            
        Next i 'next Tertiary Path Driver
        
    End If
    
    'Create and Apply the "ClearPlan Driving Path" Table, View, Group, and Filter
    SetupCPView GroupField, curproj, analysisTaskUID
    
CleanUp:

    'If error encountered, alert the user, otherwise notify of completion
    If err Then
        MsgBox "Error Encountered"
    Else
        If Not (export_to_PPT) Then MsgBox "Complete", vbOKOnly, "ClearPlan Critical Path Analyzer"
    End If

    'Clear variables
    Set tdps = Nothing
    Set tdp = Nothing
    Set t = Nothing
    Erase SecondaryDrivers, TertiaryDrivers, DrivingTasks
    Set AnalyzedTasks = Nothing
    SecondaryDriverCount = 0
    TertiaryDriverCount = 0
    drivingTasksCount = 0
    Set AnalyzedTasks = Nothing
    
    'Enable calculations and screenupdating
    curproj.Application.Calculation = pjAutomatic
    curproj.Application.ScreenUpdating = True
    
    'release project variable
    Set curproj = Nothing

End Sub

Private Sub evaluateTaskDependencies(ByVal tdp As TaskDependency, ByVal t As Task, ByVal curproj As Project, ByRef curAnalyzedTasks As Collection)
'Evaluate each task dependency, ignoring complete preds, then store as an analyzed relationship and evaluate criticality

    'Only evaluate incomplete predecessors
    If tdp.To.UniqueID = t.UniqueID And tdp.From.PercentComplete <> 100 Then
        'Check dependency for existance in analyzed tasks collection
        If ExistsInCollection(curAnalyzedTasks, curproj.Tasks.UniqueID(tdp.From).UniqueID & "-" & curproj.Tasks.UniqueID(tdp.To).UniqueID) = False Then
            'If dependency has not been analyzed, add to analyzed tasks collection
            curAnalyzedTasks.Add curproj.Tasks.UniqueID(tdp.From).UniqueID, curproj.Tasks.UniqueID(tdp.From).UniqueID & "-" & curproj.Tasks.UniqueID(tdp.To).UniqueID
            'Calculate True Float value and evaluate against list of driving tasks
            CheckCritTask curproj, tdp
        End If
    End If
    
End Sub

Private Sub SetGroupCPFieldLookupTable(ByVal GroupField As String, ByVal curproj As Project)
'Set Crit and Group field names, assign lookup table to Group Field

    'Store Field Names
    curproj.Application.CustomFieldRename FieldID:=FieldNameToFieldConstant(CritField), NewName:="CP Driving Paths"
    curproj.Application.CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField), NewName:="CP Driving Path Group ID"
    
    'Setup Lookup Table Properties
    curproj.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(GroupField), Attribute:=pjFieldAttributeNone
    curproj.Application.CustomOutlineCodeEditEx FieldID:=FieldNameToFieldConstant(GroupField), OnlyLookUpTableCodes:=True, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
    curproj.Application.CustomFieldPropertiesEx FieldID:=FieldNameToFieldConstant(GroupField), Attribute:=pjFieldAttributeValueList, SummaryCalc:=pjCalcNone, GraphicalIndicators:=False, AutomaticallyRolldownToAssn:=False
    
    'Assign Lookup Table Values
    curproj.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "1", "Primary"
    curproj.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "2", "Secondary"
    curproj.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "3", "Tertiary"
    curproj.Application.CustomFieldValueListAdd FieldNameToFieldConstant(GroupField), "0", "Noncritical"


End Sub
Private Sub SetupCPView(ByVal GroupField As String, ByVal curproj As Project, ByVal tUID As String)
'Setup CP View with Table & Grouping by Path Value

    Dim t As Task 'used to store user selected anlaysis task
    
    'Create CP Driving Path Table
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, Create:=True, ShowAddNewColumn:=True, OverwriteExisting:=True, FieldName:="ID", Width:=5, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, LockFirstColumn:=True, ColumnPosition:=0
    
    'Add fields to CP Driving Path Table
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Unique ID", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=1, LockFirstColumn:=True
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:=GroupField, Title:="Driving Path", Width:=5, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=1
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Name", Width:=45, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=2
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Duration", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=3
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Start", Width:=15, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=4
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Finish", Width:=15, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=5
    curproj.Application.TableEditEx Name:="*ClearPlan Driving Path Table", TaskTable:=True, NewFieldName:="Total Slack", Width:=10, ShowInMenu:=False, DateFormat:=pjDate_mm_dd_yy, ColumnPosition:=6

    'Create CP Driving Path Filter
    curproj.Application.FilterEdit Name:="*ClearPlan Driving Path Filter", Taskfilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=GroupField, test:="is greater than", Value:="0", ShowInMenu:=False, ShowSummaryTasks:=False
    
    'On Error Resume Next
    
    'Create CP Driving Path Group
    curproj.TaskGroups.Add Name:="*ClearPlan Driving Path Group", FieldName:=GroupField
    
    'Create CP Driving Path view if necessary
    curproj.Application.ViewEditSingle Name:="*ClearPlan Driving Path View", Create:=True, ShowInMenu:=True, Table:="*ClearPlan Driving Path Table", Filter:="*ClearPlan Driving Path Filter", Group:="*ClearPlan Driving Path Group"
    
    'Apply the CP Driving Path view
    curproj.Application.ViewApply Name:="*ClearPlan Driving Path View"
    
    'Sort the View by Finish, then by Duration to produce Waterfall Gantt
    curproj.Application.Sort Key1:="Finish", Ascending1:=True, Key2:="Duration", Ascending2:=False, outline:=False
    
    'Select all tasks and zoom the Gantt to display all tasks in view
    curproj.Application.SelectAll
    curproj.Application.ZoomTimescale Selection:=True
    
    curproj.Application.SelectRow 1
    
    'Iterate through each task in view and color the Gantt bars based on CP Group Code
    For Each t In ActiveProject.Tasks
        
        Select Case t.GetField(FieldNameToFieldConstant(GroupField))
        
            Case "1"
                t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=192, MiddleColor:=192, EndColor:=192
    
            Case "2"
                t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=3243501, MiddleColor:=3243501, EndColor:=3243501
            
            Case "3"
                t.Application.GanttBarFormatEx TaskID:=t.ID, GanttStyle:=1, StartColor:=65535, MiddleColor:=65535, EndColor:=65535
            
            Case Else
        
        End Select
    
    Next t
    
    'select the users original analysis task
    curproj.Application.FindEx "UniqueID", "equals", tUID

End Sub
Private Sub CleanCritFlag(ByVal curproj As Project)
'Remove previous analysis values from the Crit and Group fields

    Dim t As Task 'store task var
    
    'iterate through every task in the project
    For Each t In curproj.Tasks
        
        'Reset values
        t.SetField FieldNameToFieldConstant(CritField), vbNullString
        t.SetField FieldNameToFieldConstant(GroupField), "0"
    
    Next t

End Sub

Private Sub CleanViews(ByVal curproj As Project)
'Iterate through all Views, Tables, Filters, and Groups
'Delete previously created CP View Elements to avoid user modification errors

    Dim cpView As View
    Dim allViews As Views
    Dim cpTable As Table
    Dim allTables As Tables
    Dim cpFilter As Filter
    Dim allFilters As Filters
    Dim cpGroup As Group
    Dim allGroups As Groups
    
    'Set vars
    Set allViews = curproj.Views
    Set allTables = curproj.TaskTables
    Set allFilters = curproj.TaskFilters
    Set allGroups = curproj.TaskGroups
    
    'If the CPCritPathView is active, choose a different view
    curproj.Application.ViewApply Name:="Gantt Chart"

    'Clean up Views
    For Each cpView In allViews
        If cpView.Name = "*ClearPlan Driving Path View" Then
            cpView.Delete
            Exit For
        End If
    Next cpView
    
    'Clean up Tables
    For Each cpTable In allTables
        If cpTable.Name = "*ClearPlan Driving Path Table" Then
            cpTable.Delete
            Exit For
        End If
    Next cpTable
    
    'Clean up Filters
    For Each cpFilter In allFilters
        If cpFilter.Name = "*ClearPlan Driving Path Filter" Then
            cpFilter.Delete
            Exit For
        End If
    Next cpFilter
    
    'Clean up Groups
    For Each cpGroup In allGroups
        If cpGroup.Name = "*ClearPlan Driving Path Group" Then
            cpGroup.Delete
            Exit For
        End If
    Next cpGroup

End Sub
Private Function alreadyFound(ByVal t As Task) As Boolean
'Check for existing values in the Crit Field - if found, task has been evaluated previously

    If t.GetField(FieldNameToFieldConstant(CritField)) <> vbNullString Then
        alreadyFound = True
    Else
        alreadyFound = False
    End If
    
End Function

Private Sub FindNextDriver()
'Iterate through Driving Tasks array to find driving tasks based on True Float value

    Dim i As Integer 'Counter used to iterate through DrivingTasks array
    Dim driverCount As Integer 'count of driving tasks found
    Dim driverFloat As Double 'float value of driving tasks

    'If no drivers were found, exit the subroutine
    If drivingTasksCount = 0 Then
        Exit Sub
    End If

    'Find Secondary Driving Task if the Find Secondary has not yet been set to True
    If tDrivingPaths.FindSecondary = False Then
        
        'Store default float and count values
        driverFloat = 0
        driverCount = 0
    
        'Iterate through Driving Tasks array and find the least float value
        For i = 1 To UBound(DrivingTasks)
        
            'store first float value, otherwise evaluate current float value against previously stored value
            If driverFloat = 0 And DrivingTasks(i).tFloat <> 0 Then
                driverFloat = DrivingTasks(i).tFloat
            Else
                With DrivingTasks(i)
                    If .tFloat < driverFloat And .tFloat <> 0 Then
                        driverFloat = .tFloat
                    End If
                End With
            End If
        Next i 'Next Driving Task
        
        'Find all drivers with similar float and store as parallel driving tasks
        If driverFloat <> 0 Then
            For i = 1 To UBound(DrivingTasks)
                With DrivingTasks(i)
                    If .tFloat = driverFloat Then
                        driverCount = driverCount + 1
                        ReDim Preserve SecondaryDrivers(1 To driverCount)
                        SecondaryDrivers(driverCount) = .UID
                    End If
                End With
            Next i 'Next Driving Task
        End If
        
        'Set Secondary Float value equal to the evaluated driving task float
        tDrivingPaths.SecondaryFloat = driverFloat
        
        'set secondary driver count
        SecondaryDriverCount = driverCount
        
    Else 'If FindSecondary = True, indicates secondary path has been evaluated, so find the Tertiary Driving Task
    
        'Store default float and count values
        driverFloat = 0
        driverCount = 0
    
        'Iterate through Driving Tasks array and find the least float value
        For i = 1 To UBound(DrivingTasks)
        
            'store first float value, otherwise evaluate current float value against previously stored value
            If DrivingTasks(i).tFloat > tDrivingPaths.SecondaryFloat And driverFloat = 0 Then
                driverFloat = DrivingTasks(i).tFloat
            Else
                If DrivingTasks(i).tFloat > tDrivingPaths.SecondaryFloat And DrivingTasks(i).tFloat < driverFloat Then
                    driverFloat = DrivingTasks(i).tFloat
                End If
            End If
        Next i 'Next driving task
        
        'Find all drivers with similar float and store as parallel driving tasks
        If driverFloat <> 0 Then
            For i = 1 To UBound(DrivingTasks)
                With DrivingTasks(i)
                    If .tFloat = driverFloat Then
                        driverCount = driverCount + 1
                        ReDim Preserve TertiaryDrivers(1 To driverCount)
                        TertiaryDrivers(driverCount) = .UID
                    End If
                End With
            Next i 'Next Driving Task
        End If
        
        'Set Secondary Float value equal to the evaluated driving task float
        tDrivingPaths.TertiaryFloat = driverFloat
        
        'set secondary driver count
        TertiaryDriverCount = driverCount
    
    End If

End Sub

Private Function FindInArray(UID As String) As Variant
'Search DrivingTasks array for a task UID

    Dim i As Long 'counter to iterate through Driving Tasks
    
    For i = LBound(DrivingTasks) To UBound(DrivingTasks)
        If DrivingTasks(i).UID = UID Then
            FindInArray = i
            Exit Function
        End If
    Next i

    FindInArray = Null

End Function

Private Sub CheckCritTask(ByVal curproj As Project, ByVal tdp As TaskDependency)
'Compare current task dependency against full list of Driving Tasks and
'add-to/create/replace list of Path Drivers if critical

    Dim tdps As TaskDependencies 'store task dependencies
    Dim tdpI As TaskDependency 'store task dependency
    Dim tempFloat As Double 'tempFloat value used to compare float amongst all preds
    Dim i As Variant 'used to store unique ID of driving task if found in Driving Tasks array
    Dim predT As Task 'var to store pred task of evaluated dependency relationship
    Dim succT As Task 'var to store succ task of evaluated dependency relationship
    Dim predCritCoding As String 'var to store/modify existing Crit field values
    
    'Assign the dependency predecessor task to predT var
    Set predT = curproj.Tasks.UniqueID(tdp.From.UniqueID)
    
    'store predecessor task Crit path coding
    predCritCoding = predT.GetField(FieldNameToFieldConstant(CritField))
    
    'Assign the dependency successor task to the succT var
    Set succT = curproj.Tasks.UniqueID(tdp.To.UniqueID)
    
    'get the TrueFloat of Dependency relationship
    tempFloat = TrueFloat(predT, succT, tdp.Type, tdp.Lag, tdp.LagType)

    'If Evaluating Primary or Secondary Driving Tasks, and the TrueFloat value is not 0
    'Evaluate total network float and store in Driving Tasks array
    If tDrivingPaths.FindTertiary = False And tempFloat <> 0 Then
    
        'If other Driving Tasks have been found, Evaluate further
        If drivingTasksCount > 0 Then
            
            'Look for predecessor task in Driving Tasks Array
            i = FindInArray(predT.UniqueID)
    
            'If the task exists in the Driving Tasks array, evaluate further
            If Not IsNull(i) Then
            
                'if currently evaluating primary path, evaluate further
                If tDrivingPaths.FindSecondary = False Then
                
                    'if the dependency True Flaot is less than the previously stored float value
                    '(i.e. there are redundant links in the network), then store the lower float value
                    If tempFloat < DrivingTasks(i).tFloat Then
                        DrivingTasks(i).tFloat = tempFloat
                    End If
                Else 'if evaluating secondary path
                
                    'if the dependency float value + the sendary path float is less then the
                    'previously stored float vlaue, then store the lower float value
                    If tempFloat + tDrivingPaths.SecondaryFloat < DrivingTasks(i).tFloat Then
                        DrivingTasks(i).tFloat = tempFloat + tDrivingPaths.SecondaryFloat
                    End If
                End If
            Else 'If the task does not exist in the Driving Tasks array
            
                'Add new driver to the driving task count and store in the array
                drivingTasksCount = drivingTasksCount + 1
                ReDim Preserve DrivingTasks(1 To drivingTasksCount)
                DrivingTasks(drivingTasksCount).UID = predT.UniqueID
                
                'If evaluating the Primary Path, then store the float
                If tDrivingPaths.FindSecondary = False Then
                    DrivingTasks(drivingTasksCount).tFloat = tempFloat
                Else 'If evaluating secondary path, add float to the driving path network float value
                    DrivingTasks(drivingTasksCount).tFloat = tempFloat + tDrivingPaths.SecondaryFloat
                End If
            End If
        Else 'No other driving tasks found, this is the first driving task
            
            'Add the new driver to the driving tasks count and store in array
            drivingTasksCount = drivingTasksCount + 1
            ReDim DrivingTasks(1 To drivingTasksCount) 'removed Preserve - should not be neccessary when finding first driving task
            DrivingTasks(drivingTasksCount).UID = predT.UniqueID
            
            'If evaluating the Primary Path, then store the float
            If tDrivingPaths.FindSecondary = False Then
                DrivingTasks(drivingTasksCount).tFloat = tempFloat
            Else 'If evaluating secondary path, add float to the driving path network float value
                DrivingTasks(drivingTasksCount).tFloat = tempFloat + tDrivingPaths.SecondaryFloat
            End If
        End If
    End If
    
    'Evaluate new driver if True Float is 0
    If tempFloat = 0 Then
        
        'If other drivers exist, and evaluating Primary or Secondary path, evaluate further
        If drivingTasksCount > 0 And tDrivingPaths.FindTertiary = False Then
        
            'Look for predecessor task in Driving Tasks Array
            i = FindInArray(tdp.From.UniqueID)
    
            'If the task exists in the driving tasks array, update the float value
            If Not IsNull(i) Then
                DrivingTasks(i).tFloat = tempFloat
            Else 'If this is a new driver
            
                'Store the driving task in the Driving Tasks array
                drivingTasksCount = drivingTasksCount + 1
                ReDim Preserve DrivingTasks(1 To drivingTasksCount)
                With DrivingTasks(drivingTasksCount)
                    .UID = predT.UniqueID
                    .tFloat = tempFloat
                End With
            End If
            
        Else 'If no other driving tasks exists and not evaluating the tertiary path
            If tDrivingPaths.FindTertiary = False Then

                'Store the new driving task
                drivingTasksCount = drivingTasksCount + 1
                ReDim DrivingTasks(1 To drivingTasksCount) 'removed Preserve - should not be neccessary when finding first driving task
                With DrivingTasks(drivingTasksCount)
                    .UID = predT.UniqueID
                    .tFloat = tempFloat
                End With
            End If
        End If
    
        'If evaluating Primary Path, code the Crit and Group field values
        If tDrivingPaths.FindPrimary = True And tDrivingPaths.FindSecondary = False Then
            With predT
                .SetField FieldNameToFieldConstant(CritField), "1"
                .SetField FieldNameToFieldConstant(GroupField), "1"
            End With
        ElseIf tDrivingPaths.FindSecondary = True And tDrivingPaths.FindTertiary = False Then
            'If evaluating the secondary path, code the Crit and Group field values
            
            'If no existing code, then no need to concatenate
            If predCritCoding = vbNullString Then
                With predT
                    .SetField FieldNameToFieldConstant(CritField), "2"
                    .SetField FieldNameToFieldConstant(GroupField), "2"
                End With
            Else 'if existing code, then concatenate string
                If InStr(predCritCoding, "2") = 0 Then
                    predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & ",2"
                End If
            End If
            
        Else
            'If evaluating the tertiary path, code the Crit and Group field values
            
            'If no existing code, then no need to concatenate
            If predCritCoding = vbNullString Then
                With predT
                    .SetField FieldNameToFieldConstant(CritField), "3"
                    .SetField FieldNameToFieldConstant(GroupField), "3"
                End With
            Else 'if existing code, then concatenate string
                If InStr(predCritCoding, "3") = 0 Then
                    predT.SetField FieldNameToFieldConstant(CritField), predCritCoding & ",3"
                End If
            End If
        End If
    
        'store dependecies of the currently evaluted dependency
        Set tdps = predT.TaskDependencies
        
        'Iterate through the dependencies of the dependency
        For Each tdpI In tdps
        
            'evaluate task dependencies, add to analyzed tasks collection as needed, and review for criticality
            evaluateTaskDependencies tdpI, predT, curproj, AnalyzedTasks

        Next tdpI 'Next dependency of the currently evaluated dependency
    End If
        
End Sub

Private Function TrueFloat(ByVal tPred As Task, ByVal tSucc As Task, ByVal dType As Integer, ByVal dLag As Double, dlagtype As Integer) As Double
'Find True Float Value
'True Float is the dependency level 'free float' value,
'taking into consideration all duration types (including eDays),
'task calendars, leads/lags, etc

    Dim pDate As Date 'Store predecessor date (start or fin depending on link type)
    Dim sDate As Date 'Store successor date (start or fin depending on link type)
    Dim sCalObj As Calendar 'Store successor task calendar or project calendar if task cal = N/A
    Dim pCalObj As Calendar 'Store predecessor task calendar or project calendar if task cal = N/A
    Dim tempFloat As Double 'store True Float for function return
    
    'If pred task has a task calendar, store
    If tPred.Calendar <> "None" Then
        Set pCalObj = tPred.CalendarObject
    Else 'If no task calendar, store project cal
        Set pCalObj = ActiveProject.Calendar
    End If
    
    'If succ task has a task calendar, store
    If tSucc.Calendar <> "None" Then
        Set sCalObj = tSucc.CalendarObject
    Else 'If no task calendar, store project cal
        Set sCalObj = ActiveProject.Calendar
    End If

    'if dependency lag is greater than or equal to 0
    If dLag >= 0 Then
    
        'evaluate the depenency type
        Select Case dType
            
            Case 0 'Finish to Finish
                
                'Set predecessor date equal to the pred Finish date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Finish, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ finish
                sDate = tSucc.Finish
            
            Case 1 'Finish to Start
            
                'Set predecessor date equal to the pred Finish date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Finish, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ start
                sDate = tSucc.Start
            
            Case 2 'Start to Finish
            
                'Set predecessor date equal to the pred Start date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Start, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ finish
                sDate = tSucc.Finish
            
            Case 3 'Start to Start
            
                'Set predecessor date equal to the pred start date plus the lag value (considering any succ task cal)
                pDate = Application.DateAdd(tPred.Start, Application.DurationFormat(dLag, dlagtype), sCalObj)
                
                'successor date equals the succ start
                sDate = tSucc.Start
            Case Else
        End Select
    
    'if lag is less than 0 (lead)
    ElseIf dLag < 0 Then
    
        'evaluate the dependency type
        Select Case dType
            
            Case 0 'Finish to Finish
            
                'pred date equals the pred finish
                pDate = tPred.Finish
                
                'succ date equals the succ finish plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                sDate = Application.DateAdd(tSucc.Finish, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
            
            Case 1 'Finish to Start
            
                'pred date equals the pred finish
                pDate = tPred.Finish
                
                'succ date equals the succ start plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                sDate = Application.DateAdd(tSucc.Start, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
            
            Case 2 'Start to Finish
            
                'pred date equals the pred start
                pDate = tPred.Start
                
                'succ date equals the succ finish plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                sDate = Application.DateAdd(tSucc.Finish, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
            
            Case 3 'Start to Start
            
                'pred date equals the pred start
                pDate = tPred.Start
                
                'succ date equals the succ start plus the absolute value of the lag (considering any succ task cal)
                'adding the lead to the succ date is the same as subtracting it from the pred date (DateAdd cannot handle negative values)
                sDate = Application.DateAdd(tSucc.Start, Application.DurationFormat(Abs(dLag), dlagtype), sCalObj)
                
            Case Else
        End Select
    End If
    
    'subtract the pred date from the succ date, using the pred calendar, to get the True Float value
    tempFloat = Application.DateDifference(pDate, sDate, pCalObj)
    
    'Return the True Float value
    TrueFloat = tempFloat

End Function

Public Function ExistsInCollection(ByVal col As Collection, ByVal key As Variant) As Boolean
'Check for task dependency relationship in the analyzed tasks collection

    Dim f As Boolean 'stores boolean value 'True' if relationship exists in the collection
    
    'If error encountered, value does not exist in the collection
    On Error GoTo err
    
    f = IsObject(col.Item(key)) 'Store found item; if not found, will produce error
    ExistsInCollection = True 'Set True
    Exit Function
err: 'If error encountered, item does not exist - return "False" boolean vlaue
    ExistsInCollection = False
End Function
