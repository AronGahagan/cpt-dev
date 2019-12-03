Attribute VB_Name = "cptImportActuals_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptImportActuals()
'objects
Dim FSO As Scripting.FileSystemObject 'Object
Dim rst As ADODB.Recordset
Dim FileDialog As FileDialog 'Object
Dim xlApp As Excel.Application
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim Assignment As Assignment
Dim Resource As Resource
Dim Task As Task
'strings
Dim strCon As String
Dim strSchema As String
Dim strDir As String
Dim strSQL As String
Dim strFileName As String
Dim strResource As String
Dim strWPCN As String
'longs
Dim lngFile As Long
'integers
'doubles
Dim dblMatl As Double
Dim dblHours As Double
'booleans
'variants
'dates
Dim dtWeek As Date
Dim dtStart As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'setup schema.ini
  Call cptCreateSchema("actuals.csv")
  
  'setup connection string
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("temp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  'user provides import file(s)
  Set xlApp = CreateObject("Excel.Application")
  Set FileDialog = xlApp.FileDialog(msoFileDialogFilePicker)
  With FileDialog
    .AllowMultiSelect = True
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .Title = "Select Actuals import file(s):"
    .Filters.Add "Comma Separated Values (csv)", "*.csv"
    If .Show = -1 Then
      For lngFile = 1 To FileDialog.SelectedItems.Count
        
        'move file into temp directory
        strFileName = FileDialog.SelectedItems(lngFile)
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.CopyFile strFileName, Environ("temp") & "\actuals.csv", True
        
        'query the file
        strSQL = "SELECT WPCN,RESOURCE,SUM(HOURS) AS [LABOR],SUM(DOLLARS) AS [MATL],WEEK "
        strSQL = strSQL & "FROM [actuals.csv] "
        strSQL = strSQL & "GROUP BY WPCN,RESOURCE,WEEK"
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open strSQL, strCon, adOpenKeyset
        With rst
          If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
            
              'find or create the task
              strWPCN = rst("WPCN")
              Set Task = Nothing
              On Error Resume Next
              Set Task = ActiveProject.Tasks(strWPCN & " - ACTUALS")
              If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
              If Task Is Nothing Then
                Set Task = ActiveProject.Tasks.Add(strWPCN & " - ACTUALS")
              End If
              'todo: capture/reset type, effortdriven
              If Task.Type <> pjFixedDuration Then Task.Type = pjFixedDuration
              If Task.EffortDriven Then Task.EffortDriven = False
              If Task.Estimated Then Task.Estimated = False
              If Task.RemainingWork > 0 Then Task.RemainingWork = 0
              
              'find or create the resource
              strResource = rst("RESOURCE")
              'grab hours and dollars > assumption is that matl resources
              'will have a cost but no labor, and vice versa
              dblHours = rst("LABOR")
              dblMatl = rst("MATL")
                            
              Set Resource = Nothing
              On Error Resume Next
              Set Resource = ActiveProject.Resources(strResource)
              If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
              If Resource Is Nothing Then
                If MsgBox("The resource '" & strResource & "' does not exist in this project. Add it?", vbExclamation + vbYesNo, "New Resource") = vbYes Then
                  Set Resource = ActiveProject.Resources.Add(strResource)
                  If dblHours = 0 And dblMatl > 0 Then
                    Resource.Type = pjResourceTypeMaterial
                    Resource.StandardRate = 1
                  End If
                Else
                  GoTo next_record
                End If
              End If
              
              'find or create the assignment
              For Each Assignment In Task.Assignments
                If Assignment.ResourceName = strResource Then Exit For
              Next Assignment
              If Assignment Is Nothing Then
                Set Assignment = Task.Assignments.Add(Task.ID, Resource.ID, 1)
              End If
              
              'reset any remaining work on the todo: task or assignment?
              If Assignment.RemainingWork > 0 Then Assignment.RemainingWork = 0
              
              'import the values to the proper week
              dtWeek = CDate(rst("WEEK"))
              If Weekday(dtWeek) <> 6 Then
                dtWeek = DateAdd("d", 6 - Weekday(dtWeek), dtWeek)
              End If
              If Resource.Type = pjResourceTypeWork Then
                Set TSVS = Assignment.TimeScaleData(dtWeek, dtWeek, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
                TSVS(1).Value = dblHours * 60
              Else
                Set TSVS = Assignment.TimeScaleData(dtWeek, dtWeek, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
                TSVS(1).Value = dblMatl
              End If
              
              'todo: flag it somehow?
              'todo: optionally import into second file and use master to export reports?
              'If Task.Active Then Task.Active = False 'not possible on a task with actuals
              
next_record:
              .MoveNext
            Loop
            'close the recordset
            .Close
          End If
        End With
        
      'next user-selected file
      Next lngFile
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set FSO = Nothing
  If Dir(Environ("temp") & "\actuals.csv") <> vbNullString Then Kill Environ("temp") & "\actuals.csv"
  If Dir(Environ("temp") & "\Schema.ini") <> vbNullString Then Kill Environ("temp") & "\Schema.ini"
  If rst.State Then rst.Close
  Set rst = Nothing
  Set FileDialog = Nothing
  Set xlApp = Nothing
  cptSpeed False
  Set TSV = Nothing
  Set TSVS = Nothing
  Set Assignment = Nothing
  Set Resource = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptImportActuals", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowImportActualsFrm()
'objects
Dim rst As ADODB.Recordset 'Object
Dim Task As Task
'strings
Dim strSettingsFile As String
Dim strFieldName As String
'longs
Dim lngItem As Long
Dim lngField As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  With cptImportActuals_frm
    
    'populate the WPCN pick list
    .cboWPCN.Clear
    For lngItem = 1 To 30
      lngField = FieldNameToFieldConstant("Text" & lngItem)
      If Len(CustomFieldGetName(lngField)) > 0 Then
        strFieldName = "Text" & lngItem & " (" & CustomFieldGetName(lngField) & ")"
      Else
        strFieldName = "Text" & lngItem
      End If
      .cboWPCN.AddItem
      .cboWPCN.List(lngItem - 1, 0) = strFieldName
    Next lngItem
    
    Set rst = CreateObject("ADODB.Recordset")
    rst.Fields.Append "UID", adBigInt
    rst.Fields.Append "TASK_NAME", adVarChar, 120
    rst.Open
    
    'populate the task pick list
    .cboTask.Clear
    lngItem = 0
    For Each Task In ActiveProject.Tasks
      If Task Is Nothing Then GoTo next_task
      If Task.Summary Then GoTo next_task 'todo: skip summaries?
      If Not Task.Active Then GoTo next_task
      If Task.ExternalTask Then GoTo next_task
      
      'todo: also save to adtg for quicker filtering
      
      .cboTask.AddItem
      .cboTask.List(lngItem, 0) = Task.UniqueID
      .cboTask.List(lngItem, 1) = Task.Name
      
      rst.AddNew Array(0, 1), Array(Task.UniqueID, Task.Name)
      
      lngItem = lngItem + 1
      
next_task:
    Next Task
    
    strSettingsFile = cptDir & "\settings\cpt-actuals-map.adtg"
    If Dir(strSettingsFile) <> vbNullString Then Kill strSettingsFile
    rst.Save strSettingsFile, adPersistADTG
    rst.Close
    
    'set default option
    .optNewTasks = True
    
    'todo: retrieve settings
    
    'present form to user
    .Show False
  End With
  
exit_here:
  On Error Resume Next
  Set rst = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptShowImportActualsFrm", Err, Erl)
  Resume exit_here
End Sub

Sub cptAddFilesActuals(ByRef Data As MSComctlLib.DataObject)
'objects
Dim FSO As Scripting.FileSystemObject 'object
'strings
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  For lngFile = 1 To Data.Files.Count
    With cptImportActuals_frm
      'todo: validate the file before adding it to the list
      .TreeView1.Nodes.Add Text:=Data.Files(lngFile)
      'copy files to temp directory
      FSO.CopyFile Data.Files(lngFile), Environ("temp") & "\" & Dir(Data.Files(lngFile)), True
    End With
  Next lngFile

exit_here:
  On Error Resume Next
  Set FSO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptAddFilesActuals", Err, Erl)
  Resume exit_here
End Sub

Sub cptListWPCN(ByRef Node As MSComctlLib.Node)
'objects
Dim rst As Object
'strings
Dim strCon As String
Dim strSQL As String
Dim strDir As String
Dim strFileName As String
'longs
Dim lngItem As Long
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  lngFile = FreeFile
  strFileName = Dir(Node.Text)
  
  'create schema.ini
  Call cptCreateSchema(strFileName)
  
  'setup connection string
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("temp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  'build query string
  strSQL = "SELECT DISTINCT WPCN,RESOURCE FROM [" & strFileName & "] "
  strSQL = strSQL & "ORDER BY WPCN,RESOURCE"
  
  'prep form for updated list of WPCN
  cptImportActuals_frm.lboMap.Clear
  
  'query the selected file for unique wpcns
  Set rst = CreateObject("ADODB.Recordset")
  lngItem = 0
  With rst
    .Open strSQL, strCon, adOpenKeyset
    Do While Not .EOF
      cptImportActuals_frm.lboMap.AddItem
      cptImportActuals_frm.lboMap.List(lngItem, 0) = rst(0)
      cptImportActuals_frm.lboMap.List(lngItem, 1) = rst(1)
      lngItem = lngItem + 1
      .MoveNext
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  If Dir(Environ("temp") & "\Schema.ini") <> vbNullString Then Kill Environ("temp") & "\Schema.ini"
  If rst.State Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptListWPCN", Err, Erl)
  Resume exit_here
End Sub

Sub cptCreateSchema(strFileName As String)
'objects
'strings
Dim strSchema As String
'longs
Dim lngFile As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'setup schema.ini
  strSchema = Environ("temp") & "\Schema.ini"
  lngFile = FreeFile
  Open strSchema For Output As #lngFile
  Print #lngFile, "[" & strFileName & "]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=WPCN text"
  Print #lngFile, "Col2=RESOURCE text"
  Print #lngFile, "Col3=HOURS double"
  Print #lngFile, "Col4=DOLLARS double"
  Print #lngFile, "Col5=WEEK date"
  Close #lngFile

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptCreateSchema", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateTaskMapList(Optional strText As String)
'objects
Dim rst As ADODB.Recordset 'Object
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set rst = CreateObject("ADODB.Recordset")

  With cptImportActuals_frm
    .cboTask.Clear
    rst.Open cptDir & "\settings\cpt-actuals-map.adtg"
    If Len(strText) > 0 Then
      rst.Filter = "TASK_NAME like '%" & strText & "%'"
    End If
    If rst.RecordCount > 0 Then
      rst.MoveFirst
      Do While Not rst.EOF
        .cboTask.AddItem
        .cboTask.List(.cboTask.ListCount - 1, 0) = rst(0)
        .cboTask.List(.cboTask.ListCount - 1, 1) = rst(1)
        rst.MoveNext
      Loop
    End If
    rst.Close
  End With

exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptImportActuals_bas", "cptUpdateTaskMapList", Err, Erl)
  Resume exit_here
End Sub
