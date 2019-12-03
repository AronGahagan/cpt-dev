Attribute VB_Name = "cptImportActuals_bas"
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
Dim strSchema As String
Dim strCon As String
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

  'user provides import file: indicate current or cumulative
  Set xlApp = CreateObject("Excel.Application")
  Set FileDialog = xlApp.FileDialog(msoFileDialogFilePicker)
  With FileDialog
    .AllowMultiSelect = False 'todo: yes or no?
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .Title = "Select Actuals import file:"
    '.Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
    .Filters.Add "Comma Separated Values (csv)", "*.csv"
    If .Show = -1 Then
      strFileName = FileDialog.SelectedItems(1)
    End If
  End With
  
  If Len(strFileName) = 0 Then GoTo exit_here
  
  'move file into temp directory
  Set FSO = CreateObject("Scripting.FileSystemObject")
  FSO.CopyFile strFileName, Environ("temp") & "\actuals.csv", True
  
  'setup schema.ini
  strSchema = Environ("temp") & "\Schema.ini"
  lngFile = FreeFile
  Open strSchema For Output As #lngFile
  Print #lngFile, "[actuals.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeader=True"
  Print #lngFile, "Col1=WPCN text"
  Print #lngFile, "Col2=RESOURCE text"
  Print #lngFile, "Col3=HOURS double"
  Print #lngFile, "Col4=DOLLARS double"
  Print #lngFile, "Col5=WEEK date"
  Close #lngFile
  
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("temp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  
  strSQL = "SELECT WPCN,RESOURCE,SUM(HOURS) AS [LABOR],SUM(DOLLARS) AS [MATL],WEEK "
  strSQL = strSQL & "FROM [actuals.csv] "
  strSQL = strSQL & "GROUP BY WPCN,RESOURCE,WEEK"
  
  Set rst = CreateObject("ADODB.Recordset")
  rst.Open strSQL, strCon, adOpenKeyset
  With rst
    If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
        strWPCN = rst("WPCN")
        Set Task = Nothing
        On Error Resume Next
        Set Task = ActiveProject.Tasks(strWPCN & " - ACTUALS")
        If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
        If Task Is Nothing Then
          Set Task = ActiveProject.Tasks.Add(strWPCN & " - ACTUALS")
        End If
        If Task.Type <> pjFixedDuration Then Task.Type = pjFixedDuration
        If Task.EffortDriven Then Task.EffortDriven = False
        If Task.Estimated Then Task.Estimated = False
      
        dblHours = rst("LABOR")
        dblMatl = rst("MATL")
        
        strResource = rst("RESOURCE")
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
        
        For Each Assignment In Task.Assignments
          If Assignment.ResourceName = strResource Then Exit For
        Next Assignment
        If Assignment Is Nothing Then
          Set Assignment = Task.Assignments.Add(Task.ID, Resource.ID, 1)
        End If
        
        If Assignment.RemainingWork > 0 Then Task.RemainingWork = 0
        
        dtWeek = CDate(rst("WEEK"))
        'make it a friday - todo: is this really necessary?
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
        
        'todo: flag it somehow
        'todo: optionally import into second file and use master to export reports?
        'If Task.Active Then Task.Active = False 'not possible on a task with actuals
        
next_record:
        .MoveNext
      Loop
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
