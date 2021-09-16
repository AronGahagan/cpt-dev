Attribute VB_Name = "cptResourceDemand_bas"
'<cpt_version>v1.3.3</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptExportResourceDemand(Optional lngTaskCount As Long)
  'objects
  Dim oShell As Object
  Dim oSettings As Object
  Dim oListObject As Object
  Dim oSubproject As Object
  Dim oTask As Task
  Dim oResource As Resource
  Dim oAssignment As Assignment
  Dim tsv As TimeScaleValue
  Dim TSVS_BCWS As TimeScaleValues
  Dim TSVS_WORK As TimeScaleValues
  Dim TSVS_AW As TimeScaleValues
  Dim TSVS_COST As TimeScaleValues
  Dim TSVS_AC As TimeScaleValues
  Dim oCostRateTable As CostRateTable
  Dim oPayRate As PayRate
  Dim oExcel As Object
  Dim oWorksheet As Object
  Dim oWorkbook As Object
  Dim oRange As Object
  Dim oPivotTable As Object
  'dates
  Dim dtWeek As Date
  Dim dtStart As Date, dtFinish As Date, dtMin As Date, dtMax As Date
  'doubles
  Dim dblWork As Double, dblCost As Double
  'strings
  Dim strSettings As String
  Dim strTask As String
  Dim strMsg As String
  Dim strView As String
  Dim strFile As String, strRange As String
  Dim strTitle As String, strHeaders As String
  Dim strRecord As String, strFileName As String
  Dim strCost As String
  'longs
  Dim lngOffset As Long
  Dim lngRateSets As Long
  Dim lngCol As Long
  Dim lngOriginalRateSet As Long
  Dim lngFile As Long, lngTasks As Long, lngTask As Long
  Dim lngWeekCol As Long, lngExport As Long, lngField As Long
  Dim lngRateSet As Long
  Dim lngRow As Long
  'variants
  Dim vChk As Variant
  Dim vRateSet As Variant
  Dim aUserFields() As Variant
  'booleans
  Dim blnExportBaseline As Boolean
  Dim blnIncludeCosts As Boolean

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
    If ActiveProject.ProjectStart > dtStart Then dtStart = ActiveProject.ProjectStart
  Else
    MsgBox "Please enter a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    GoTo exit_here
  End If

  'save settings
  strFileName = Environ("USERPROFILE") & "\cpt-backup\settings\cpt-export-resource-userfields.adtg."
  aUserFields = cptResourceDemand_frm.lboExport.List()
  Set oSettings = CreateObject("ADODB.Recordset")
  With oSettings
    .Fields.Append "Field Constant", adVarChar, 255
    .Fields.Append "Custom Field Name", adVarChar, 255
    .Open
    strSettings = "Week=" & cptResourceDemand_frm.cboWeeks & ";"
    strSettings = strSettings & "Weekday=" & cptResourceDemand_frm.cboWeekday & ";"
    strSettings = strSettings & "Costs=" & cptResourceDemand_frm.chkCosts & ";"
    strSettings = strSettings & "Baseline=" & cptResourceDemand_frm.chkBaseline & ";"
    strSettings = strSettings & "RateSets="
    For Each vChk In Array("A", "B", "C", "D", "E")
      strSettings = strSettings & IIf(cptResourceDemand_frm.Controls("chk" & vChk), vChk & ",", "")
    Next vChk
    .AddNew Array(0, 1), Array("settings", strSettings)
    .Update
    'save userfields
    For lngExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
      .AddNew Array(0, 1), Array(aUserFields(lngExport, 0), aUserFields(lngExport, 1))
      .Update
    Next lngExport
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    .Save strFileName, adPersistADTG
    .Close
  End With
  
  lngFile = FreeFile
  Set oShell = CreateObject("WScript.Shell")
  strFile = oShell.SpecialFolders("Desktop") & "\" & Replace(Replace(ActiveProject.Name, ".mpp", ""), " ", "_") & "_ResourceDemand" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".csv"
  
  If Dir(strFile) <> vbNullString Then Kill strFile
  
  Open strFile For Output As #lngFile
  strHeaders = "PROJECT,[UID] TASK,RESOURCE_NAME,"
  '<issue42> get selected rate sets
  With cptResourceDemand_frm
    If .chkBaseline Then
      strHeaders = strHeaders & "BL_HOURS,BL_COST,"
    End If
    blnExportBaseline = .chkBaseline = True
    strHeaders = strHeaders & "HOURS,"
    'get rate sets
    blnIncludeCosts = .chkA Or .chkB Or .chkC Or .chkD Or .chkE
    If blnIncludeCosts Then strHeaders = strHeaders & "RATE_TABLE,COST,"
    If .chkA Then
      strHeaders = strHeaders & "COST_A,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkB Then
      strHeaders = strHeaders & "COST_B,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkC Then
      strHeaders = strHeaders & "COST_C,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkD Then
      strHeaders = strHeaders & "COST_D,"
      lngRateSets = lngRateSets + 1
    End If
    If .chkE Then
      strHeaders = strHeaders & "COST_E,"
      lngRateSets = lngRateSets + 1
    End If
    'get custom fields
    For lngExport = 0 To .lboExport.ListCount - 1
      lngField = .lboExport.List(lngExport, 0)
      If Len(CustomFieldGetName(lngField)) > 0 Then
        strHeaders = strHeaders & CustomFieldGetName(lngField) & ","
      Else
        strHeaders = strHeaders & FieldConstantToFieldName(lngField) & ","
      End If
    Next lngExport
    strHeaders = strHeaders & "WEEK"
  End With '</issue42>
  Print #lngFile, strHeaders

  If ActiveProject.Subprojects.Count = 0 Then
    lngTasks = ActiveProject.Tasks.Count
  Else
    cptSpeed True
    strView = ActiveWindow.TopPane.View.Name
    ViewApply "Gantt Chart"
    FilterClear
    GroupClear
    SelectAll
    OptionsViewEx displaysummarytasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        Sort "ID", , , , , , False, True
        OutlineShowAllTasks
      End If
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    SelectAll
    lngTasks = ActiveSelection.Tasks.Count
    ViewApply strView
    cptSpeed False
  End If

  'iterate over tasks
  Set oExcel = CreateObject("Excel.Application")
  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then 'skip blank lines
      If oTask.ExternalTask Then GoTo next_task 'skip external tasks
      If Not oTask.Summary And oTask.RemainingDuration > 0 And oTask.Active Then 'skip summary, complete tasks/milestones, and inactive
        
        'get earliest start and latest finish
        If cptResourceDemand_frm.chkBaseline Then
          dtStart = oExcel.WorksheetFunction.Min(oTask.Start, IIf(oTask.BaselineStart = "NA", oTask.Start, oTask.BaselineStart)) 'works with forecast, actual, and baseline start
          dtFinish = oExcel.WorksheetFunction.Max(oTask.Finish, IIf(oTask.BaselineFinish = "NA", oTask.Finish, oTask.BaselineFinish)) 'works with forecast, actual, and baseline finish
        Else
          If IsDate(oTask.Stop) Then 'capture the unstatused / remaining portion
            dtStart = oTask.Stop
          Else 'capture the entire unstarted task
            dtStart = oTask.Start
          End If
          dtFinish = oTask.Finish
        End If
        
        'capture oTask data common to all oAssignments
        strTask = oTask.Project & "," & Chr(34) & "[" & oTask.UniqueID & "] " & Replace(oTask.Name, Chr(34), Chr(39)) & Chr(34) & ","
        
        'examine every oAssignment on the task
        For Each oAssignment In oTask.Assignments
          
          If oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment
          
          'capture timephased work
          Set TSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
          For Each TSV In TSVS_WORK
            
            'capture common oAssignment data
            strRecord = strTask & oAssignment.ResourceName & ","
            
            'optionally capture baseline work and cost
            If cptResourceDemand_frm.chkBaseline Then
              Set TSVS_BCWS = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
              If oAssignment.ResourceType = pjResourceTypeWork Then
                strRecord = strRecord & Val(TSVS_BCWS(1).Value) / 60 & ","
              Else
                strRecord = strRecord & "0,"
              End If
              Set TSVS_BCWS = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks, 1)
              strRecord = strRecord & Val(TSVS_BCWS(1).Value) & ","
            End If
            'capture (and subtract) actual work, leaving ETC/Remaining Work
            Set TSVS_AW = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleWeeks, 1)
            dblWork = Val(TSV.Value) - Val(TSVS_AW(1))
            If oAssignment.ResourceType = pjResourceTypeWork Then
              strRecord = strRecord & dblWork / 60 & ","
            Else
              strRecord = strRecord & "0,"
            End If
            'get costs
            If blnIncludeCosts Then
              'rate set
              strRecord = strRecord & Choose(oAssignment.CostRateTable + 1, "A", "B", "C", "D", "E") & ","
              Set TSVS_COST = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
              'get actual cost
              Set TSVS_AC = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
              'subtract actual cost from cost to get remaining cost
              dblCost = Val(TSVS_COST(1).Value) - Val(TSVS_AC(1))
              'get cost
              If dblWork > 0 Or dblCost > 0 Then 'there is remaining work or cost
                strRecord = strRecord & dblCost & ","
              Else
                strRecord = strRecord & "0,"
              End If
            End If
            
            'if default rate set is included then include it
            If cptResourceDemand_frm.chkA Then
              strRecord = strRecord & IIf(oAssignment.CostRateTable = 0, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkB Then
              strRecord = strRecord & IIf(oAssignment.CostRateTable = 1, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkC Then
              strRecord = strRecord & IIf(oAssignment.CostRateTable = 2, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkD Then
              strRecord = strRecord & IIf(oAssignment.CostRateTable = 3, dblCost, 0) & ","
            End If
            If cptResourceDemand_frm.chkE Then
              strRecord = strRecord & IIf(oAssignment.CostRateTable = 4, dblCost, 0) & ","
            End If
            
            'get custom field values
            For lngExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
              lngField = cptResourceDemand_frm.lboExport.List(lngExport, 0)
              strRecord = strRecord & Chr(34) & Trim(Replace(oTask.GetField(lngField), ",", "-")) & Chr(34) & ","
            Next lngExport
            
            'apply user settings for week identification
            With cptResourceDemand_frm
              If .cboWeeks = "Beginning" Then
                dtWeek = TSV.StartDate
                If .cboWeekday = "Monday" Then
                  dtWeek = DateAdd("d", 1, dtWeek)
                End If
              ElseIf .cboWeeks = "Ending" Then
                dtWeek = TSV.EndDate
                If .cboWeekday = "Friday" Then
                  dtWeek = DateAdd("d", -2, dtWeek)
                ElseIf .cboWeekday = "Saturday" Then
                  dtWeek = DateAdd("d", -1, dtWeek)
                End If
              End If
            End With
            strRecord = strRecord & Format(dtWeek, "mm/dd/yyyy") & "," 'week
            Print #lngFile, strRecord
          Next TSV
          
          'get rate set and cost
          lngOriginalRateSet = oAssignment.CostRateTable
          'todo: only include baseline cost if both baseline and costs are checked
          If cptResourceDemand_frm.chkBaseline Then strRecord = strRecord & "0,0," 'BL HOURS, BL COST
          For lngRateSet = 0 To 4
            'need msproj to calculate the cost
            If cptResourceDemand_frm.Controls(Choose(lngRateSet + 1, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
              If lngRateSet = lngOriginalRateSet Then GoTo next_rate_set
              Application.StatusBar = "Exporting Rate Set " & Replace(Choose(lngRateSet + 1, "chkA", "chkB", "chkC", "chkD", "chkE"), "chk", "") & "..."
              If oAssignment.CostRateTable <> lngRateSet Then oAssignment.CostRateTable = lngRateSet 'recalculation not needed
              'extract timephased date
              'get work
              Set TSVS_WORK = oAssignment.TimeScaleData(dtStart, dtFinish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
              For Each TSV In TSVS_WORK
                strRecord = oTask.Project & "," & Chr(34) & "[" & oTask.UniqueID & "] " & Replace(oTask.Name, Chr(34), Chr(39)) & Chr(34) & ","
                strRecord = strRecord & oAssignment.ResourceName & ","
                If cptResourceDemand_frm.chkBaseline Then strRecord = strRecord & "0,0," 'baseline placeholder
                strRecord = strRecord & "0," 'hours
                strRecord = strRecord & Choose(lngOriginalRateSet + 1, "A", "B", "C", "D", "E") & ","
                strRecord = strRecord & "0," 'cost
                'get cost
                Set TSVS_COST = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledCost, pjTimescaleWeeks, 1)
                'get actual cost
                Set TSVS_AC = oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleWeeks, 1)
                'subtract actual cost from cost to get remaining cost
                dblCost = Val(TSVS_COST(1).Value) - Val(TSVS_AC(1))
                'hacky way of figuring out how many zeroes to include
                'and how to replace the right one with the dblCost
                With cptResourceDemand_frm
                  If .chkA Then strCost = "[0],"
                  If .chkB Then strCost = strCost & "[1],"
                  If .chkC Then strCost = strCost & "[2],"
                  If .chkD Then strCost = strCost & "[3],"
                  If .chkE Then strCost = strCost & "[4],"
                End With
                If dblCost > 0 Then
                  strCost = Replace(strCost, "[" & lngRateSet & "]", dblCost)
                  strCost = Replace(strCost, "[0]", "0")
                  strCost = Replace(strCost, "[1]", "0")
                  strCost = Replace(strCost, "[2]", "0")
                  strCost = Replace(strCost, "[3]", "0")
                  strCost = Replace(strCost, "[4]", "0")
                  strRecord = strRecord & strCost
                Else
                  strRecord = strRecord & Replace(String(lngRateSets, "0"), "0", "0,")
                End If
                
                'get custom field values
                For lngExport = 0 To cptResourceDemand_frm.lboExport.ListCount - 1
                  lngField = cptResourceDemand_frm.lboExport.List(lngExport, 0)
                  strRecord = strRecord & oTask.GetField(lngField) & ","
                Next lngExport
                
                'apply user settings for week identification
                With cptResourceDemand_frm
                  If .cboWeeks = "Beginning" Then
                    dtWeek = TSV.StartDate
                    If .cboWeekday = "Monday" Then
                      dtWeek = DateAdd("d", 1, dtWeek)
                    End If
                  ElseIf .cboWeeks = "Ending" Then
                    dtWeek = TSV.EndDate
                    If .cboWeekday = "Friday" Then
                      dtWeek = DateAdd("d", -2, dtWeek)
                    ElseIf .cboWeekday = "Saturday" Then
                      dtWeek = DateAdd("d", -1, dtWeek)
                    End If
                  End If
                End With
                strRecord = strRecord & Format(dtWeek, "mm/dd/yyyy") & "," 'week
                Print #lngFile, strRecord
              Next TSV
            End If
next_rate_set:
          Next lngRateSet
          If oAssignment.CostRateTable <> lngOriginalRateSet Then oAssignment.CostRateTable = lngOriginalRateSet

next_assignment:
        Next oAssignment
      End If 'skip external tasks
    End If 'skip blank lines
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Exporting " & Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & "...(" & Format(lngTask / lngTasks, "0%") & ")"
    cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
    cptResourceDemand_frm.lblProgress.Width = (lngTask / lngTasks) * cptResourceDemand_frm.lblStatus.Width
    DoEvents
  Next oTask

  'close the CSV
  Close #lngFile

  Application.StatusBar = "Creating Workbook..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  'set reference to Excel
  '<issue14-15> added
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If

  'is previous run still open?
  On Error Resume Next
  Set oWorkbook = oExcel.oWorkbooks(strFile)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oWorkbook Is Nothing Then oWorkbook.Close False
  On Error Resume Next
  Set oWorkbook = oExcel.Workbooks(Environ("TEMP") & "\ExportResourceDemand.xlsx")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oWorkbook Is Nothing Then 'add timestamp to existing file
    If oWorkbook.Application.Visible = False Then oWorkbook.Application.Visible = True
    strMsg = "'" & strFile & "' already exists and is open."
    strFile = Replace(strFile, ".xlsx", "_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx")
    strMsg = strMsg & "Your new file will be saved as:" & vbCrLf & strFile
    MsgBox strMsg, vbExclamation + vbOKOnly, "File Exists and is Open"
  End If '</issue14-15>

  'create a new Workbook
  Set oWorkbook = oExcel.Workbooks.Open(strFile)

  On Error Resume Next
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then Kill Environ("TEMP") & "\ExportResourceDemand.xlsx"
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Dir(Environ("TEMP") & "\ExportResourceDemand.xlsx") <> vbNullString Then 'kill failed, rename it
    oWorkbook.SaveAs Environ("TEMP") & "\ExportResourceDemand_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".xlsx", 51
  Else
    oWorkbook.SaveAs Environ("TEMP") & "\ExportResourceDemand.xlsx", 51
  End If
  If Dir(strFile) <> vbNullString Then Kill strFile '</issue14-15>

  'set reference to oWorksheet to manipulate it
  Set oWorksheet = oWorkbook.Sheets(1)
  'rename the oWorksheet
  oWorksheet.Name = "SourceData"
  lngWeekCol = oWorksheet.Rows(1).Find(what:="WEEK").Column
  dtMin = oExcel.WorksheetFunction.Min(oWorksheet.Columns(lngWeekCol))
  dtMax = oExcel.WorksheetFunction.Max(oWorksheet.Columns(lngWeekCol))
  
  Set oWorksheet = oWorkbook.Sheets("SourceData")
  
  'format currencies
  For lngCol = 1 To lngWeekCol
    If InStr(oWorksheet.Cells(1, lngCol), "COST") > 0 Then oWorksheet.Columns(lngCol).Style = "Currency"
  Next lngCol
  
  'add note on CostRateTable column
  If blnIncludeCosts Then
    lngCol = oWorksheet.Rows(1).Find("RATE_TABLE", lookat:=1).Column
    oWorksheet.Cells(1, lngCol).AddComment "Rate Table Applied in the Project"
  End If
  
  'todo: export exceptions
  'todo: setup HPM Table
  
  'create FTE_WEEK column
  Set oRange = oWorksheet.[A1].End(xlToRight).End(xlDown).Offset(0, 1)
  Set oRange = oWorksheet.Range(oRange, oWorksheet.[A1].End(xlToRight).Offset(1, 1))
  lngCol = oWorksheet.Rows(1).Find("HOURS", lookat:=1).Column '1=xlWhole
  oRange.FormulaR1C1 = "=RC" & lngCol & "/40"
  oWorksheet.[A1].End(xlToRight).Offset(0, 1).Value = "FTE_WEEK"
  
  'create FTE_MONTH column
  Set oRange = oWorksheet.[A1].End(xlToRight).End(xlDown).Offset(0, 1)
  Set oRange = oWorksheet.Range(oRange, oWorksheet.[A1].End(xlToRight).Offset(1, 1))
  lngCol = oWorksheet.Rows(1).Find("HOURS", lookat:=1).Column '1=xlWhole
  'todo: allow for Holidays, Fiscal Periods
  oRange.FormulaR1C1 = "=RC" & lngCol & "/160"
  oWorksheet.[A1].End(xlToRight).Offset(0, 1).Value = "FTE_MONTH"
  
  If blnExportBaseline Then
    'include FTE_BL_WEEK
    Set oRange = oWorksheet.[A1].End(xlToRight).End(xlDown).Offset(0, 1)
    Set oRange = oWorksheet.Range(oRange, oWorksheet.[A1].End(xlToRight).Offset(1, 1))
    lngCol = oWorksheet.Rows(1).Find("BL_HOURS", lookat:=1).Column '1=xlWhole
    oRange.FormulaR1C1 = "=RC" & lngCol & "/40"
    oWorksheet.[A1].End(xlToRight).Offset(0, 1).Value = "FTE_BL_WEEK"
    
    'include FTE_BL_MONTH
    Set oRange = oWorksheet.[A1].End(xlToRight).End(xlDown).Offset(0, 1)
    Set oRange = oWorksheet.Range(oRange, oWorksheet.[A1].End(xlToRight).Offset(1, 1))
    lngCol = oWorksheet.Rows(1).Find("BL_HOURS", lookat:=1).Column '1=xlWhole
    oRange.FormulaR1C1 = "=RC" & lngCol & "/160"
    oWorksheet.[A1].End(xlToRight).Offset(0, 1).Value = "FTE_BL_MONTH"
  End If
  
  'capture the range of data to feed as variable to PivotTable
  Set oRange = oWorksheet.Range(oWorksheet.[A1].End(xlDown), oWorksheet.[A1].End(xlToRight))
  strRange = oWorksheet.Name & "!" & Replace(oRange.Address, "$", "")
  'add a new Worksheet for the oPivotTable
  Set oWorksheet = oWorkbook.Sheets.Add
  'rename the new Worksheet
  oWorksheet.Name = "ResourceDemand"

  Application.StatusBar = "Creating PivotTable..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  'create the PivotTable
  oWorkbook.PivotCaches.Create(SourceType:=1, _
        SourceData:=strRange, Version:= _
        3).CreatePivotTable TableDestination:="ResourceDemand!R3C1", TableName:="RESOURCE_DEMAND", DefaultVersion:=3
  Set oPivotTable = oWorksheet.PivotTables(1)
  oPivotTable.AddFields Array("RESOURCE_NAME", "PROJECT", "[UID] TASK"), Array("WEEK") 'Array("FISCAL_YEAR", "FISCAL_MONTH", "WEEK")
  oPivotTable.AddDataField oPivotTable.PivotFields("FTE_WEEK"), "FTE_WEEK ", -4157
  'format the oPivotTable
  oPivotTable.PivotFields("RESOURCE_NAME").ShowDetail = False
  oPivotTable.TableStyle2 = "PivotStyleMedium2"
  oPivotTable.PivotSelect "", 2, True
  oExcel.Selection.Style = "Comma"

  Application.StatusBar = "Building header..."
  cptResourceDemand_frm.lblStatus = Application.StatusBar

  'add a title
  oWorksheet.[A2] = "Status Date: " & FormatDateTime(ActiveProject.StatusDate, vbShortDate)
  oWorksheet.[A2].EntireColumn.AutoFit
  oWorksheet.[A1] = "REMAINING WORK IN IMS: " & Replace(ActiveProject.Name, " ", "_")
  oWorksheet.[A1].Font.Bold = True
  oWorksheet.[A1].Font.Italic = True
  oWorksheet.[A1].Font.Size = 14
  oWorksheet.[A1:F1].Merge
  'revise according to user options
  oWorksheet.[B2] = "Weeks " & cptResourceDemand_frm.cboWeeks.Value & " " & cptResourceDemand_frm.cboWeekday.Value
  oWorksheet.[B4].Select
  oWorksheet.[B5].Select

  'make it nice
  oExcel.ActiveWindow.Zoom = 85
  'todo: add 'heatmap' like coloring
'    Cells.FormatConditions.Delete
'    Range("B5:HT25").Select
'    Range("B8").Activate
'    Selection.FormatConditions.AddColorScale ColorScaleType:=2
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
'    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
'        .Color = 10285055
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueHighestValue
'    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
'        .Color = 2650623
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).ScopeType = xlFieldsScope
'    ActiveSheet.PivotTables("RESOURCE_DEMAND").TableStyle2 = "PivotStyleLight22"

  Application.StatusBar = "Creating PivotChart..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  'create a PivotChart
  Set oWorksheet = oWorkbook.Sheets("SourceData")
  oWorksheet.Activate
  oWorksheet.[A2].Select
  oWorksheet.[A2].EntireColumn.AutoFit
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Cells.EntireColumn.AutoFit
  Set oWorksheet = oWorkbook.Sheets.Add
  oWorksheet.Name = "PivotChart_Source"
  oWorkbook.Worksheets("ResourceDemand").PivotTables("RESOURCE_DEMAND"). _
        PivotCache.CreatePivotTable TableDestination:="PivotChart_Source!R1C1", TableName:= _
        "PivotTable1", DefaultVersion:=3
  Set oWorksheet = oWorkbook.Sheets("PivotChart_Source")
  oWorksheet.[A1].Select
  oExcel.ActiveSheet.Shapes.AddChart.Select
  Set oRange = oWorksheet.Range(oWorksheet.[A1].End(-4161), oWorksheet.[A1].End(-4121))
  oExcel.ActiveChart.SetSourceData Source:=oRange
  oWorkbook.ShowPivotChartActiveFields = True
  oExcel.ActiveChart.ChartType = 76 'xlAreaStacked
  With oExcel.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK")
    .Orientation = 1 'xlRowField
    .Position = 1
  End With
  oExcel.ActiveChart.PivotLayout.PivotTable.AddDataField oExcel.ActiveChart.PivotLayout. _
        PivotTable.PivotFields("HOURS"), "Sum of HOURS", -4157
  With oExcel.ActiveChart.PivotLayout.PivotTable.PivotFields("RESOURCE_NAME")
    .Orientation = 2 'xlColumnField
    .Position = 1
  End With
  With oExcel.ActiveChart.PivotLayout.PivotTable.PivotFields("WEEK")
    .Orientation = 1 'xlRowField
    .Position = 1
  End With
  If Not cptResourceDemand_frm.chkBaseline Then oExcel.ActiveSheet.PivotTables("PivotTable1").PivotFields("WEEK").PivotFilters.Add _
        Type:=33, Value1:=ActiveProject.StatusDate '33 = xlAfter
  oExcel.ActiveChart.ClearToMatchStyle
  oExcel.ActiveChart.ChartStyle = 34
  oExcel.ActiveChart.ClearToMatchStyle
  oExcel.ActiveSheet.ChartObjects(1).Activate
  oExcel.ActiveChart.SetElement (msoElementChartTitleAboveChart)
  oExcel.ActiveChart.ChartTitle.Text = "Resource Demand"
  oExcel.ActiveChart.Location 1, "PivotChart" 'xlLocationAsNewSheet = 1
  Set oWorksheet = oWorkbook.Sheets("PivotChart_Source")
  oWorksheet.Visible = False

  'add legend
  oExcel.ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
  oExcel.ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "HOURS"
  
  'export selected cost rate tables to oWorksheet
  If blnIncludeCosts Then
    Application.StatusBar = "Exporting Cost Rate Tables..."
    cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
    Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets("SourceData"))
    oWorksheet.Name = "Cost Rate Tables"
    oWorksheet.[A1:I1].Value = Array("PROJECT", "RESOURCE_NAME", "RESOURCE_TYPE", "ENTERPRISE", "RATE_TABLE", "EFFECTIVE_DATE", "STANDARD_RATE", "OVERTIME_RATE", "PER_USE_COST")
    lngRow = 2
    'make compatible with master/sub projects
    If ActiveProject.ResourceCount > 0 Then
      For Each oResource In ActiveProject.Resources
        oWorksheet.Cells(lngRow, 1) = oResource.Name
        For Each oCostRateTable In oResource.CostRateTables
          If cptResourceDemand_frm.Controls(Choose(oCostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
            For Each oPayRate In oCostRateTable.PayRates
              oWorksheet.Cells(lngRow, 1) = ActiveProject.Name
              oWorksheet.Cells(lngRow, 2) = oResource.Name
              oWorksheet.Cells(lngRow, 3) = Choose(oResource.Type + 1, "Work", "Material", "Cost")
              oWorksheet.Cells(lngRow, 4) = oResource.Enterprise
              oWorksheet.Cells(lngRow, 5) = oCostRateTable.Name
              oWorksheet.Cells(lngRow, 6) = Format(oPayRate.EffectiveDate, "mm/dd/yyyy")
              oWorksheet.Cells(lngRow, 7) = Replace(oPayRate.StandardRate, "/h", "")
              oWorksheet.Cells(lngRow, 8) = Replace(oPayRate.OvertimeRate, "/h", "")
              oWorksheet.Cells(lngRow, 9) = oPayRate.CostPerUse
              lngRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162).Row + 1
            Next oPayRate
          End If
        Next oCostRateTable
      Next oResource
    ElseIf ActiveProject.Subprojects.Count > 0 Then
      For Each oSubproject In ActiveProject.Subprojects
        For Each oResource In oSubproject.SourceProject.Resources
          oWorksheet.Cells(lngRow, 1) = oResource.Name
          For Each oCostRateTable In oResource.CostRateTables
            If cptResourceDemand_frm.Controls(Choose(oCostRateTable.Index, "chkA", "chkB", "chkC", "chkD", "chkE")).Value = True Then
              For Each oPayRate In oCostRateTable.PayRates
                oWorksheet.Cells(lngRow, 1) = oSubproject.SourceProject.Name
                oWorksheet.Cells(lngRow, 2) = oResource.Name
                oWorksheet.Cells(lngRow, 3) = Choose(oResource.Type + 1, "Work", "Material", "Cost")
                oWorksheet.Cells(lngRow, 4) = oResource.Enterprise
                oWorksheet.Cells(lngRow, 5) = oCostRateTable.Name
                oWorksheet.Cells(lngRow, 6) = Format(oPayRate.EffectiveDate, "mm/dd/yyyy")
                oWorksheet.Cells(lngRow, 7) = Replace(oPayRate.StandardRate, "/h", "")
                oWorksheet.Cells(lngRow, 8) = Replace(oPayRate.OvertimeRate, "/h", "")
                oWorksheet.Cells(lngRow, 9) = oPayRate.CostPerUse
                lngRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162).Row + 1
              Next oPayRate
            End If
          Next oCostRateTable
        Next oResource
      Next oSubproject
    End If
  
    'make it a oListObject
    Set oListObject = oWorksheet.oListObjects.Add(1, oWorksheet.Range(oWorksheet.[A1].End(-4161), oWorksheet.[A1].End(-4121)).Address, , 1)
    oListObject.Name = "CostRateTables"
    oListObject.TableStyle = ""
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.[A2].Select
    oExcel.ActiveWindow.FreezePanes = True
    oWorksheet.Columns.AutoFit
    
  End If
    
  'PivotTable Worksheet active by default
  oWorkbook.Sheets("ResourceDemand").Activate
  
  'todo: add conditional formatting?
  
  'provide user feedback
  Application.StatusBar = "Saving the Workbook..."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar
  
  'save the file
  '<issue49> - file exists in location
  strFile = oShell.SpecialFolders("Desktop") & "\" & Replace(oWorkbook.Name, ".xlsx", "_" & Format(Now(), "yyyy-mm-dd-hh-nn-ss") & ".xlsx") '<issue49>
  If Dir(strFile) <> vbNullString Then '<issue49>
    If MsgBox("A file named '" & strFile & "' already exists in this location. Replace?", vbYesNo + vbExclamation, "Overwrite?") = vbYes Then '<issue49>
      Kill strFile '<issue49>
      oWorkbook.SaveAs strFile, 51 '<issue49>
      MsgBox "Saved to your Desktop:" & vbCrLf & vbCrLf & Dir(strFile), vbInformation + vbOKOnly, "Resource Demand Exported" '<issue49>
    End If '<issue49>
  Else '<issue49>
    oWorkbook.SaveAs strFile, 51  '<issue49>
  End If '</issue49>
  
  MsgBox "Export Complete", vbInformation + vbOKOnly, "FTE"
  
  Application.StatusBar = "Complete."
  cptResourceDemand_frm.lblStatus.Caption = Application.StatusBar

  oExcel.Visible = True
  Application.ActivateMicrosoftApp pjMicrosoftExcel
  
exit_here:
  On Error Resume Next
  Set oShell = Nothing
  Set oSettings = Nothing
  Set oListObject = Nothing
  Set oSubproject = Nothing
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Application.StatusBar = ""
  cptResourceDemand_frm.lblStatus.Caption = "Ready..."
  For lngFile = 1 To FreeFile
    Close #lngFile
  Next lngFile
  cptSpeed False
  Set oTask = Nothing
  Set oResource = Nothing
  Set oAssignment = Nothing
  Set oCostRateTable = Nothing
  Set oPayRate = Nothing
  Set oExcel = Nothing
  Set oPivotTable = Nothing
  Set oListObject = Nothing
  Set oWorkbook = Nothing
  Set oWorksheet = Nothing
  Set tsv = Nothing
  Set TSVS_BCWS = Nothing
  Set TSVS_WORK = Nothing
  Set TSVS_AW = Nothing
  Set TSVS_COST = Nothing
  Set TSVS_AC = Nothing
  Set oRange = Nothing

  If Not oWorkbook Is Nothing Then oWorkbook.Close False
  If Not oExcel Is Nothing Then oExcel.Quit
  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_bas", "cptExportResourceDemand", Err, Erl)
  On Error Resume Next
  Resume exit_here

End Sub

Sub cptShowExportResourceDemand_frm()
'objects
Dim rst As Object 'ADODB.Recordset
Dim rstResources As Object 'ADODB.Recordset
Dim objProject As Object
Dim rstFields As Object 'ADODB.Recordset
'strings
Dim strMissing As String
Dim strActiveView As String
Dim strFieldName As String, strFileName As String
'longs
Dim lngResourceCount As Long, lngResource As Long
Dim lngField As Long, lngItem As Long
'integers
'booleans
'variants
Dim vCostSet As Variant
Dim vCostSets As Variant
Dim vFieldType As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'requires ms excel
  If Not cptCheckReference("Excel") Then
    MsgBox "This feature requires MS Excel.", vbCritical + vbOKOnly, "Resource Demand"
    GoTo exit_here
  End If
  'todo: why do this?
  If ActiveProject.Subprojects.Count = 0 And ActiveProject.ResourceCount = 0 Then
    MsgBox "This project has no resources to export.", vbExclamation + vbOKOnly, "No Resources"
    GoTo exit_here
  Else
    cptSpeed True
    lngResourceCount = ActiveProject.ResourceCount
    Set rstResources = CreateObject("ADODB.Recordset")
    rstResources.Fields.Append "RESOURCE_NAME", adVarChar, 200
    rstResources.Open
    For lngItem = 1 To ActiveProject.Subprojects.Count
      Set objProject = ActiveProject.Subprojects(lngItem).SourceProject
      Application.StatusBar = "Loading " & objProject.Name & "..."
      For lngResource = 1 To objProject.Resources.Count
        With rstResources
          .Filter = "[RESOURCE_NAME]='" & objProject.Resources(lngResource).Name & "'"
          If rstResources.RecordCount = 0 Then
            .AddNew Array(0), Array("'" & objProject.Resources(lngResource).Name & "'")
          Else
            Debug.Print "duplicate found"
          End If
          .Filter = ""
        End With
      Next lngResource
      Set objProject = Nothing
    Next lngItem
    rstResources.Close 'todo: save for later?
    Application.StatusBar = ""
    cptSpeed False
  End If

  cptResourceDemand_frm.lboFields.Clear
  cptResourceDemand_frm.lboExport.Clear

  Set rstFields = CreateObject("ADODB.Recordset")
  rstFields.Fields.Append "CONSTANT", adInteger
  rstFields.Fields.Append "CUSTOM_NAME", adVarChar, 200
  rstFields.Open
  
  'add the 'Critical' field
  rstFields.AddNew Array(0, 1), Array(FieldNameToFieldConstant("Critical"), "Critical")
  
  'todo: add the TrueFloat Fields - get from ini?
  
  For Each vFieldType In Array("Text", "Outline Code")
    On Error GoTo err_here
    For lngItem = 1 To 30
      lngField = FieldNameToFieldConstant(vFieldType & lngItem) ',lngFieldType)
      strFieldName = CustomFieldGetName(lngField)
      If Len(strFieldName) > 0 Then
        'todo: handle duplicates if master/subprojects
        rstFields.AddNew Array(0, 1), Array(lngField, strFieldName)
        rstFields.Update
      End If
next_field:
    Next lngItem
  Next vFieldType

  'get enterprise custom fields
  For lngField = 188776000 To 188778000
    If FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      strFieldName = Application.FieldConstantToFieldName(lngField)
      'todo: avoid conflicts between local and custom fields?
      'If rstFields.Contains(strFieldName) Then
      '  MsgBox "An Enterprise Field named '" & strFieldName & "' conflicts with a local custom field of the same name. The local field will be ignored.", vbExclamation + vbOKOnly, "Conflict"
        'rstFields.Remove Application.FieldConstantToFieldName(lngField)
      'End If
      rstFields.AddNew Array(0, 1), Array(lngField, strFieldName)
      rstFields.Update
    End If
next_field1:
  Next lngField

  'add fields to listbox
  rstFields.Sort = "CUSTOM_NAME"
  rstFields.MoveFirst
  lngItem = 0
  Do While Not rstFields.EOF
    cptResourceDemand_frm.lboFields.AddItem
    cptResourceDemand_frm.lboFields.List(lngItem, 0) = rstFields(0)
    If rstFields(0) > 188776000 Then
      cptResourceDemand_frm.lboFields.List(lngItem, 1) = rstFields(1) & " (Enterprise)"
    Else
      cptResourceDemand_frm.lboFields.List(lngItem, 1) = rstFields(1) & " (" & FieldConstantToFieldName(rstFields(0)) & ")"
    End If
    rstFields.MoveNext
    lngItem = lngItem + 1
  Loop

  'save the fields to a file for fast searching
  If rstFields.RecordCount > 0 Then
    strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
    If Dir(strFileName) <> vbNullString Then Kill strFileName
    rstFields.Save strFileName, adPersistADTG
  End If
  rstFields.Close
  
  'populate options and set defaults
  With cptResourceDemand_frm
    .cboWeeks.AddItem "Beginning"
    .cboWeeks.AddItem "Ending"
    'allow to trigger, it populates the form
    .cboWeeks.Value = "Beginning"
    .cboWeekday = "Monday"
    .chkA.Value = False
    .chkB.Value = False
    .chkC.Value = False
    .chkD.Value = False
    .chkE.Value = False
    .chkCosts.Value = False
    .chkBaseline = False
  End With
  
  'import saved fields if exists
  strFileName = Environ("USERPROFILE") & "\cpt-backup\settings\cpt-export-resource-userfields.adtg"
  If Dir(strFileName) <> vbNullString Then
    Set rst = CreateObject("ADODB.Recordset")
    With rst
      .Open strFileName, , adOpenKeyset, adLockReadOnly
      .MoveFirst
      lngItem = 0
      Do While Not .EOF
        If .Fields(0) = "settings" Then
          cptResourceDemand_frm.cboWeeks.Value = Replace(Replace(cptRegEx(.Fields(1), "Week\=[A-z]*;"), "Week=", ""), ";", "")
          cptResourceDemand_frm.cboWeekday = Replace(Replace(cptRegEx(.Fields(1), "Weekday\=[A-z]*;"), "Weekday=", ""), ";", "")
          cptResourceDemand_frm.chkCosts = Replace(Replace(cptRegEx(.Fields(1), "Costs\=[A-z]*;"), "Costs=", ""), ";", "")
          cptResourceDemand_frm.chkBaseline = Replace(Replace(cptRegEx(.Fields(1), "Baseline\=[A-z]*;"), "Baseline=", ""), ";", "")
          vCostSets = Split(Replace(cptRegEx(.Fields(1), "RateSets\=[A-z\,]*"), "RateSets=", ""), ",")
          If cptResourceDemand_frm.chkCosts Then
            For vCostSet = 0 To UBound(vCostSets) - 1
              cptResourceDemand_frm.Controls("chk" & vCostSets(vCostSet)).Value = True
            Next vCostSet
          Else
            'todo: convert to bitwise?
            For Each vCostSet In Array("A", "B", "C", "D", "E")
              cptResourceDemand_frm.Controls("chk" & vCostSet) = False
              cptResourceDemand_frm.Controls("chk" & vCostSet).Enabled = False
            Next vCostSet
          End If
        Else
          '.Fields(0) is adVarChar in order to store "settings" todo: convert to .ini
          'FieldNameToFieldConstant somehow recognizes the string as a long
          If .Fields(0) >= 188776000 Then 'check enterprise field
            If FieldConstantToFieldName(.Fields(0)) <> Replace(.Fields(1), cptRegEx(.Fields(1), " \([A-z0-9]*\)$"), "") Then
              strMissing = strMissing & "- " & .Fields(1) & vbCrLf
              GoTo next_saved_field
            End If
          Else 'check local field
            If CustomFieldGetName(.Fields(0)) <> Trim(Replace(.Fields(1), cptRegEx(.Fields(1), "\([^\(].*\)$"), "")) Then
              'limit this check to Custom Fields
              If IsNumeric(Right(FieldConstantToFieldName(.Fields(0)), 1)) Then
                strMissing = strMissing & "- " & .Fields(1) & vbCrLf
                GoTo next_saved_field
              End If
            End If
          End If

          cptResourceDemand_frm.lboExport.AddItem
          cptResourceDemand_frm.lboExport.List(lngItem, 0) = .Fields(0) 'Field Constant
          cptResourceDemand_frm.lboExport.List(lngItem, 1) = .Fields(1) 'Custom Field Name
          lngItem = lngItem + 1
        End If
next_saved_field:
        .MoveNext
      Loop
      .Close
    End With
  End If
  
  cptResourceDemand_frm.Show False

  If Len(strMissing) > 0 Then
    MsgBox "The following saved fields do not exist in this project:" & vbCrLf & strMissing, vbInformation + vbOKOnly, "Saved Settings"
  End If

exit_here:
  On Error Resume Next
  Set rst = Nothing
  If rstResources.State Then rstResources.Close
  Set rstResources = Nothing
  Set objProject = Nothing
  If rstFields.State Then rstFields.Close
  Set rstFields = Nothing
  Exit Sub

err_here:
  If Err.Number = 1101 Or Err.Number = 1004 Then
    Err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptResourceDemand_bas", "cptShowExportResourceDemand_frm", Err, Erl)
    Resume exit_here
  End If

End Sub
