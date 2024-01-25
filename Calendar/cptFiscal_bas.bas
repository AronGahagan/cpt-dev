Attribute VB_Name = "cptFiscal_bas"
'<cpt_version>v1.1.0</cpt_version>
Option Explicit

Sub cptShowFiscal_frm()
'objects
Dim oException As MSProject.Exception
Dim oCal As MSProject.Calendar
'strings
Dim strSetting As String
Dim strExceptions As String
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  'get/create fiscal calendar
  On Error Resume Next
  Set oCal = ActiveProject.BaseCalendars("cptFiscalCalendar")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oCal Is Nothing Then
    BaseCalendarCreate Name:="cptFiscalCalendar", FromName:="Standard" ' [" & ActiveProject.Name & "]"
    Set oCal = ActiveProject.BaseCalendars("cptFiscalCalendar")
    If oCal.Exceptions.Count > 0 Then
      For Each oException In oCal.Exceptions
        oException.Delete
      Next oException
    End If
    cptFiscal_frm.txtExceptions.ControlTipText = "Paste list of fiscal end dates here (e.g., from a vertical list or column in Excel), with or without a corresponding label"
    cptFiscal_frm.lboExceptions.ControlTipText = "Paste list of fiscal end dates here (e.g., from a vertical list or column in Excel), with or without a corresponding label"
  Else
    If oCal.Exceptions.Count = 0 Then
      cptFiscal_frm.txtExceptions.ControlTipText = "Paste list of fiscal end dates here (e.g., from a vertical list or column in Excel), with or without a corresponding label"
      cptFiscal_frm.lboExceptions.ControlTipText = "Paste list of fiscal end dates here (e.g., from a vertical list or column in Excel), with or without a corresponding label"
    Else
      cptFiscal_frm.txtExceptions.ControlTipText = ""
      cptFiscal_frm.lboExceptions.ControlTipText = ""
    End If
  End If
  
  With cptFiscal_frm
    
    .lboExceptions.Clear
    .Caption = "Fiscal Calendar (" & cptGetVersion("cptFiscal_bas") & ")"
    'load exceptions
    '.lboExceptions.ColumnWidths = 45
    For Each oException In oCal.Exceptions
      .lboExceptions.AddItem oException.Start
      strExceptions = strExceptions & oException.Start & vbTab
      .lboExceptions.List(.lboExceptions.ListCount - 1, 1) = oException.Name
      strExceptions = strExceptions & oException.Name & vbCrLf
    Next
    
    'load headers
    .lboHeaders.AddItem "Fiscal End Date"
    .lboHeaders.List(0, 1) = "Label"
    '.lboHeaders.ColumnWidths = 45
  
    If .lboExceptions.ListCount = 0 Then
      .txtExceptions.Visible = True
      .lboExceptions.Visible = False
    Else
      .txtExceptions.Visible = False
      .lboExceptions.Visible = True
    End If
    
    .cmdImport.Enabled = False
    
    .lblCount.Caption = oCal.Exceptions.Count & " exception" & IIf(oCal.Exceptions.Count = 1, "", "s") & "."
    
    'warn if baseline or forecast finish exceends fiscal calendar
    If oCal.Exceptions.Count > 0 Then
      If IsDate(ActiveProject.BaselineSavedDate(pjBaseline)) Then
        If ActiveProject.ProjectSummaryTask.BaselineFinish > oCal.Exceptions(oCal.Exceptions.Count).Finish Then
          MsgBox "The project's baseline finish date is after the latest fiscal period end date.", vbInformation + vbOKOnly, "Heads up"
        End If
      End If
      If ActiveProject.ProjectSummaryTask.Finish > oCal.Exceptions(oCal.Exceptions.Count).Finish Then
        MsgBox "The project's forecast finish date is after the latest fiscal period end date.", vbInformation + vbOKOnly, "Heads up"
      End If
    End If
    
    .cmdAnalyzeEVT.Enabled = False
    strSetting = cptGetSetting("Integration", "EVT")
    If Len(strSetting) > 0 Then
      .cboUse.AddItem
      .cboUse.List(.cboUse.ListCount - 1, 0) = Split(strSetting, "|")(0)
      .cboUse.List(.cboUse.ListCount - 1, 1) = "EVT"
      .cboUse.List(.cboUse.ListCount - 1, 2) = Split(strSetting, "|")(1)
      .cboUse.Value = Split(strSetting, "|")(0)
      .cmdAnalyzeEVT.Enabled = True
    End If
    strSetting = cptGetSetting("Integration", "EVT_MS")
    If Len(strSetting) > 0 Then
      .cboUse.AddItem
      .cboUse.List(.cboUse.ListCount - 1, 0) = Split(strSetting, "|")(0)
      .cboUse.List(.cboUse.ListCount - 1, 1) = "EVT_MS"
      .cboUse.List(.cboUse.ListCount - 1, 2) = Split(strSetting, "|")(1)
      .cmdAnalyzeEVT.Enabled = True
    End If
    
    .Show 'Modal=True
    
  End With
    
exit_here:
  On Error Resume Next
  Set oException = Nothing
  Set oCal = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptShowFiscal_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportFiscalCalendar()
'objects
Dim oExcel As Excel.Application
Dim oWorkbook As Excel.Workbook
Dim oWorksheet As Excel.Worksheet
Dim oCalendar As MSProject.Calendar
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not cptCalendarExists("cptFiscalCalendar") Then
    MsgBox "cptFiscalCalendar has been deleted! Please re-open the form to re-create it.", vbCritical + vbOKOnly, "What happened?"
    GoTo exit_here
  End If
  
  Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
  If oCalendar.Exceptions.Count > 0 Then
    
    Application.StatusBar = "Getting Excel..."
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = "Fiscal Calendar"
    'add header
    Application.StatusBar = "Adding header..."
    oWorksheet.[A1:B1] = cptFiscal_frm.lboHeaders.List
    'export oExceptions
    oWorksheet.Range(oWorksheet.Cells(2, 1), oWorksheet.Cells(cptFiscal_frm.lboExceptions.ListCount + 1, 2)) = cptFiscal_frm.lboExceptions.List
    'make it pretty
    Application.StatusBar = "Formatting..."
    oWorksheet.ListObjects.Add xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)), False, xlYes 'xlSrcRange;
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.[A2].Select
    oExcel.ActiveWindow.FreezePanes = True
    oWorksheet.Columns.AutoFit
    Application.StatusBar = "Complete."
    Application.ActivateMicrosoftApp pjMicrosoftExcel
  Else
    Application.StatusBar = "Fiscal Calendar is empty"
    MsgBox "Fiscal Calendar has not yet been populated.", vbInformation + vbOKOnly, "No Exceptions"
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oCalendar = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  
  Exit Sub
  
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptExportCalendarExceptions", Err, Erl)
  Resume exit_here

End Sub

Sub cptExportExceptionsTemplate()
  'objects
  Dim oListObject As Excel.ListObject
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  Dim strMsg As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtLastFriday As Date
  
  strMsg = "Note: MPM's Fiscal Start dates must be adjusted to Fiscal End dates (in Excel) before importing."
  MsgBox strMsg, vbCritical + vbOKOnly, "MPM Warning"
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "Fiscal Calendar"
  oExcel.ActiveWindow.Zoom = 85
  oWorksheet.[A1:B1] = Split("fisc_end,label", ",")
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range("A1:B1"), , xlYes)
  oListObject.TableStyle = ""
  oListObject.ListColumns(1).Range.ColumnWidth = 12
  oListObject.ListColumns(2).Range.ColumnWidth = 12
  oWorksheet.[A2].Select
  oExcel.ActiveWindow.FreezePanes = True
  dtLastFriday = CDate("1/31/" & Year(Now))
  Do Until Weekday(dtLastFriday, vbSunday) = vbFriday
    dtLastFriday = DateAdd("d", -1, dtLastFriday)
  Loop
  oWorksheet.[A2:B2] = Array(dtLastFriday, "'" & Year(Now) & "01")
  
exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptExportExceptionsTemplate", Err)
  Resume exit_here
End Sub

Sub cptImportCalendarExceptions()
'objects
Dim oException As MSProject.Exception
Dim oCalendar As Calendar
Dim oCell As Excel.Range
Dim oWorksheet As Excel.Worksheet
Dim oWorkbook As Excel.Workbook
Dim fd As Object 'FileDialog
Dim oExcel As Excel.Application
'strings
Dim strSkipCalendar As String
'longs
'integers
'doubles
'booleans
'variants
'dates
Dim dtFiscalEnd As Date

  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Visible = True
  End If
  Set fd = oExcel.FileDialog(msoFileDialogFilePicker)
  With fd
    .AllowMultiSelect = False
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .InitialFileName = Environ("USERPROFILE") & "\"
    .Title = "Select Fiscal Calendar source file:"
    .Filters.Add "Microsoft Excel Workbook", "*.xls"
    .Filters.Add "Microsoft Excel Workbook", "*.xlsx"
    .Filters.Add "Microsoft Excel Macro-Enabled Workbook", "*.xlsm"
    .Filters.Add "Comma-Separated Values", "*.csv"
    If .Show = -1 Then
        
      Set oWorkbook = oExcel.Workbooks.Open(.SelectedItems(1))
      On Error Resume Next
      Set oWorksheet = oWorkbook.Sheets("Fiscal Calendar")
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If oWorksheet Is Nothing Then
        MsgBox "A worksheet named 'Fiscal Calendar' was not found in the selected workbook.", vbExclamation + vbOKOnly, "Invalid Selection"
        GoTo exit_here
      Else
        cptFiscal_frm.txtExceptions.Visible = False
        cptFiscal_frm.lboExceptions.Visible = True
        On Error Resume Next
        Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
        If oCalendar Is Nothing Then
          BaseCalendarCreate Name:="cptFiscalCalendar", FromName:="Standard" ' [" & ActiveProject.Name & "]"
          Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
          If oCalendar.Exceptions.Count > 0 Then
            For Each oException In oCalendar.Exceptions
              oException.Delete
            Next oException
          End If
        End If
        For Each oCell In oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].End(xlDown))
          If Not IsDate(oCell.Value) Then
            oCell.Style = "Bad"
            MsgBox "Row " & oCell.Row & " has invalid data and will be skipped.", vbExclamation + vbOKOnly, "Not a Date"
            GoTo next_record
          End If
          dtFiscalEnd = CDate(oCell.Value)
          On Error Resume Next
          Set oException = oCalendar.Exceptions.Add(Type:=pjDaily, Start:=CStr(dtFiscalEnd), Finish:=CStr(dtFiscalEnd), Name:=CStr(oCell.Offset(0, 1).Value))
          If oException Is Nothing Then
            MsgBox "Failed to add exception " & oCell.Value & " - " & oCell.Offset(0, 1).Value & "!", vbExclamation + vbOKOnly, "Unknown Error"
          Else
            cptFiscal_frm.lboExceptions.AddItem
            cptFiscal_frm.lboExceptions.List(cptFiscal_frm.lboExceptions.ListCount - 1, 0) = oException.Start 'CStr(oCell.Value)
            cptFiscal_frm.lboExceptions.List(cptFiscal_frm.lboExceptions.ListCount - 1, 1) = oException.Name 'CStr(oCell.Offset(0, 1).Value)
            cptFiscal_frm.lblCount.Caption = oCalendar.Exceptions.Count & " exception" & IIf(oCalendar.Exceptions.Count = 1, "", "s") & "."
          End If
          If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
next_record:
        Next oCell
        oWorkbook.Close False
      End If
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set oException = Nothing
  Set oCalendar = Nothing
  Set oCell = Nothing
  Set oWorksheet = Nothing
  oWorkbook.Close False
  Set oWorkbook = Nothing
  Set fd = Nothing
  oExcel.Quit
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptImportCalendarExceptions", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateFiscal()
'objects
Dim oCalendar As MSProject.Calendar
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With cptFiscal_frm
    If .txtExceptions.Visible = True Then GoTo exit_here
  
    Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
  
    With .lboExceptions
      .Clear
      For lngItem = 1 To oCalendar.Exceptions.Count
        .AddItem oCalendar.Exceptions(lngItem).Start
        .List(.ListCount - 1, 1) = oCalendar.Exceptions(lngItem).Name
      Next lngItem
    End With
  End With

exit_here:
  On Error Resume Next
  Set oCalendar = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptUpdateFiscal", Err, Erl)
  Resume exit_here
End Sub

Sub cptAnalyzeEVT(Optional lngImportField As Long)
  'objects
  Dim oRange As Excel.Range
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim rst As ADODB.Recordset
  Dim oException As MSProject.Exception
  Dim oCalendar As MSProject.Calendar
  Dim oProject As MSProject.Project
  Dim oTask As MSProject.Task
  'strings
  Dim strMissingBaselines As String
  Dim strLOE As String
  Dim strEVT As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  'longs
  Dim lngFiscalPeriodsCol As Long
  Dim lngFiscalEndCol As Long
  Dim lngLastRow As Long
  Dim lngFile As Long
  Dim lngEVT As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  'integers
  'doubles
  'booleans
  Dim blnExists As Boolean
  'variants
  Dim vbResponse As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If IsNull(cptFiscal_frm.cboUse) Then
    cptFiscal_frm.cboUse.BorderColor = 192
    GoTo exit_here
  Else
    cptFiscal_frm.cboUse.BorderColor = -2147483642
  End If
  
  Set oProject = ActiveProject
  
  'ensure project is baselined
  If Not IsDate(oProject.BaselineSavedDate(pjBaseline)) Then
    MsgBox "This project is not yet baselined.", vbCritical + vbOKOnly, "No Baseline"
    GoTo exit_here
  End If
  
  'ensure fiscal calendar is still loaded
  If Not cptCalendarExists("cptFiscalCalendar") Then
    MsgBox "The Fiscal Calendar (cptFiscalCalendar) is missing! Please reset it and try again.", vbCritical + vbOKOnly, "What happened?"
    GoTo exit_here
  End If
    
  'either EVT or EVT_MS
  strEVT = cptFiscal_frm.cboUse.List(cptFiscal_frm.cboUse.ListIndex, 1)
  lngEVT = CLng(cptFiscal_frm.cboUse.List(cptFiscal_frm.cboUse.ListIndex, 0))
  strLOE = cptGetSetting("Integration", "LOE")
  
  'todo: allow user to add other fields?
  
  'create the Schema.ini
  lngFile = FreeFile
  strFile = Environ("tmp") & "\Schema.ini"
  Open strFile For Output As #lngFile
  Print #1, "[fiscal.csv]"
  Print #1, "ColNameHeader=True"
  Print #1, "Format=CSVDelimited"
  Print #1, "Col1=FISCAL_END date"
  Print #1, "Col2=LABEL text"
  Print #1, "[tasks.csv]"
  Print #1, "ColNameHeader=True"
  Print #1, "Format=CSVDelimited"
  Print #1, "Col1=UID integer"
  Print #1, "Col2=BLS date"
  Print #1, "Col3=BLF date"
  Print #1, "Col4=" & strEVT & " text"
  Close #1
  
  'export the calendar
  Set oCalendar = ActiveProject.BaseCalendars("cptFiscalCalendar")
  lngFile = FreeFile
  strFile = Environ("tmp") & "\fiscal.csv"
  Open strFile For Output As #lngFile
  Print #lngFile, "fisc_end,label,"
  For Each oException In oCalendar.Exceptions
    Print #lngFile, oException.Finish & "," & oException.Name
  Next oException
  Close #lngFile
  
  'export discrete, PMB tasks
  lngFile = FreeFile
  strFile = Environ("tmp") & "\tasks.csv"
  Open strFile For Output As #lngFile
  Print #lngFile, "UID,BLS,BLF," & strEVT & ","
  For Each oTask In oProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.Assignments.Count = 0 Then GoTo next_task
    If oTask.BaselineWork = 0 And oTask.BaselineCost = 0 Then GoTo next_task
    If oTask.GetField(lngEVT) = strLOE Then GoTo next_task
    If Len(oTask.GetField(lngEVT)) = 0 And strEVT = "EVT_MS" Then GoTo next_task
    If Not IsDate(oTask.BaselineStart) Or Not IsDate(oTask.BaselineFinish) Then
      strMissingBaselines = strMissingBaselines = oTask.UniqueID & ","
    End If
    Print #lngFile, oTask.UniqueID & "," & FormatDateTime(oTask.BaselineStart, vbShortDate) & "," & FormatDateTime(oTask.BaselineFinish, vbShortDate) & "," & oTask.GetField(lngEVT)
next_task:
  Next oTask
  Close #lngFile
  
  If Len(strMissingBaselines) > 0 Then
    Debug.Print "MISSING BASELINES: " & strMissingBaselines
  End If
  
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  oExcel.Visible = True
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "EVT Analysis"
  oWorksheet.[A1:E1] = Split("UID,BLS,BLF," & strEVT & ",FiscalPeriods", ",")
  
  Set rst = CreateObject("ADODB.Recordset")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & Environ("tmp") & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT * FROM [tasks.csv]"
  rst.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  oWorksheet.[A2].CopyFromRecordset rst
  rst.Close
  
  strSQL = "SELECT * FROM [fiscal.csv]"
  rst.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  oWorksheet.[G1:H1] = Split("fisc_end,label", ",")
  oWorksheet.[G2].CopyFromRecordset rst
  rst.Close
  
  Set oRange = oWorksheet.Range(oWorksheet.[A1].End(xlToRight).Offset(1, 0), oWorksheet.[A1].End(xlDown).Offset(0, 4))
  lngFiscalEndCol = oWorksheet.rows(1).Find(what:="fisc_end").Column
  lngLastRow = oWorksheet.Cells(2, lngFiscalEndCol).End(xlDown).Row
  'Excel 2016 compatibility
  'oRange.FormulaR1C1 = "=COUNTIFS(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ","">=""&RC[-3],R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ",""<""&RC[-2])+1"
  '=SUMPRODUCT(--($G$2:$G$109>=B15)*--($G$2:$G$109<C15)*1)+1
  oRange.FormulaR1C1 = "=SUMPRODUCT(--(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & ">=RC[-3])*--(R2C" & lngFiscalEndCol & ":R" & lngLastRow & "C" & lngFiscalEndCol & "<RC[-2])*1)+1"
  lngFiscalPeriodsCol = oWorksheet.rows(1).Find(what:="FiscalPeriods").Column
  oWorksheet.Columns(lngFiscalPeriodsCol).NumberFormat = "#0"
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.[A1].AutoFilter
  oWorksheet.Columns.AutoFit
  If lngImportField > 0 Then
    cptSpeed True
    Dim oCell As Excel.Range
    lngLastRow = oWorksheet.Cells(1048576, 1).End(xlUp).Row
    Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.Cells(lngLastRow, 1))
    lngTasks = oRange.Cells.Count
    lngTask = 1
    For Each oCell In oRange.Cells
      Set oTask = ActiveProject.Tasks.UniqueID(oCell.Value)
      oTask.SetField lngImportField, oCell.Offset(0, 4)
      lngTask = lngTask + 1
      cptFiscal_frm.lblProgress.Width = (lngTask / lngTasks) * cptFiscal_frm.lblStatus.Width
      cptFiscal_frm.lblStatus.Caption = "Importing...(" & Format(lngTask / lngTasks, "0%") & ")"
    Next oCell
    cptFiscal_frm.lblStatus.Caption = "Complete"
    cptFiscal_frm.lblProgress.Width = cptFiscal_frm.lblStatus.Width
    cptSpeed False
  Else
    MsgBox "Copy UIDs from Excel to 'Filter By Clipboard' to apply bulk " & strEVT & " changes.", vbInformation + vbOKOnly, "Hint:"
  End If
  Application.ActivateMicrosoftApp (pjMicrosoftExcel)
  
exit_here:
  On Error Resume Next
  cptSpeed False
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  For lngFile = 1 To FreeFile
    Close #lngFile
  Next lngFile
  If Dir(Environ("tmp") & "\Schema.ini") <> vbNullString Then
    Kill Environ("tmp") & "\Schema.ini"
  End If
  If Dir(Environ("tmp") & "\fiscal.csv") <> vbNullString Then
    Kill Environ("tmp") & "\fiscal.csv"
  End If
  If Dir(Environ("tmp") & "\tasks.csv") <> vbNullString Then
    Kill Environ("tmp") & "\tasks.csv"
  End If
  If rst.State = 1 Then rst.Close
  Set rst = Nothing
  Set oException = Nothing
  Set oCalendar = Nothing
  Set oTask = Nothing
  Set oProject = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal", "cptAnalyzeEVT", Err, Erl)
  Resume exit_here
End Sub
