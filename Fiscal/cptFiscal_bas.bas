Attribute VB_Name = "cptFiscal_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowCptFiscal_frm()
'objects
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'load data
  If Dir(cptDir & "\cpt-fiscal.adtg") <> vbNullString Then
    'load fiscal calendar
  End If
  
  With cptFiscal_frm
    'load import options
    .cboImport.AddItem "COBRA Export"
    .cboImport.AddItem "MPM Export"
    .cboImport.AddItem "Custom"
    
    'load export options
    .cboExport.AddItem "For COBRA"
    .cboExport.AddItem "For MPM"
    .cboExport.AddItem "To Excel Workbook"

    'load calendars
    .lboCalendars.Clear
    For lngItem = 1 To ActiveProject.BaseCalendars.Count
      .lboCalendars.AddItem ActiveProject.BaseCalendars(lngItem).Name
    Next
    
    .Show False
    
  End With
    
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptShowCptFiscalFrm", err, Erl)
  Resume exit_here
End Sub

Sub cptExportCalendarExceptions()
'objects
Dim xlApp As Object 'Excel.Application
Dim Workbook As Object 'Workbook
Dim Worksheet As Worksheet
Dim Calendar As Calendar
Dim Exception As Exception
'strings
'longs
Dim lngRow As Long
Dim lngCalendar As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = "Calendar Exceptions"
  'add header
  Worksheet.[A1:D1] = Array("Calendar", "Name", "Start", "Finish")
  'export exceptions
  For lngCalendar = 1 To ActiveProject.BaseCalendars.Count
    Set Calendar = ActiveProject.BaseCalendars(lngCalendar)
    For Each Exception In Calendar.Exceptions
      lngRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Row + 1
      Worksheet.Cells(lngRow, 1) = Calendar.Name
      Worksheet.Cells(lngRow, 2) = Exception.Name
      Worksheet.Cells(lngRow, 3) = Exception.Start
      Worksheet.Cells(lngRow, 4) = Exception.Finish
    Next Exception
  Next lngCalendar
  'make it pretty
  Worksheet.ListObjects.Add 1, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), 1
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  Worksheet.Columns.AutoFit
  xlApp.Visible = True

exit_here:
  On Error Resume Next
  Set Exception = Nothing
  Set Calendar = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing

  Exit Sub
  
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptExportCalendarExceptions", err, Erl)
  Resume exit_here

End Sub

Sub cptExportExceptionsTemplate()
'objects
Dim ListObject As Object
Dim Worksheet As Object
Dim Workbook As Object
Dim xlApp As Object
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.Name = "Calendar Exceptions"
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.[A1:D1] = Array("Calendar", "Name", "Start", "Finish")
  Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range("A1:C1"), , xlYes)
  Worksheet.Columns(1).ColumnWidth = 33.72
  Worksheet.Columns(2).ColumnWidth = 33.72
  Worksheet.Columns(3).ColumnWidth = 12
  Worksheet.Columns(4).ColumnWidth = 12
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  Worksheet.[A2:D2] = Array("MyProject Calendar", "Independence Day", #7/3/2020#, #7/3/2020#)
  xlApp.Visible = True
  Application.ActivateMicrosoftApp pjMicrosoftExcel
  
exit_here:
  On Error Resume Next
  Set ListObject = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptExportExceptionsTemplate", err)
  Resume exit_here
End Sub

Sub cptImportCalendarExceptions()
'objects
Dim Calendar As Calendar
Dim c As Object
Dim Worksheet As Object
Dim Workbook As Object
Dim fd As Object 'FileDialog
Dim xlApp As Object 'Excel Application
'strings
Dim strSkipCalendar As String
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  Set fd = xlApp.FileDialog(msoFileDialogFilePicker)
  With fd
    .AllowMultiSelect = False
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .InitialFileName = Environ("USERPROFILE") & "\"
    .Title = "Select Calendar Exceptions source file:"
    .Filters.Add "Microsoft Excel Workbook", "*.xlsx"
    .Filters.Add "Microsoft Excel Macro-Enabled Workbook", "*.xlsm"
    If .Show = -1 Then
        
      Set Workbook = xlApp.Workbooks.Open(.SelectedItems(1))
      On Error Resume Next
      Set Worksheet = Workbook.Sheets("Calendar Exceptions")
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If Worksheet Is Nothing Then
        MsgBox "A worksheet named 'Calendar Exceptions' was not found in the selected workbook.", vbExclamation + vbOKOnly, "Invalid Selection"
        GoTo exit_here
      Else
        For Each c In Worksheet.Range(Worksheet.[A2], Worksheet.[A2].End(xlDown))
          On Error Resume Next
          Set Calendar = ActiveProject.BaseCalendars(c.Value)
          If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          If Calendar Is Nothing Then
            MsgBox "Calendar '" & c.Value & "' not found in this project. Associated exceptions will be skipped.", vbExclamation + vbOKOnly
            strSkipCalendar = c.Value
          End If
          If c.Value <> strSkipCalendar Then
            On Error Resume Next
            Calendar.Exceptions.Add Type:=pjDaily, Start:=CStr(c.Offset(0, 2).Value), Finish:=CStr(c.Offset(0, 3).Value), Name:=c.Offset(0, 1).Value
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
          End If
          Set Calendar = Nothing
        Next c
        Workbook.Close False
      End If
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set Calendar = Nothing
  Set c = Nothing
  Set Worksheet = Nothing
  Workbook.Close False
  Set Workbook = Nothing
  Set fd = Nothing
  xlApp.Quit
  Set xlApp = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_bas", "cptImportCalendarExceptions", err, Erl)
  Resume exit_here
End Sub
