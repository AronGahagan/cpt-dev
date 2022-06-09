Attribute VB_Name = "cptBackbone_bas"
'<cpt_version>v1.1.3</cpt_version>
Option Explicit

Sub cptImportCWBSFromExcel(lngOutlineCode As Long)
  'objects
  Dim oTask As Task
  Dim oLookupTable As LookupTable
  Dim oOutlineCode As OutlineCode
  Dim c As Object
  Dim oRange As Object
  Dim oFileDialog As Object 'FileDialog
  Dim oWorksheet As Object
  Dim oWorkbook As Object
  Dim oExcel As Object 'Excel.Application
  'strings
  Dim strOutlineCode As String
  'longs
  Dim lngItems As Long
  Dim lngOutlineLevel As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    
  If MsgBox("Epected fields/column headers, in range [A1:C1], are CODE,LEVEL,DESCRIPTION and there should be no blank rows." & vbCrLf & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Confirm CWBS Import") = vbNo Then
    'export a sample template
    If MsgBox("Would you like an example?", vbQuestion + vbYesNo, "A little help") = vbYes Then Call cptExportTemplate
  Else
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    Set oExcel = CreateObject("Excel.Application")
    'allow user to select excel file and import it to chosen
    Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
    With oFileDialog
      .AllowMultiSelect = False
      .ButtonName = "Import"
      .InitialView = msoFileDialogViewDetails
      .InitialFileName = Environ("USERPROFILE") & "\"
      .Title = "Select " & strOutlineCode & " source file:"
      .Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
      .Filters.Add "Comma Separated Values (csv)", "*.csv"
      If .Show = -1 Then
      
        Application.OpenUndoTransaction "Import " & strOutlineCode & " from Excel Workbook"
      
        cptSpeed True
      
        'set up the outline code field
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=2, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=3, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=4, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=5, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=6, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=7, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=8, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=9, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=10, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
        
        Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
        'open the workbook
        Set oWorkbook = oExcel.Workbooks.Open(oFileDialog.SelectedItems(1))
        'find the sheet
        For Each oWorksheet In oWorkbook.Sheets
          If oWorksheet.[A1].Value = "CODE" And oWorksheet.[B1].Value = "LEVEL" And oWorksheet.[C1].Value = "DESCRIPTION" Then
            strOutlineCode = CustomFieldGetName(lngOutlineCode)
            Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162)) '-4162 = xlUp
            lngItems = oRange.Cells.Count
            lngItem = 0
            For Each c In oRange.Cells
              lngItem = lngItem + 1
              Set oTask = ActiveProject.Tasks.Add(c.Offset(0, 2).Value)
              oTask.SetField lngOutlineCode, c.Value
              If oOutlineCode Is Nothing Then Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
              If oLookupTable Is Nothing Then Set oLookupTable = oOutlineCode.LookupTable
              oLookupTable.Item(lngItem).Description = c.Offset(0, 2)
              If cptBackbone_frm.chkAlsoCreateTasks Then
                lngOutlineLevel = Len(c.Value) - Len(Replace(c.Value, ".", ""))
                If lngOutlineLevel > 0 Then
                  oTask.OutlineLevel = lngOutlineLevel + 1
                End If
              Else
                oTask.Delete
              End If
              cptBackbone_frm.lblStatus.Caption = "Importing " & lngItem & " of " & lngItems & "(" & Format(lngItem / lngItems, "0%") & ")..."
              cptBackbone_frm.lblProgress.Width = (lngItem / lngItems) * cptBackbone_frm.lblStatus.Width
            Next c
            cptBackbone_frm.lblStatus.Caption = "Ready..."
            cptBackbone_frm.lblProgress.Width = cptBackbone_frm.lblStatus.Width
            'reset outline code to disallow new entries
            CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
            'refresh the form
            cptBackbone_frm.cboOutlineCodes.List(cptBackbone_frm.cboOutlineCodes.ListIndex, 1) = FieldConstantToFieldName(lngOutlineCode) & " (" & strOutlineCode & ")"
            Exit For
          End If
        Next oWorksheet
      Else
        MsgBox "No worksheet found where [A1:C1] contains CODE, LEVEL, DESCRIPTION.", vbExclamation + vbOKOnly, "Invalid Workbook"
      End If 'proper headers found
    End With
  End If 'proceed

exit_here:
  On Error Resume Next
  cptSpeed False
  Application.CloseUndoTransaction
  Set oTask = Nothing
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Set c = Nothing
  Set oRange = Nothing
  Set oFileDialog = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  oWorkbook.Close False
  oExcel.Quit
  Set oExcel = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptImportCWBSFromExcel", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportCWBSFromServer(lngOutlineCode As Long)
  'objects
  Dim c As Object
  Dim oTask As Task
  Dim oRange As Object
  Dim oWorksheet As Object
  Dim oWorkbook As Object
  Dim oLookupTable As LookupTable
  Dim oOutlineCode As OutlineCode
  Dim oFileDialog As Object 'FileDialog
  Dim oExcel As Object
  'strings
  Dim strDescription As String
  Dim strCode As String
  Dim strOutlineCode As String
  'longs
  Dim lngItems As Long
  Dim lngOutlineLevel As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Expected fields/column headers, in range [A1:C1], are LEVEL,VALUE,DESCRIPTION and there should be no blank rows." & vbCrLf & vbCrLf & "Proceed?", vbQuestion + vbYesNo, "Confirm CWBS Import") = vbYes Then
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    Set oExcel = CreateObject("Excel.Application")
    'allow user to select excel file and import it to chosen
    Set oFileDialog = oExcel.FileDialog(msoFileDialogFilePicker)
    With oFileDialog
      .AllowMultiSelect = False
      .ButtonName = "Import"
      .InitialView = 2 'msoFileDialogViewDetails
      .InitialFileName = Environ("USERPROFILE") & "\"
      .Title = "Select " & strOutlineCode & " source file:"
      .Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
      .Filters.Add "Comma Separated Values (csv)", "*.csv"
      If .Show = -1 Then
      
        Application.OpenUndoTransaction "Import " & strOutlineCode & " from MSP Server Outline Code Export"
      
        cptSpeed True
      
        'set up the outline code field
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=2, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=3, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=4, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=5, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=6, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=7, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=8, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=9, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=10, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
        CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
        
        Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
        'open the workbook
        Set oWorkbook = oExcel.Workbooks.Open(oFileDialog.SelectedItems(1))
        'find the sheet
        For Each oWorksheet In oWorkbook.Sheets
          If UCase(oWorksheet.[A1].Value) = "LEVEL" And UCase(oWorksheet.[B1].Value) = "VALUE" And UCase(oWorksheet.[C1].Value) = "DESCRIPTION" Then
            strOutlineCode = CustomFieldGetName(lngOutlineCode)
            Set oRange = oWorksheet.Range(oWorksheet.[A2], oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162)) '-4162 = xlUp
            lngItems = oRange.Cells.Count
            lngItem = 0
            For Each c In oRange.Cells
              lngItem = lngItem + 1
              If oOutlineCode Is Nothing Then Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
              If oLookupTable Is Nothing Then Set oLookupTable = oOutlineCode.LookupTable
              Set oTask = ActiveProject.Tasks.Add(c.Offset(0, 2).Value)
              strCode = Left(c.Offset(0, 2), InStr(c.Offset(0, 2), " ") - 1)
              strDescription = Replace(c.Offset(0, 2), strCode & " - ", "")
              oTask.SetField lngOutlineCode, strCode
              oLookupTable.Item(lngItem).Description = strDescription
              cptBackbone_frm.lblStatus.Caption = "Importing " & lngItem & " of " & lngItems & "(" & Format(lngItem / lngItems, "0%") & ")..."
              cptBackbone_frm.lblProgress.Width = (lngItem / lngItems) * cptBackbone_frm.lblStatus.Width
            Next c
            cptBackbone_frm.lblStatus.Caption = "Ready..."
            cptBackbone_frm.lblProgress.Width = cptBackbone_frm.lblStatus.Width
            'reset outline code to disallow new entries
            CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
            'refresh the form
            cptBackbone_frm.cboOutlineCodes.List(cptBackbone_frm.cboOutlineCodes.ListIndex, 1) = FieldConstantToFieldName(lngOutlineCode) & " (" & strOutlineCode & ")"
            Exit For
          End If
        Next oWorksheet
      Else
        MsgBox "No worksheet found where [A1:C1] contains LEVEL, VALUE, DESCRIPTION.", vbExclamation + vbOKOnly, "Invalid Workbook"
      End If 'proper headers found
    End With
  End If 'proceed

exit_here:
  On Error Resume Next
  cptSpeed False
  Set c = Nothing
  Set oTask = Nothing
  Set oRange = Nothing
  Set oWorksheet = Nothing
  oWorkbook.Close False
  Set oWorkbook = Nothing
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  Set oFileDialog = Nothing
  oExcel.Quit False
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptImportCWBSFromServer", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportAppendixB(lngOutlineCode As Long)
  'objects
  Dim TaskTable As Object 'TaskTable
  Dim Task As Task
  'strings
  'longs
  Dim lngItem As Long
  Dim lngField As Long
  Dim lngOutlineLevel As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Application.OpenUndoTransaction "Import Appendix B"
  
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=2, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=3, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=4, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=5, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=6, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=7, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=8, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=9, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=10, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0

  With CreateObject("ADODB.Recordset")
    'delete existing
    If Dir(cptDir & "\cwbs-appendix-b.adtg") <> vbNullString Then
      Kill cptDir & "\cwbs-appendix-b.adtg"
    End If
    .Fields.Append "CWBS", adVarChar, 10
    .Fields.Append "CWBS TITLE", adVarChar, 75
    .Open
    .AddNew Array("CWBS", "CWBS Title"), Array("1", "Electronic System/Generic System")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1", "Prime Mission Product (PMP) 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.1", "PMP Integration, Assembly, Test, and Checkout")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.2", "PMP Subsystem 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.2.1", "Subsystem Integration, Assembly, Test, and Checkout")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.2.2", "Subsystem Hardware 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.2.3", "Subsystem Software Release 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.3", "PMP Software Release 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.3.1", "Computer Software Configuration Item (CSCI) 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.3.2", "PMP Software Integration, Assembly, Test, and Checkout")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2", "Platform Integration, Assembly, Test, and Checkout")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3", "Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.1", "Software Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.2", "Integrated Logistics Support (ILS) Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.3", "Cybersecurity Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.4", "Core Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.5", "Other Systems Engineering 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4", "Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.1", "Software Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.2", "Integrated Logistics Support (ILS) Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.3", "Cybersecurity Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.4", "Core Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.5", "Other Program Management 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5", "System Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.1", "Development Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.2", "Operational Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.3", "Cybersecurity Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.4", "Mock-ups/System Integration Labs (SILs)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.5", "Test and Evaluation Support")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.6", "Test Facilities")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6", "Training")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.1", "Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.1.1", "Operator Instructional Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.1.2", "Maintainer Instructional Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.2", "Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.2.1", "Operator Instructional Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.2.2", "Maintainer Instructional Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.3", "Facilities")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.4", "Training Software 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7", "Data")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.1", "Data Deliverables 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.2", "Data Repository")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.3", "Data Rights 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8", "Peculiar Support Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1", "Test and Measurement Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.1", "Test and Measurement Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.2", "Test and Measurement Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.3", "Test and Measurement Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.4", "Test and Measurement Equipment (Other Major Subsystems 1...n (Sif))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2", "Support and Handling Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.1", "Support and Handling Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.2", "Support and Handling Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.3", "Support and Handling Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.4", "Support and Handling Equipment (Other Major Subsystems 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9", "Common Support Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.1", "Test and Measurement Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.1.1", "Test and Measurement Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.1.2", "Test and Measurement Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.1.3", "Test and Measurement Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.1.4", "Test and Measurement Equipment (Other Major Subsystems 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.2", "Support and Handling Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.2.1", "Support and Handling Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.2.2", "Support and Handling Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.2.3", "Support and Handling Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.2.4", "Support and Handling Equipment (Other Major Subsystems 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10", "Operational/Site Activation by Site 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10.1", "System Assembly, Installation, and Checkout on Site")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10.2", "Contractor Technical Support")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10.3", "Site Construction")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10.4", "Site /Ship/Vehicle Conversion")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10.5", "Interim Contractor Support (ICS)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.11", "Contractor Logistics Support (CLS)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.12", "Industrial Facilities")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.12.1", "Construction/Conversion/Expansion")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.12.2", "Equipment Acquisition or Modernization")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.12.3", "Maintenance (Industrial Facilities)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.13", "Initial Spares and Repair Parts")
    .Save cptDir & "\cwbs-appendix-b.adtg"
    .MoveFirst
    lngItem = 0
    Do While Not .EOF
      lngItem = lngItem + 1
      Set Task = ActiveProject.Tasks.Add(.Fields(1).Value)
      Task.SetField lngOutlineCode, .Fields(0)
      ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode)).LookupTable.Item(lngItem).Description = .Fields(1).Value

      lngOutlineLevel = Len(.Fields(0).Value) - Len(Replace(.Fields(0).Value, ".", ""))
      If lngOutlineLevel > 0 Then
        Task.OutlineLevel = lngOutlineLevel + 1
      End If
      
      .MoveNext
    Loop
    .Close
  End With
  
  'pretty up the task table
  If Len(ActiveProject.CurrentTable) > 0 Then
    SelectBeginning
    SetRowHeight 1, "all"
    Set TaskTable = ActiveProject.TaskTables(ActiveProject.CurrentTable)
    For lngField = 1 To TaskTable.TableFields.Count
      If FieldConstantToFieldName(TaskTable.TableFields(lngField).Field) = "Name" Then
        ColumnBestFit lngField
        Exit For
      End If
    Next lngField
  End If
  
  'reset outline code to disallow new entries
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
  Call cptRefreshOutlineCodePreview(CustomFieldGetName(lngOutlineCode))
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set TaskTable = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptImportAppendixB", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportAppendixE(lngOutlineCode As Long)
  'objects
  Dim TaskTable As Object 'TaskTable
  Dim Task As Task
  'strings
  'longs
  Dim lngItem As Long
  Dim lngField As Long
  Dim lngOutlineLevel As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Application.OpenUndoTransaction "Import Appendix E"
  
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=2, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=3, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=4, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=5, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=6, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=7, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=8, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=9, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, Level:=10, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0

  With CreateObject("ADODB.Recordset")
    'delete existing
    If Dir(cptDir & "\cwbs-appendix-e.adtg") <> vbNullString Then
      Kill cptDir & "\cwbs-appendix-e.adtg"
    End If
    .Fields.Append "CWBS", adVarChar, 10
    .Fields.Append "CWBS TITLE", adVarChar, 75
    .Open
    .AddNew Array("CWBS", "CWBS Title"), Array("1", "Sea System")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1", "Ship")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.1", "Hull Structure")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.2", "Propulsion Plant")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.3", "Electric Plant")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.4", "Command, Communications, and Surveillance")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.5", "Auxiliary Systems")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.6", "Outfit and Furnishings")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.7", "Armament")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.8", "Total Ship Integration/Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.1.9", "Ship Assembly and Support Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2", "Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2.1", "Software Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2.2", "Integrated Logistics Support (ILS) Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2.3", "Cybersecurity Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2.4", "Core Systems Engineering")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.2.5", "Other Systems Engineering 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3", "Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.1", "Software Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.2", "Integrated Logistics Support (ILS) Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.3", "Cybersecurity Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.4", "Core Program Management")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.3.5", "Other Program Management 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4", "System Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.1", "Development Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.2", "Operational Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.3", "Cybersecurity Test and Evaluation")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.4", "Mock-ups/System Integration Labs (SILs)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.5", "Test and Evaluation Support")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.4.6", "Test Facilities")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5", "Training")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.1", "Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.1.1", "Operator Instructional Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.1.2", "Maintainer Instructional Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.2", "Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.2.1", "Operator Instructional Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.2.2", "Maintainer Instructional Services")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.3", "Facilities")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.5.4", "Training Software 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6", "Data")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.1", "Data Deliverables 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.2", "Data Repository")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.6.3", "Data Rights 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7", "Peculiar Support Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.1", "Test and Measurement Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.1.1", "Test and Measurement Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.1.2", "Test and Measurement Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.1.3", "Test and Measurement Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.1.4", "Test and Measurement Equipment (Other Major Subsystem 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.2", "Support and Handling Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.2.1", "Support and Handling Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.2.2", "Support and Handling Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.2.3", "Support and Handling Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.7.2.4", "Support and Handling Equipment (Other Major Subsystem 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8", "Common Support Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1", "Test and Measurement Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.1", "Test and Measurement Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.2", "Test and Measurement Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.3", "Test and Measurement Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.1.4", "Test and Measurement Equipment (Other Major Subsystem 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2", "Support and Handling Equipment")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.1", "Support and Handling Equipment (Airframe/Hull/Vehicle)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.2", "Support and Handling Equipment (Propulsion)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.3", "Support and Handling Equipment (Electronics/Avionics)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.8.2.4", "Support and Handling Equipment (Other Major Subsystem 1...n (Specify))")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9", "Operational/Site Activation by Site 1...n (Specify)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.1", "System Assembly, Installation, and Checkout")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.2", "Contractor Technical Support")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.3", "Site Construction")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.4", "Site/Ship/Vehicle Conversion")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.9.5", "Interim Contractor Support (ICS)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.10", "Contractor Logistics Support (CLS)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.11", "Industrial Facilities")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.11.1", "Construction/Conversion/Expansion")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.11.2", "Equipment Acquisition or Modernization")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.11.3", "Maintenance (Industrial Facilities)")
    .AddNew Array("CWBS", "CWBS Title"), Array("1.12", "Initial Spares and Repair Parts")
    .Save cptDir & "\cwbs-appendix-e.adtg"
    .MoveFirst
    lngItem = 0
    Do While Not .EOF
      lngItem = lngItem + 1
      Set Task = ActiveProject.Tasks.Add(.Fields(1).Value)
      Task.SetField lngOutlineCode, .Fields(0)
      ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode)).LookupTable.Item(lngItem).Description = .Fields(1).Value

      lngOutlineLevel = Len(.Fields(0).Value) - Len(Replace(.Fields(0).Value, ".", ""))
      If lngOutlineLevel > 0 Then
        Task.OutlineLevel = lngOutlineLevel + 1
      End If
      
      .MoveNext
    Loop
    .Close
  End With
  
  'pretty up the task table
  If Len(ActiveProject.CurrentTable) > 0 Then
    SelectBeginning
    SetRowHeight 1, "all"
    Set TaskTable = ActiveProject.TaskTables(ActiveProject.CurrentTable)
    For lngField = 1 To TaskTable.TableFields.Count
      If FieldConstantToFieldName(TaskTable.TableFields(lngField).Field) = "Name" Then
        ColumnBestFit lngField
        Exit For
      End If
    Next lngField
  End If
  
  'reset outline code to disallow new entries
  CustomOutlineCodeEditEx FieldID:=lngOutlineCode, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
  Call cptRefreshOutlineCodePreview(CustomFieldGetName(lngOutlineCode))
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set TaskTable = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptImportAppendixE", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportOutlineCodeToExcel(lngOutlineCode As Long)
  'objects
  Dim oExcel As Object 'Excel.Application
  Dim oWorkbook As Object 'Workbook
  Dim oWorksheet As Object 'Worksheet
  Dim oListObject As Object 'ListObject
  Dim oLookupTable As LookupTable
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strOutlineCode As String
  'longs
  Dim lngLastRow As Long
  Dim lngLookupItems As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strOutlineCode = CustomFieldGetName(lngOutlineCode)
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oLookupTable Is Nothing Then
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbCritical + vbOKOnly, "No Code Defined"
    GoTo exit_here
  End If
  Application.StatusBar = "Exporting Outline Code '" & strOutlineCode & "'..."
  cptBackbone_frm.lblStatus.Caption = Application.StatusBar
  
  'get excel
  Application.StatusBar = "Setting up Excel..."
  cptBackbone_frm.lblStatus.Caption = Application.StatusBar
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = -4135 'xlCalculationManual
  oExcel.ScreenUpdating = False
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Outline.SummaryRow = 0 'xlSummaryAbove
  oWorksheet.[A1:C1] = Array("CODE", "LEVEL", "DESCRIPTION")
  
  'export the codes
  For lngLookupItems = 1 To oLookupTable.Count
    lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(-4162).Row + 1 '-4162 = xlUp
    oWorksheet.Cells(lngLastRow, 1).Value = "'" & oLookupTable.Item(lngLookupItems).FullName
    oWorksheet.Cells(lngLastRow, 2).Value = oLookupTable.Item(lngLookupItems).Level
    oWorksheet.Cells(lngLastRow, 3).Value = oLookupTable.Item(lngLookupItems).Description
    oWorksheet.Cells(lngLastRow, 3).IndentLevel = oLookupTable.Item(lngLookupItems).Level - 1
    If oLookupTable.Item(lngLookupItems).Level > 8 Then
      oWorksheet.Rows(lngLastRow).OutlineLevel = 8
      oWorksheet.Cells(lngLastRow, 2).AddComment "Excel grouping limited to 8 levels"
    Else
      oWorksheet.Rows(lngLastRow).OutlineLevel = oLookupTable.Item(lngLookupItems).Level
    End If
    cptBackbone_frm.lblProgress.Width = (lngLookupItems / oLookupTable.Count) * cptBackbone_frm.lblStatus.Width
    cptBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngLookupItems / oLookupTable.Count, "0%") & ")"
  Next lngLookupItems
  
  Application.StatusBar = "Formatting Worksheet..."
  cptBackbone_frm.lblStatus.Caption = Application.StatusBar
  
  'format the table
  oExcel.ActiveWindow.Zoom = 85
  'Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), , xlYes)
  Set oListObject = oWorksheet.ListObjects.Add(1, oWorksheet.Range(oWorksheet.[A1].End(-4161), oWorksheet.[A1].End(-4121)), , 1)
  oListObject.Name = strOutlineCode
  oListObject.TableStyle = ""
  oListObject.HeaderRowRange.Font.Bold = True
  oListObject.Range.Borders(5).LineStyle = -4142 'xlDiagonalDown = xlNone
  oListObject.Range.Borders(6).LineStyle = -4142 'xlDiagonalUp = xlNone
  With oListObject.Range.Borders(7) 'xlEdgeLeft
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(8) 'xlEdgeTop
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(9) 'xlEdgeBottom
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(10) 'xlEdgeRight
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(11) 'xlInsideVertical
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = 2 'xlThin
  End With
  With oListObject.Range.Borders(12) 'xlInsideHorizontal
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = 2 'xlThin
  End With
  With oListObject.HeaderRowRange.Interior
    .Pattern = 1 'xlSolid
    .PatternColorIndex = -4105 'xlAutomatic
    .ThemeColor = 1 'xlThemeColorDark1
    .TintAndShade = -0.149998474074526
    .PatternTintAndShade = 0
  End With
  oWorksheet.Name = strOutlineCode
  oWorksheet.[A2].Select
  oExcel.ActiveWindow.FreezePanes = True
  oWorksheet.Columns.AutoFit
    
exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing
  Application.StatusBar = "Ready..."
  cptBackbone_frm.lblStatus.Caption = Application.StatusBar
  cptBackbone_frm.lblProgress.Width = cptBackbone_frm.lblStatus.Width
  oExcel.Visible = True
  oExcel.ScreenUpdating = True
  oExcel.Calculation = -4105 'xlCalculationAutomatic
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oOutlineCode = Nothing

  Exit Sub
  
err_here:
  Call cptHandleErr("cptBackbone_bas", "ExportOutlineCode", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptExport81334D(lngOutlineCode As Long)
  'objects
  Dim oMailItem As Object 'MailItem
  Dim oOutlook As Object 'Outlook.Application
  Dim oLookupTable As LookupTable
  Dim oOutlineCode As OutlineCode
  Dim wsDictionary As Object 'Worksheet
  Dim wsIndex As Object 'Worksheet
  Dim oWorkbook As Object 'Workbok
  Dim oExcel As Object 'Excel.Application
  Dim oStream As Object 'ADODB.Stream
  Dim oXMLHttpDoc As Object
  Dim oShell As Object
  'strings
  Dim strOutlineCode As String
  Dim strURL As String
  Dim strTemplateDir As String
  Dim strTemplate As String
  'longs
  Dim lngBorder As Long
  Dim lngRow As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'get outline code name and export it
  cptBackbone_frm.lblStatus.Caption = "Exporting..."
  strOutlineCode = CustomFieldGetName(lngOutlineCode)
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbCritical + vbOKOnly, "No Code Defined"
    GoTo exit_here
  Else
  
    'first determine if user has the template installed
    Set oShell = CreateObject("WScript.Shell")
    strTemplateDir = oShell.SpecialFolders("Templates")
    strTemplate = "81334D_CWBS_TEMPLATE.xltm"
    
    If Dir(strTemplateDir & "\" & strTemplate) = vbNullString Then
      'provide user feedback
      cptBackbone_frm.lblStatus.Caption = "Downloading template..."
      Set oXMLHttpDoc = CreateObject("Microsoft.XMLHTTP")
      strURL = strGitHub & "Templates/" & strTemplate
      oXMLHttpDoc.Open "GET", strURL, False
      oXMLHttpDoc.Send
      If oXMLHttpDoc.Status = 200 And oXMLHttpDoc.readyState = 4 Then
        'success: save it to templates directory
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'adTypeBinary
        oStream.Write oXMLHttpDoc.responseBody
        oStream.SaveToFile strTemplateDir & "\" & strTemplate
        oStream.Close
      Else
        cptBackbone_frm.lblStatus.Caption = "Download failed."
        'fail: prompt to request by email
        If MsgBox("Unable to download template. Request via email?", vbExclamation + vbYesNo, "No Connection") = vbYes Then
          On Error Resume Next
          Set oOutlook = GetObject(, "Outlook.Application")
          If oOutlook Is Nothing Then
            Set oOutlook = CreateObject("Outlook.Application")
          End If
          Set oMailItem = oOutlook.CreateItem(0) '0 = olMailItem
          oMailItem.To = "cpt@ClearPlanConsulting.com"
          oMailItem.Importance = 2 'olImportanceHigh
          oMailItem.Subject = "Template Request: " & strTemplate
          oMailItem.HTMLBody = "Please forward the subject-referenced template. Thank you." & oMailItem.HTMLBody
          oMailItem.Display False
        End If
        GoTo exit_here
      End If
      
    End If
  
    'open excel and create template
    Set oExcel = CreateObject("Excel.Application")
    Set oWorkbook = oExcel.Workbooks.Add(strTemplateDir & "\" & strTemplate)
    oExcel.Calculation = -4135 'xlManual
    oExcel.ScreenUpdating = False
    Set wsIndex = oWorkbook.Sheets("CWBS Index")
    wsIndex.Outline.SummaryRow = 0 'xlSummaryAbove
    Set wsDictionary = oWorkbook.Sheets("CWBS Dictionary")
    wsDictionary.Outline.SummaryRow = 0 'xlSummaryAbove
    lngRow = 7
    For lngItem = 1 To oLookupTable.Count
      'index: code=col1; name=col9
      wsIndex.Cells(lngRow, 1).Value = "'" & oLookupTable.Item(lngItem).FullName
      wsIndex.Cells(lngRow, 10).Value = oLookupTable.Item(lngItem).Description
      wsIndex.Cells(lngRow, 10).HorizontalAlignment = -4131 'xlLeft
      wsIndex.Cells(lngRow, 10).IndentLevel = Len(CStr(oLookupTable.Item(lngItem).FullName)) - Len(Replace(CStr(oLookupTable.Item(lngItem).FullName), ".", ""))
      wsIndex.Rows(lngRow).OutlineLevel = Len(CStr(oLookupTable.Item(lngItem).FullName)) - Len(Replace(CStr(oLookupTable.Item(lngItem).FullName), ".", "")) + 1
      If lngRow >= 8 Then
        wsIndex.Range(wsIndex.Cells(lngRow, 10), wsIndex.Cells(lngRow, 19)).Merge
      End If
      'dictionary: code=col1; name=col2
      wsDictionary.Cells(lngRow, 1).Value = "'" & oLookupTable.Item(lngItem).FullName
      wsDictionary.Cells(lngRow, 2).Value = oLookupTable.Item(lngItem).Description
      wsDictionary.Cells(lngRow, 2).HorizontalAlignment = -4131 'xlLeft
      wsDictionary.Cells(lngRow, 2).IndentLevel = wsIndex.Cells(lngRow, 10).IndentLevel
      wsDictionary.Rows(lngRow).OutlineLevel = wsIndex.Rows(lngRow).OutlineLevel
      If lngRow >= 8 Then
        wsDictionary.Range(wsDictionary.Cells(lngRow, 2), wsDictionary.Cells(lngRow, 3)).Merge
        wsDictionary.Range(wsDictionary.Cells(lngRow, 4), wsDictionary.Cells(lngRow, 11)).Merge
      End If
      cptBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngItem / oLookupTable.Count, "0%") & ")"
      cptBackbone_frm.lblProgress.Width = (lngItem / oLookupTable.Count) * cptBackbone_frm.lblStatus.Width
      lngRow = lngRow + 1
    Next
  End If
  
  'format it
  '-4121=-4121; -4161=xlToRight; 1=xlContinuous; 2=xlThin; -4105=xlColorIndexAutomatic
  wsIndex.[B8:I8].AutoFill Destination:=wsIndex.Range(wsIndex.Cells(8, 2), wsIndex.Cells(7 + oLookupTable.Count - 1, 9))
  For lngBorder = 7 To 12 'left,top,bottom,right,insidevertical,insidehorizontal
    With wsIndex.Range(wsIndex.[A7].End(-4121), wsIndex.Cells(7, 19)).Borders(lngBorder)
      .LineStyle = 1
      .Weight = 2
      .ColorIndex = -4105
    End With
    With wsDictionary.Range(wsDictionary.[A7].End(-4121), wsDictionary.Cells(7, 11)).Borders(lngBorder)
      .LineStyle = 1
      .Weight = 2
      .ColorIndex = -4105
    End With
  Next lngBorder
  wsDictionary.Range(wsDictionary.[A7].End(-4121), wsDictionary.[A7].End(-4161)).BorderAround 1, 2, -4105
  
  'freeze panes
  wsDictionary.Activate
  wsDictionary.[A7].Select
  oExcel.ActiveWindow.FreezePanes = True
  wsIndex.Activate
  wsIndex.[A7].Select
  oExcel.ActiveWindow.FreezePanes = True
  oExcel.Visible = True
  
  'provide user feedback
  cptBackbone_frm.lblStatus.Caption = "Complete."
  
exit_here:
  On Error Resume Next
  Set oMailItem = Nothing
  Set oExcel = Nothing
  cptBackbone_frm.lblStatus.Caption = "Ready..."
  cptBackbone_frm.lblProgress.Width = cptBackbone_frm.lblStatus.Width
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  Set wsDictionary = Nothing
  Set wsIndex = Nothing
  Set oWorkbook = Nothing
  oExcel.Calculation = -4105 'xlAutomatic
  oExcel.ScreenUpdating = True
  Set oExcel = Nothing
  Set oStream = Nothing
  Set oXMLHttpDoc = Nothing
  Set oShell = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptExport81334D", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportTemplate()
  'objects
  Dim oWorksheet As Object
  Dim oWorkbook As Object
  Dim oExcel As Object
  'strings
  Dim strMsg As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strMsg = "Instructions:" & vbCrLf
  strMsg = strMsg & "1. Do not add, edit, move, or remove columns." & vbCrLf
  strMsg = strMsg & "2. No empty rows from row 2 to the end of your Code." & vbCrLf
  strMsg = strMsg & "3. Save and import when done." & vbCrLf & vbCrLf
  strMsg = strMsg & "- CWBS SUGGESTION: Include down to Control Account levels, suffixed with ' CA'" & vbCrLf
  strMsg = strMsg & "- IMP SUGGESTION: Include down to an accomplishment criteria milestone." & vbCrLf & vbCrLf
  strMsg = strMsg & "Proceed?"
  If MsgBox(strMsg, vbInformation + vbYesNo, "Instructions:") = vbYes Then
    Set oExcel = CreateObject("Excel.Application")
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = "CWBS"
    oWorksheet.[A1:C1] = Array("CODE", "LEVEL", "DESCRIPTION")
    oWorksheet.[A1:C1].Font.Bold = True
    oWorksheet.[A2].Select
    oWorksheet.Columns(1).ColumnWidth = 10
    oWorksheet.Columns(2).ColumnWidth = 5.2
    oWorksheet.Columns(3).ColumnWidth = 59.14
    oExcel.ActiveWindow.FreezePanes = True
    oExcel.ActiveWindow.Zoom = 85
    oExcel.Visible = True
    Application.ActivateMicrosoftApp pjMicrosoftExcel
  End If
  
exit_here:
  On Error Resume Next
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptExportTemplate", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowBackbone_frm()
  'longs
  Dim lngCode As Long, lngOutlineCode As Long
  'strings
  Dim strOutlineCode As String, strOutlineCodeName As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptBackbone_frm.cboOutlineCodes
    .Clear
    'populate the listbox/combobox
    For lngCode = 1 To 10
      strOutlineCode = "Outline Code" & lngCode
      lngOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
      strOutlineCodeName = Application.CustomFieldGetName(lngOutlineCode)
      .AddItem
      If Len(strOutlineCodeName) > 0 Then
        strOutlineCode = strOutlineCode & " (" & strOutlineCodeName & ")"
      End If
      .List(lngCode - 1, 0) = lngOutlineCode
      .List(lngCode - 1, 1) = strOutlineCode
    Next lngCode
  End With
  
  'add Import Actions
  With cptBackbone_frm.cboImport
    .Clear
    .AddItem "From Excel Workbook"
    .AddItem "From MSP Server Outline Code Export"
    .AddItem "From MIL-STD-881D Appendix B"
    .AddItem "From MIL-STD-881D Appendix E"
    .AddItem "From Existing Tasks"
  End With
  
  'add Export Actions
  With cptBackbone_frm.cboExport
    .Clear
    .AddItem "To Excel Workbook"
    .AddItem "To CSV for MPM"
    .AddItem "To CSV for COBRA"
    .AddItem "To DI-MGMT-81334D Template"
  End With
  
  'pre-select Outline Code 1
  With cptBackbone_frm
    .cboOutlineCodes.ListIndex = 0
    .txtNameIt = CustomFieldGetName(.cboOutlineCodes.List(0, 0))
    .Caption = "Backbone (" & cptGetVersion("cptBackbone_frm") & ")"
    .cboOutlineCodes.SetFocus
   Call cptBackboneHideControls
   .Show 'False
  End With

exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptShowBackbone_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptCreateCode(lngOutlineCode As Long)
  'objects
  Dim objOutlineCode As OutlineCode
  Dim objLookupTable As LookupTable
  Dim objLookupTableEntry As LookupTableEntry
  Dim oTask As Task
  'strings
  Dim strWBS As String, strParent As String, strChild As String
  'longs
  Dim lngUID As Long, lngTasks As Long, lngTask As Long, lngLevel As Long
  'variants
  Dim aOutlineCode As Variant, tmr As Date

  tmr = Now
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure name doesn't already exist - trust form formatting
  If cptBackbone_frm.txtNameIt.BorderColor = 255 Then GoTo exit_here

  'first name the field and create the code mask
  For lngLevel = 1 To 10
    CustomOutlineCodeEditEx lngOutlineCode, Level:=lngLevel, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  Next lngLevel
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
  Set objOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode))
  Set objLookupTable = objOutlineCode.LookupTable
  
  lngTasks = ActiveProject.Tasks.Count
  
  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then
      lngTask = lngTask + 1
      oTask.SetField lngOutlineCode, oTask.WBS
      objLookupTable.Item(lngTask).Description = oTask.Name
      cptBackbone_frm.lblProgress.Width = ((lngTask - 1) / lngTasks) * cptBackbone_frm.lblStatus.Width
      cptBackbone_frm.lblStatus.Caption = Format(lngTask - 1, "#,##0") & " / " & Format(lngTasks, "#,##0") & " (" & Format((lngTask - 1) / lngTasks, "0%") & ") [" & Format(Now - tmr, "hh:nn:ss") & "]"
    End If 'task is nothing
  Next oTask
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLeaves:=True, OnlyLookUpTableCodes:=True
  cptBackbone_frm.lblStatus.Caption = "Complete."
  Application.StatusBar = "Complete."
  cptBackbone_frm.cmdCancel.Caption = "Done"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  cptSpeed False
  Set objOutlineCode = Nothing
  Set objLookupTable = Nothing
  Set objLookupTableEntry = Nothing
  Set oTask = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptCreateCode", Err, Erl)
  Resume exit_here
End Sub

Sub cptRenameInsideOutlineCode(strOutlineCode As String, strFind As String, strReplace As String)
  'usage: Call RenameOutlineCode("CWBS","BOSS","IBRS")
  'objects
  Dim oOutlineCode As OutlineCode, oLookupTable As LookupTable, oLookupTableEntry As LookupTableEntry
  'longs
  Dim lngEntry As Long
  Dim lngReplaced As Long
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  Set oLookupTable = oOutlineCode.LookupTable
  For lngEntry = 1 To oLookupTable.Count
    If InStr(oLookupTable(lngEntry).Description, strFind) > 0 Then
      oLookupTable(lngEntry).Description = Replace(oLookupTable(lngEntry).Description, strFind, strReplace)
      lngReplaced = lngReplaced + 1
    End If
  Next lngEntry
  
  cptBackbone_frm.lblFeedback.Caption = Format(lngReplaced, "#,##0") & " replaced"
  
exit_here:
  On Error Resume Next
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptRenameInsideOutlineCode", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshOutlineCodePreview(strOutlineCode As String)
  'objects
  Dim oOutlineCode As OutlineCode, oLookupTable As LookupTable, oLookupTableEntry As LookupTableEntry
  Dim oNode As Object 'Node
  'strings
  'longs
  Dim lngEntry As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  strOutlineCode = Replace(Replace(strOutlineCode, cptRegEx(strOutlineCode, "Outline Code[0-9]{1,}") & " (", ""), ")", "")
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oLookupTable Is Nothing Then
    If oLookupTable.Count > 0 Then
      cptBackbone_frm.lboOutlineCode.Clear
      For lngEntry = 1 To oLookupTable.Count
        With cptBackbone_frm.lboOutlineCode
          .AddItem
          .List(.ListCount - 1, 0) = oLookupTable(lngEntry).UniqueID
          .List(.ListCount - 1, 1) = oLookupTable(lngEntry).Level
          .List(.ListCount - 1, 2) = oLookupTable(lngEntry).FullName & " - " & oLookupTable(lngEntry).Description
          Application.StatusBar = "Adding: " & oLookupTable(lngEntry).FullName & " - " & oLookupTable(lngEntry).Description
        End With
      Next lngEntry
    End If 'lookuptable.count > 0
  End If 'lookuptable is nothing
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Set oLookupTableEntry = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptRefreshOutlineCodePreview", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptExportOutlineCodeForMPM(lngOutlineCode As Long)
  'exports local Outline Code to CSV for MPM Upload
  'objects
  Dim oOutlineCode As OutlineCode
  Dim oLookupTable As LookupTable
  'longs
  Dim lngItem As Long, lngFile As Long
  'strings
  Dim strHeader As String
  Dim strMsg As String
  Dim strCode As String, strDescription As String, strParent As String
  Dim strDir As String, strFile As String, strOutlineCode As String
  'booleans
  Dim blnCA As Boolean

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'confirm lookuptable exists
  Set oOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode))
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbExclamation + vbOKOnly, "No LookupTable"
    GoTo exit_here
  End If
  
  'set directory
  strDir = Environ("TEMP") & "\"
  strFile = "WBS_DESCRIPTIVE_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".csv"
  If Dir(strDir & strFile) <> vbNullString Then Kill strDir & strFile

  lngFile = FreeFile
  Open strDir & strFile For Output As #lngFile
  
  If cptBackbone_frm.chkIncludeHeaders Then
    strHeader = "WBS ID,"
    strHeader = strHeader & "WBS Description,"
    strHeader = strHeader & "Alias,"
    strHeader = strHeader & "XREF1,"
    strHeader = strHeader & "XREF2,"
    strHeader = strHeader & "XREF3,"
    strHeader = strHeader & "XREF4,"
    strHeader = strHeader & "XREF5,"
    strHeader = strHeader & "XREF6,"
    strHeader = strHeader & "XREF7,"
    strHeader = strHeader & "XREF8,"
    strHeader = strHeader & "XREF9,"
    strHeader = strHeader & "XREF10,"
    strHeader = strHeader & "Manager,"
    strHeader = strHeader & "Charge Number,"
    strHeader = strHeader & "Performing Department,"
    strHeader = strHeader & "Responsible Department,"
    strHeader = strHeader & "Element Type,"
    strHeader = strHeader & "Earned Value Method,"
    strHeader = strHeader & "CLIN,"
    strHeader = strHeader & "Recurring or non-recurring,"
    strHeader = strHeader & "Fee %,"
    strHeader = strHeader & "Fee Limit Amount,"
    strHeader = strHeader & "BCWP Base Unit,"
    strHeader = strHeader & "Parent WBS ID,"
    strHeader = strHeader & "Base WBS,"
    Print #lngFile, strHeader
  End If
  
  'output top level
  Print #lngFile, "*" & "," & Chr(34) & ActiveProject.Name & Chr(34) & String(25, ",")
  For lngItem = 1 To oLookupTable.Count
    strCode = oLookupTable(lngItem).FullName
    strDescription = oLookupTable(lngItem).Description
    If Not oLookupTable(lngItem).IsValid Then
      MsgBox "Invalid Code Found! See " & strCode & " : " & strDescription, vbCritical + vbOKOnly, "Error"
      GoTo kill_file
    End If
    blnCA = Right(strDescription, 3) = " CA"
    If Len(strCode) = 1 Then
      strParent = "*"
    Else
      strParent = Left(strCode, InStrRev(strCode, ".") - 1)
    End If
    cptBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngItem / oLookupTable.Count, "0%") & ")"
    cptBackbone_frm.lblProgress.Width = (lngItem / oLookupTable.Count) * cptBackbone_frm.lblStatus.Width
    Print #lngFile, strCode & "," & Chr(34) & strDescription & Chr(34) & String(16, ",") & IIf(blnCA, "C", "") & String(7, ",") & strParent & ",,"
  Next lngItem
  
  Close #lngFile
  
  'open it in notepad
  Shell "C:\Windows\notepad.exe '" & strDir & strFile & "'", vbNormalFocus
  
exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  For lngFile = 1 To FreeFile: Close #lngFile: Next lngFile
  Exit Sub
  
kill_file:
  On Error Resume Next
  Close #lngFile
  Kill strDir & strFile
  Resume exit_here
  
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptExportOutlineCodeForMPM", Err, Erl)
  Resume exit_here

End Sub

Sub cptBackboneHideControls()

  With cptBackbone_frm
    'Replace
    .lblFeedback.Visible = .optReplace
    .txtReplace.Enabled = .optReplace
    .txtReplacement.Enabled = .optReplace
    .cmdReplace.Enabled = .optReplace
    'Import
    .txtNameIt.Enabled = .optImport
    .cboImport.Enabled = .optImport
    .chkAlsoCreateTasks.Enabled = .optImport
    .cmdExportTemplate.Visible = False
    .cmdImport.Enabled = .optImport
    'Export
    .cboExport.Enabled = .optExport
    .chkIncludeHeaders.Enabled = .optExport
    .chkIncludeThresholds.Enabled = .optExport
    .cmdExport.Enabled = .optExport
  End With

End Sub

Sub cptExportOutlineCodeForCOBRA(lngOutlineCode)
  'objects
  Dim oLookupTable As LookupTable
  Dim oOutlineCode As OutlineCode
  'strings
  Dim strOutlineCode As String
  Dim strDescription As String
  Dim strCode As String
  Dim strFile As String
  Dim strHeader As String
  'longs
  Dim lngItem As Long
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm lookuptable exists
  Set oOutlineCode = ActiveProject.OutlineCodes(CustomFieldGetName(lngOutlineCode))
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then
    strOutlineCode = CustomFieldGetName(lngOutlineCode)
    MsgBox "There is no LookupTable associated with " & FieldConstantToFieldName(lngOutlineCode) & IIf(Len(strOutlineCode) > 0, " (" & strOutlineCode & ")", "") & ".", vbExclamation + vbOKOnly, "No LookupTable"
    GoTo exit_here
  End If
  
  'setup the export file
  strFile = Environ("TEMP") & "\CODE_FILE_WBS.csv"
  If Dir(strFile) <> vbNullString Then Kill strFile
  lngFile = FreeFile
  Open strFile For Output As #lngFile
  
  'export header
  strHeader = "Code,"
  strHeader = strHeader & "Description,"
  If cptBackbone_frm.chkIncludeThresholds Then
    strHeader = strHeader & "Threshold SV Value Current Period Favorable,"
    strHeader = strHeader & "Threshold SV Value Current Period Unfavorable,"
    strHeader = strHeader & "Threshold SV % Current Period Favorable,"
    strHeader = strHeader & "Threshold SV % Current Period Unfavorable,"
    strHeader = strHeader & "Threshold SV Value Cumulative Favorable,"
    strHeader = strHeader & "Threshold SV Value Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold SV % Cumulative Favorable,"
    strHeader = strHeader & "Threshold SV % Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold CV Value Current Period Favorable,"
    strHeader = strHeader & "Threshold CV Value Current Period Unfavorable,"
    strHeader = strHeader & "Threshold CV % Current Period Favorable,"
    strHeader = strHeader & "Threshold CV % Current Period Unfavorable,"
    strHeader = strHeader & "Threshold CV Value Cumulative Favorable,"
    strHeader = strHeader & "Threshold CV Value Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold CV % Cumulative Favorable,"
    strHeader = strHeader & "Threshold CV % Cumulative Unfavorable,"
    strHeader = strHeader & "Threshold At Complete Value Favorable,"
    strHeader = strHeader & "Threshold At Complete Value Unfavorable,"
    strHeader = strHeader & "Threshold At Complete % Favorable,"
    strHeader = strHeader & "Threshold At Complete % Unfavorable,"
  End If
  
  Print #lngFile, strHeader
  
  'export outline code
  For lngItem = 1 To oLookupTable.Count
    strCode = oLookupTable(lngItem).FullName
    strDescription = oLookupTable(lngItem).Description
    If Not oLookupTable(lngItem).IsValid Then
      MsgBox "Invalid Code Found! See " & strCode & " : " & strDescription, vbCritical + vbOKOnly, "Error"
      GoTo kill_file
    End If
    cptBackbone_frm.lblStatus.Caption = "Exporting...(" & Format(lngItem / oLookupTable.Count, "0%") & ")"
    cptBackbone_frm.lblProgress.Width = (lngItem / oLookupTable.Count) * cptBackbone_frm.lblStatus.Width
    Print #lngFile, strCode & "," & Chr(34) & strDescription & Chr(34) & ","
  Next lngItem

  Close #lngFile
  
  Shell "C:\Windows\notepad.exe '" & strFile & "'", vbNormalFocus

exit_here:
  On Error Resume Next
  Set oLookupTable = Nothing
  Set oOutlineCode = Nothing
  For lngFile = 1 To FreeFile: Close #lngFile: Next lngFile
  Exit Sub
  
kill_file:
  On Error Resume Next
  Close #lngFile
  Kill strFile
  Resume exit_here
  
err_here:
  Call cptHandleErr("cptBackbone_bas", "cptExportOutlineCodeForCOBRA", Err, Erl)
  Resume exit_here
  
End Sub
