Attribute VB_Name = "cptBackbone_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptImportAppendixB()
'objects
Dim TaskTable As Object
Dim Task As Object
'strings
'longs
Dim lngItem As Long
Dim lngField As Long
Dim lngOutlineLevel As Long
'integers
'doubles
'booleans
Dim blnOutlineCode As Boolean
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnOutlineCode = MsgBox("Populate Outline Code 1?", vbQuestion + vbYesNo) = vbYes

  If blnOutlineCode Then
    Application.CustomFieldRename pjCustomTaskOutlineCode1, "CWBS"
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=2, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=3, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=4, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=5, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=6, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=7, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=8, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=9, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, Level:=10, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, OnlyLookUpTableCodes:=False, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
  End If

  With CreateObject("ADODB.Recordset")
    If Dir(cptDir & "\cwbs-appendix-b.adtg") = vbNullString Then
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
    Else
      .Open cptDir & "\cwbs-appendix-b.adtg"
    End If
    .MoveFirst
    lngItem = 0
    Do While Not .EOF
      lngItem = lngItem + 1
      Set Task = ActiveProject.Tasks.Add(.Fields(1).Value)
      EditGoTo Task.ID
      Task.WBS = .Fields(0).Value
      If blnOutlineCode Then
        Task.SetField pjCustomTaskOutlineCode1, .Fields(0)
        ActiveProject.OutlineCodes("CWBS").LookupTable.Item(lngItem).Description = .Fields(1).Value
      End If
      lngOutlineLevel = Len(.Fields(0).Value) - Len(Replace(.Fields(0).Value, ".", ""))
      If lngOutlineLevel > 0 Then
        Task.OutlineLevel = lngOutlineLevel + 1
      End If
      
      .MoveNext
    Loop
    .Close
  End With
  SelectBeginning
  SetRowHeight 1, "all"
  
  Set TaskTable = ActiveProject.TaskTables(ActiveProject.CurrentTable)
  For lngField = 1 To TaskTable.TableFields.Count
    If FieldConstantToFieldName(TaskTable.TableFields(lngField).Field) = "Name" Then Exit For
  Next lngField
  ColumnBestFit lngField
  
  'reset outline code to disallow new entries
  If blnOutlineCode Then
    CustomOutlineCodeEditEx FieldID:=pjCustomTaskOutlineCode1, OnlyLookUpTableCodes:=True, OnlyLeaves:=True, LookupDefault:=False, SortOrder:=0
  End If
  
exit_here:
  On Error Resume Next
  Set TaskTable = Nothing
  Set Task = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptWBS_bas", "cptImportAppendixB", err)
  Resume exit_here
End Sub


Sub cptExportOutlineCode()
'objects
Dim xlApp As Object 'Excel.Application
Dim Workbook As Object 'Workbook
Dim Worksheet As Object 'Worksheet
Dim ListObject As Object 'ListObject
Dim OutlineCode As Object 'OutlineCode
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

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  '====== UPDATE FIELD NAME =======
  strOutlineCode = "CWBS v2-2"
  '================================
  
  Application.StatusBar = "Exporting Outline Code '" & strOutlineCode & "'..."
  
  'get excel
  Application.StatusBar = "Setting up Excel..."
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  xlApp.Calculation = -4135 'xlCalculationManual
  xlApp.ScreenUpdating = False
  Set Worksheet = Workbook.Sheets(1)
  Worksheet.[A1:C1] = Array("CODE", "LEVEL", "DESCRIPTION")
  
  'export the codes
  Set OutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  For lngLookupItems = 1 To OutlineCode.LookupTable.Count
    lngLastRow = Worksheet.Cells(Worksheet.Rows.Count, 1).End(xlUp).Row + 1
    Worksheet.Cells(lngLastRow, 1).Value = "'" & OutlineCode.LookupTable.Item(lngLookupItems).FullName
    Worksheet.Cells(lngLastRow, 2).Value = OutlineCode.LookupTable.Item(lngLookupItems).Level
    Worksheet.Cells(lngLastRow, 3).Value = OutlineCode.LookupTable.Item(lngLookupItems).Description
    Worksheet.Cells(lngLastRow, 3).IndentLevel = OutlineCode.LookupTable.Item(lngLookupItems).Level - 1
    Application.StatusBar = "Exporting Outline Code '" & strOutlineCode & "'...(" & Format((lngLastRow - 1) / lngLookupItems, "0%") & ")"
  Next lngLookupItems
  
  Application.StatusBar = "Formatting Worksheet..."
  
  'format the table
  xlApp.ActiveWindow.Zoom = 85
  'Set ListObject = Worksheet.ListObjects.Add(xlSrcRange, Worksheet.Range(Worksheet.[A1].End(xlToRight), Worksheet.[A1].End(xlDown)), , xlYes)
  Set ListObject = Worksheet.ListObjects.Add(1, Worksheet.Range(Worksheet.[A1].End(-4161), Worksheet.[A1].End(-4121)), , 1)
  ListObject.Name = strOutlineCode
  ListObject.TableStyle = ""
  ListObject.HeaderRowRange.Font.Bold = True
  ListObject.Range.Borders(xlDiagonalDown).LineStyle = -4142 'xlNone
  ListObject.Range.Borders(xlDiagonalUp).LineStyle = -4142 'xlNone
  With ListObject.Range.Borders(7) 'xlEdgeLeft
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With ListObject.Range.Borders(8) 'xlEdgeTop
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With ListObject.Range.Borders(9) 'xlEdgeBottom
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With ListObject.Range.Borders(10) 'xlEdgeRight
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.499984740745262
    .Weight = 2 'xlThin
  End With
  With ListObject.Range.Borders(11) 'xlInsideVertical
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = 2 'xlThin
  End With
  With ListObject.Range.Borders(12) 'xlInsideHorizontal
    .LineStyle = 1 'xlContinuous
    .ThemeColor = 1
    .TintAndShade = -0.249946592608417
    .Weight = 2 'xlThin
  End With
    With ListObject.HeaderRowRange.Interior
      .Pattern = 1 'xlSolid
      .PatternColorIndex = -4105 'xlAutomatic
      .ThemeColor = 1 'xlThemeColorDark1
      .TintAndShade = -0.149998474074526
      .PatternTintAndShade = 0
    End With
    Worksheet.[A2].Select
    xlApp.ActiveWindow.FreezePanes = True
    Worksheet.Columns.AutoFit
    
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  xlApp.Visible = True
  xlApp.ScreenUpdating = True
  xlApp.Calculation = xlCalculationAutomatic
  Set ListObject = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Set OutlineCode = Nothing

  Exit Sub
  
err_here:
  Call cptHandleErr("cptOutlineCodes", "ExportOutlineCode", err, Erl)
  Resume exit_here
  
End Sub

Sub cptCreate81334D()

End Sub
