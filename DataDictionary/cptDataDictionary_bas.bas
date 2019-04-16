Attribute VB_Name = "cptDataDictionary_bas"
'<cpt_version>0.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptExportDataDictionary()
'objects
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Range
'strings
Dim strAttributes As String
Dim strFieldName As String
'longs
Dim lngCol As Long
Dim lngMax As Long
Dim lngHeaderRow As Long
Dim lngRow  As Long
Dim lngField As Long
'integers
Dim intListItem As Integer
Dim intField As Integer
'doubles
'booleans
'variants
Dim arrColumns As Variant
Dim vFieldType As Variant
Dim vFieldScope As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'set up a workbook/worksheet
  Application.StatusBar = "Creating Excel Workbook..."
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Worksheets(1)
  Worksheet.Name = "Data Dictionary"
  'IMS Title
  Worksheet.[A1].Value = "IMS Data Dictionary"
  Worksheet.[A1].Font.Size = 18
  Worksheet.[A1].Font.Bold = True
  'subtitle
  Worksheet.[A2].Value = ActiveProject.Name
  Worksheet.[A2].Font.Size = 14
  Worksheet.[A2].Font.Bold = True
  'date
  Worksheet.[A3].Value = FormatDateTime(Now, vbLongDate)
  
  'set header row
  lngHeaderRow = 5
  
  'set up columns
  arrColumns = Array("Enterprise", "Scope", "Type", "Field", "Custom Name", "Attributes", "Description")
  Worksheet.Range(Worksheet.Cells(lngHeaderRow, 1), Worksheet.Cells(lngHeaderRow, 1).Offset(0, UBound(arrColumns))) = arrColumns
  
  'freezepanes
  Worksheet.Cells(lngHeaderRow + 1, 1).Select
  xlApp.ActiveWindow.FreezePanes = True
  xlApp.ActiveWindow.Zoom = 85

  Application.StatusBar = "Exporting local custom fields..."

  'prep for data dump
  lngRow = lngHeaderRow
  'export local custom fields
  For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
    For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Number", "Start", "Text", "Outline Code")
      'avoid the errors
      Select Case vFieldType
        Case "Text"
          lngMax = 30
        Case "Flag"
          lngMax = 20
        Case "Number"
          lngMax = 20
        Case Else
          lngMax = 10
      End Select
      For intField = 1 To lngMax
        On Error GoTo err_here
        lngField = FieldNameToFieldConstant(vFieldType & intField, vFieldScope)
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          lngRow = lngRow + 1
          'xlApp.ActiveWindow.ScrollRow = lngRow - 1
          Worksheet.Cells(lngRow, 1).Value = False
          Worksheet.Cells(lngRow, 2).Value = Choose(CInt(vFieldScope) + 1, "Task", "Resource", "Project")
          Worksheet.Cells(lngRow, 3).Value = CStr(vFieldType)
          Worksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
          Worksheet.Cells(lngRow, 5).Value = strFieldName
          'get attributes
          If Len(CustomFieldGetFormula(lngField)) > 0 Then
            Worksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
          End If
          strAttributes = ""
          On Error Resume Next
          For intListItem = 1 To 1000
            If vFieldType = "Outline Code" Then
              strAttributes = strAttributes & vbCrLf & "- " & Application.CustomFieldValueListGetItem(lngField, pjValueListDescription, intListItem)
            Else
              strAttributes = strAttributes & vbCrLf & "- " & Application.CustomFieldValueListGetItem(lngField, pjValueListValue, intListItem) + " (" + Application.CustomFieldValueListGetItem(lngField, pjValueListDescription, intListItem) + ")"
            End If
            If err > 0 Then GoTo exit_for
          Next intListItem
exit_for:
          If Len(strAttributes) > 0 Then Worksheet.Cells(lngRow, 6).Value = "Lookup Values:" & strAttributes
        End If
next_field:
      Next intField
    Next vFieldType
  Next vFieldScope
  
  Application.StatusBar = "Exporting Enterprise Custom Fields..."
  
  'get enterprise custom fields
  For lngField = 188776000 To 188778000
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      lngRow = lngRow + 1
      'xlApp.ActiveWindow.ScrollRow = lngRow - 1
      Worksheet.Cells(lngRow, 1).Value = True
      Worksheet.Cells(lngRow, 2).Value = "n/a"
      Worksheet.Cells(lngRow, 3).Value = "n/a"
      Worksheet.Cells(lngRow, 4).Value = "n/a"
      Worksheet.Cells(lngRow, 5).Value = FieldConstantToFieldName(lngField)
      'field attributes like formulae and pick lists not exposed to VBA
    End If
  Next lngField
  
  Application.StatusBar = "Formatting..."
  
  'make it nice
  'convert to table / format it
  xlApp.ActiveWindow.ScrollRow = lngHeaderRow
  Set rng = Worksheet.Range(Worksheet.Cells(lngHeaderRow, 1).End(xlToRight), Worksheet.Cells(lngHeaderRow, 1).End(xlDown))
  Worksheet.ListObjects.Add(xlSrcRange, rng, , xlYes).Name = "DATA_DICTIONARY"
  'autofit
  Worksheet.Range("DATA_DICTIONARY[#All]").Select
  rng.Columns.AutoFit
  rng.Rows.AutoFit
  rng.VerticalAlignment = xlCenter
  lngCol = Worksheet.Rows(lngHeaderRow).Find("Attributes", lookat:=xlWhole).Column
  Worksheet.Columns(lngCol).ColumnWidth = 100
  Worksheet.Columns(lngCol).WrapText = True
  lngCol = Worksheet.Rows(lngHeaderRow).Find("Description", lookat:=xlWhole).Column
  Worksheet.Columns(lngCol).ColumnWidth = 100
  Worksheet.Cells(lngHeaderRow + 1, 1).Select
  
  Application.StatusBar = "Complete."
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  If Not xlApp Is Nothing Then xlApp.Visible = True
  Set rng = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing

  Exit Sub
  
err_here:
  If err.Number = 1101 Or err.Number = 1004 Then
    err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptExportCustomFields_bas", "cptExportDataDictionary", err, Erl)
  End If
  
End Sub
