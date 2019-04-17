Attribute VB_Name = "cptDataDictionary_bas"
'<cpt_version>0.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptExportDataDictionary()
'objects
Dim rst As Object
Dim xlApp As Excel.Application, Workbook As Workbook, Worksheet As Worksheet, rng As Range
'strings
Dim strGUID As String
Dim strAttributes As String
Dim strFieldName As String
'longs
Dim lngItem As Long
Dim lngItems As Long
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
Dim blnExists As Boolean
'variants
Dim arrColumns As Variant
Dim vFieldType As Variant
Dim vFieldScope As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'get project uid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'set up a workbook/worksheet
  cptDataDictionary_frm.lblStatus.Caption = "Creating Excel Workbook..."
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

  cptDataDictionary_frm.lblStatus.Caption = "Exporting local custom fields..."
  
  blnExists = Dir(cptDir & "\settings\data-dictionary.adtg") <> vbNullString

  If blnExists Then
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open cptDir & "\settings\data-dictionary.adtg"
  End If
  
  lngItems = 260 + (188778000 - 188776000)
  
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

        If blnExists Then
          rst.Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & lngField
          If Not rst.EOF Then Worksheet.Cells(lngRow, 7).Value = rst("DESCRIPTION")
          rst.Filter = ""
        End If
        
next_field:
        lngItem = lngItem + 1
        cptDataDictionary_frm.lblStatus.Caption = "Exporting Local Custom Fields..." & lngItem & "/" & lngItems & " (" & Format(lngItem / lngItems, "0%") & ")"
        cptDataDictionary_frm.lblProgress.Width = (lngItem / lngItems) * cptDataDictionary_frm.lblStatus.Width
      Next intField
    Next vFieldType
  Next vFieldScope
  
  cptDataDictionary_frm.lblStatus.Caption = "Exporting Enterprise Custom Fields..."
  
  'get enterprise custom fields
  For lngField = 188776000 To 188778000
    If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
      lngRow = lngRow + 1
      'xlApp.ActiveWindow.ScrollRow = lngRow - 1
      Worksheet.Cells(lngRow, 1).Value = True
      Worksheet.Cells(lngRow, 2).Value = "Enterprise"
      Worksheet.Cells(lngRow, 3).Value = "Enterprise"
      Worksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
      Worksheet.Cells(lngRow, 5).Value = FieldConstantToFieldName(lngField)
      'field attributes like formulae and pick lists not exposed to VBA
      If blnExists Then
        rst.Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & lngField
        If Not rst.EOF Then Worksheet.Cells(lngRow, 7).Value = rst("DESCRIPTION")
        rst.Filter = ""
      End If
    End If
    lngItem = lngItem + 1
    cptDataDictionary_frm.lblStatus.Caption = "Exporting Enterprise Custom Fields..." & lngItem & "/" & lngItems & " (" & Format(lngItem / lngItems, "0%") & ")"
    cptDataDictionary_frm.lblProgress.Width = (lngItem / lngItems) * cptDataDictionary_frm.lblStatus.Width
  Next lngField
    
  cptDataDictionary_frm.lblStatus.Caption = "Formatting..."
  
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
  
  cptDataDictionary_frm.lblStatus.Caption = "Opening..."
  
exit_here:
  On Error Resume Next
  If rst.State = 1 Then rst.Close
  Set rst = Nothing
  cptDataDictionary_frm.lblStatus.Caption = "Ready..."
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

Sub ShowFrmCptDataDictionary()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  Call cptRefreshDictionary
  cptDataDictionary_frm.txtFilter.SetFocus
  cptDataDictionary_frm.Show
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "ShowFrmCptDataDictionary()", err)
  Resume exit_here
End Sub

Sub cptRefreshDictionary()
'objects
Dim rst As Object 'ADODB.Recordset
'strings
Dim strFieldName As String
Dim strCustomName As String
Dim strGUID As String
'longs
Dim lngItem As Long
Dim lngField As Long
Dim lngMax As Long
'integers
Dim intField As Integer
'doubles
'booleans
Dim blnCreate As Boolean
Dim blnExists As Boolean
'variants
Dim vFieldType As Variant
Dim vFieldScope As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'clear the form if it's visible
  If cptDataDictionary_frm.Visible Then
    cptDataDictionary_frm.lboCustomFields.Clear
    cptDataDictionary_frm.txtDescription.Value = ""
  End If
  
  'get unique id of the current project
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'if data file exists then use it else create it
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    If Dir(cptDir & "\settings\data-dictionary.adtg") = vbNullString Then
      blnCreate = True
      .Fields.Append "PROJECT_ID", 200, 50 'adVarChar
      .Fields.Append "FIELD_ID", 3 'adInteger = Long
      .Fields.Append "FIELD_NAME", 200, 50
      .Fields.Append "CUSTOM_NAME", 200, 50
      .Fields.Append "DESCRIPTION", 203, 500 'adLongVarWChar
      .Open
    Else
      blnCreate = False
      .Open cptDir & "\settings\data-dictionary.adtg"
    End If
    
    'get local custom fields
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
          lngField = FieldNameToFieldConstant(vFieldType & intField, vFieldScope)
          strFieldName = FieldConstantToFieldName(lngField)
          strCustomName = CustomFieldGetName(lngField)
          If Len(strCustomName) > 0 Then
            If blnCreate Then
              'add to data store
              .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>")
            Else
              'does it exist?
              .Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & CLng(lngField)
              'if not then add it
              If .EOF Then
                .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>")
              End If
              .Filter = ""
            End If
          End If
        Next intField
      Next vFieldType
    Next vFieldScope
    
    'get enterprise custom fields
    For lngField = 188776000 To 188778000
      If Application.FieldConstantToFieldName(lngField) <> "<Unavailable>" Then
        strFieldName = "Enterprise"
        strCustomName = FieldConstantToFieldName(lngField)
        If blnCreate Then
          .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>")
        Else
          'does it exist?
          .Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & lngField
          'if not, then add it
          If .EOF Then
            .AddNew Array("PROJECT_ID", "FIELD_ID", "FIELD_NAME", "CUSTOM_NAME", "DESCRIPTION"), Array(strGUID, lngField, strFieldName, strCustomName, "<missing>")
          End If
          .Filter = ""
        End If
      End If
    Next lngField
    
    'save the data
    .Save cptDir & "\settings\data-dictionary.adtg"
    
    'populate the list
    If Not .EOF Then
      .Sort = "CUSTOM_NAME"
      .MoveFirst
      lngItem = 0
      Do While Not .EOF
        cptDataDictionary_frm.lboCustomFields.AddItem
        cptDataDictionary_frm.lboCustomFields.List(lngItem, 0) = .Fields("FIELD_ID")
        If .Fields("FIELD_ID") >= 188776000 Then
          cptDataDictionary_frm.lboCustomFields.List(lngItem, 1) = .Fields("CUSTOM_NAME") & " (Enterprise)"
        Else
          cptDataDictionary_frm.lboCustomFields.List(lngItem, 1) = .Fields("CUSTOM_NAME") & " (" & .Fields("FIELD_NAME") & ")"
        End If
        cptDataDictionary_frm.lboCustomFields.List(lngItem, 2) = .Fields("DESCRIPTION")
        .MoveNext
        lngItem = lngItem + 1
      Loop
      .Close
    End If
  End With
  
exit_here:
  On Error Resume Next
  If rst.State = 1 Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptRefreshDictionary", err)
  Resume exit_here
End Sub
