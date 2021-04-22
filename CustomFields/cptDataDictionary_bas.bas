Attribute VB_Name = "cptDataDictionary_bas"
'<cpt_version>v1.2.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptExportDataDictionary()
'objects
Dim oListObject As Object 'ListObject
Dim wsLookups As Object 'Worksheet
Dim dFields As Scripting.Dictionary 'Object
Dim oLookupTable As LookupTable
Dim rstDictionary As Object 'ADODB.Recordset
Dim oExcel As Object 'Excel.Application
Dim oWorkbook As Object 'Workbook
Dim oWorksheet As Object 'Worksheet
Dim oRange As Object 'Excel.Range
'strings
Dim strDescription As String
Dim strValue As String
Dim strGUID As String
Dim strAttributes As String
Dim strFieldName As String
'longs
Dim lngLookupCol As Long
Dim lngItem As Long
Dim lngItems As Long
Dim lngCol As Long
Dim lngHeaderRow As Long
Dim lngRow  As Long
Dim lngField As Long
'integers
Dim intListItem As Integer
Dim intField As Integer
'doubles
'booleans
Dim blnLookupTable As Boolean
Dim blnLookups As Boolean
Dim blnExists As Boolean
'variants
Dim vColumns As Variant
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
  
  blnLookups = MsgBox("Replicate Pick Lists in Excel?", vbQuestion + vbYesNo, "Data Dictionary") = vbYes
  
  'set up a workbook/worksheet
  cptDataDictionary_frm.lblStatus.Caption = "Creating Excel Workbook..."
  Set oExcel = CreateObject("Excel.Application")
  Set oWorkbook = oExcel.Workbooks.Add
  Set oWorksheet = oWorkbook.Worksheets(1)
  oWorksheet.Name = "Data Dictionary"
  
  'create lookups worksheet (optional)
  If blnLookups Then
    Set wsLookups = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(1))
    wsLookups.Name = "LOOKUPS"
    wsLookups.Activate
    oExcel.ActiveWindow.Zoom = 83
    oWorksheet.Activate
  End If
  
  'IMS Title
  oWorksheet.[A1].Value = "IMS Data Dictionary"
  oWorksheet.[A1].Font.Size = 18
  oWorksheet.[A1].Font.Bold = True
  'subtitle
  oWorksheet.[A2].Value = ActiveProject.Name
  oWorksheet.[A2].Font.Size = 14
  oWorksheet.[A2].Font.Bold = True
  'date
  oWorksheet.[A3].Value = FormatDateTime(Now, vbLongDate)
  
  'set header row
  lngHeaderRow = 5
  
  'set up columns
  vColumns = Array("Enterprise", "Scope", "Type", "Field", "Custom Name", "Attributes", "Description")
  oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1), oWorksheet.Cells(lngHeaderRow, 1).Offset(0, UBound(vColumns))) = vColumns
  
  'freezepanes
  oWorksheet.Cells(lngHeaderRow + 1, 1).Select
  oExcel.ActiveWindow.FreezePanes = True
  oExcel.ActiveWindow.Zoom = 85

  cptDataDictionary_frm.lblStatus.Caption = "Exporting local custom fields..."
  
  blnExists = Dir(cptDir & "\settings\cpt-data-dictionary.adtg") <> vbNullString

  If blnExists Then
    Set rstDictionary = CreateObject("ADODB.Recordset")
    rstDictionary.Open cptDir & "\settings\cpt-data-dictionary.adtg"
  End If
  
  'count of custom fields = local + enterprise
  lngItems = 260 + (188778000 - 188776000)
  
  Set dFields = CreateObject("Scripting.Dictionary")
  dFields.Add "Cost", 10
  dFields.Add "Date", 10
  dFields.Add "Duration", 10
  dFields.Add "Flag", 20
  dFields.Add "Finish", 10
  dFields.Add "Outline Code", 10
  dFields.Add "Number", 20
  dFields.Add "Start", 10
  dFields.Add "Text", 30
  
  'prep for data dump
  lngRow = lngHeaderRow
  'export local custom fields
  For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
    For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Outline Code", "Number", "Start", "Text")
      For intField = 1 To dFields.Item(vFieldType)
        lngField = FieldNameToFieldConstant(vFieldType & intField, vFieldScope)
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          lngRow = lngRow + 1
          oWorksheet.Cells(lngRow, 1).Value = False
          oWorksheet.Cells(lngRow, 2).Value = Choose(CInt(vFieldScope) + 1, "Task", "Resource", "Project")
          oWorksheet.Cells(lngRow, 3).Value = CStr(vFieldType)
          oWorksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
          oWorksheet.Cells(lngRow, 5).Value = strFieldName
          'get attributes
          If Len(CustomFieldGetFormula(lngField)) > 0 Then
            oWorksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
          End If
          blnLookupTable = False
          On Error Resume Next
          blnLookupTable = Len(CustomFieldValueListGetItem(lngField, pjValueListValue, 1)) > 0
          'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0 <- don't put this here
          
          strAttributes = ""
          If blnLookupTable Then
            
            If blnLookups Then
              lngLookupCol = wsLookups.[XFD2].End(-4159).Column
              If wsLookups.Cells(1, lngLookupCol) <> "" Then lngLookupCol = lngLookupCol + 2
              wsLookups.Cells(1, lngLookupCol) = UCase(strFieldName)
              wsLookups.Cells(2, lngLookupCol) = UCase(strFieldName) & " LOOKUP:"
            End If

            For intListItem = 1 To 1000
              If vFieldType = "Outline Code" Then
                Set oLookupTable = ActiveProject.OutlineCodes(CustomFieldGetName(lngField)).LookupTable
                If Len(oLookupTable(intListItem).Description) > 0 Then
                  If Left(oLookupTable(intListItem).Description, Len(oLookupTable(intListItem).FullName)) = oLookupTable(intListItem).FullName Then
                    strAttributes = strAttributes & vbCrLf & oLookupTable(intListItem).Description
                    If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = oLookupTable(intListItem).Description
                  Else
                    strAttributes = strAttributes & vbCrLf & oLookupTable(intListItem).FullName & " - " & oLookupTable(intListItem).Description
                    If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = oLookupTable(intListItem).FullName & " - " & oLookupTable(intListItem).Description
                  End If
                Else
                  strAttributes = strAttributes & vbCrLf & oLookupTable(intListItem).FullName
                  If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = oLookupTable(intListItem).FullName
                End If
              Else
                strValue = ""
                strValue = CustomFieldValueListGetItem(lngField, pjValueListValue, intListItem)
                strDescription = ""
                strDescription = CustomFieldValueListGetItem(lngField, pjValueListDescription, intListItem)
                If Len(strDescription) > 0 Then
                  strAttributes = strAttributes & vbCrLf & strValue & " - " & strDescription
                  If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = strValue & " - " & strDescription
                Else
                  strAttributes = strAttributes & vbCrLf & strValue
                  If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = strValue
                End If
              End If
              If Err > 0 Then
                Err.Clear
                Exit For
              End If
            Next intListItem
            
            If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
            
            If blnLookups Then 'use data validation
              'name the range
              wsLookups.ListObjects.Add(SourceType:=1, Source:=wsLookups.Range(wsLookups.Cells(1, lngLookupCol), wsLookups.Cells(2 + intListItem, lngLookupCol)).Address(True, True), xllistobjecthasheaders:=1).Name = UCase(Replace(FieldConstantToFieldName(lngField), " ", "_"))
              wsLookups.Columns(lngLookupCol).AutoFit
              wsLookups.Columns(lngLookupCol + 1).ColumnWidth = 2
              With oWorksheet.Cells(lngRow, 6).Validation
                 .Delete
                 .Add Type:=3, AlertStyle:=1, Operator:= _
                 1, Formula1:="=INDIRECT(""" & UCase(Replace(FieldConstantToFieldName(lngField), " ", "_")) & """)"
                 .IgnoreBlank = True
                 .InCellDropdown = True
                 .InputTitle = ""
                 .ErrorTitle = ""
                 .InputMessage = ""
                 .ErrorMessage = ""
                 .ShowInput = True
                 .ShowError = True
               End With
               oWorksheet.Cells(lngRow, 6).Value = UCase(strFieldName) & " LOOKUP:"
            Else 'don't
              If Len(strAttributes) > 0 Then oWorksheet.Cells(lngRow, 6).Value = "Lookup Values:" & strAttributes
            End If 'blnLookups
            
          End If 'Not LookupTable Is Nothing Then
          
        End If 'Len(strFieldName) > 0

        If blnExists Then
          rstDictionary.Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & lngField
          If Not rstDictionary.EOF Then oWorksheet.Cells(lngRow, 7).Value = rstDictionary("DESCRIPTION")
          rstDictionary.Filter = ""
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
      oWorksheet.Cells(lngRow, 1).Value = True
      oWorksheet.Cells(lngRow, 2).Value = "Enterprise"
      oWorksheet.Cells(lngRow, 3).Value = "Enterprise"
      oWorksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
      oWorksheet.Cells(lngRow, 5).Value = FieldConstantToFieldName(lngField)
      If Len(CustomFieldGetFormula(lngField)) > 0 Then
        oWorksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
      End If
      strAttributes = ""
      Set oLookupTable = Nothing
      On Error Resume Next
      Set oLookupTable = GlobalOutlineCodes(FieldConstantToFieldName(lngField)).LookupTable
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If Not oLookupTable Is Nothing Then
        If blnLookups Then
          lngLookupCol = wsLookups.[XFD2].End(-4159).Column
          If wsLookups.Cells(1, lngLookupCol) <> "" Then lngLookupCol = lngLookupCol + 2
          wsLookups.Cells(1, lngLookupCol) = UCase(FieldConstantToFieldName(lngField))
          wsLookups.Cells(2, lngLookupCol) = UCase(FieldConstantToFieldName(lngField)) & " LOOKUP:"
        End If
        For intListItem = 1 To oLookupTable.Count
          If Len(oLookupTable(intListItem).Description) > 0 Then
            If Left(oLookupTable(intListItem).Description, Len(oLookupTable(intListItem).FullName)) = oLookupTable(intListItem).FullName Then
              strAttributes = strAttributes & vbCrLf & oLookupTable(intListItem).Description
              If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = oLookupTable(intListItem).Description
            Else
              strAttributes = strAttributes & vbCrLf & oLookupTable(intListItem).FullName & " - " & oLookupTable(intListItem).Description
              If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = oLookupTable(intListItem).FullName & " - " & oLookupTable(intListItem).Description
            End If
          Else
            strAttributes = strAttributes & vbCrLf & oLookupTable(intListItem).FullName
            If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = oLookupTable(intListItem).FullName
          End If
        Next intListItem
        
        If blnLookups Then 'use validation
          'name the range
          wsLookups.ListObjects.Add(SourceType:=1, Source:=wsLookups.Range(wsLookups.Cells(1, lngLookupCol), wsLookups.Cells(2 + oLookupTable.Count, lngLookupCol)).Address(True, True), xllistobjecthasheaders:=1).Name = UCase(Replace(FieldConstantToFieldName(lngField), " ", "_"))
          wsLookups.Columns(lngLookupCol).AutoFit
          wsLookups.Columns(lngLookupCol + 1).ColumnWidth = 2
          With oWorksheet.Cells(lngRow, 6).Validation
            .Delete
            .Add Type:=3, AlertStyle:=1, Operator:= _
            1, Formula1:="=INDIRECT(""" & UCase(Replace(FieldConstantToFieldName(lngField), " ", "_")) & """)"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
          End With
          oWorksheet.Cells(lngRow, 6).Value = UCase(FieldConstantToFieldName(lngField)) & " LOOKUP:"
        Else
          If Len(strAttributes) > 0 Then oWorksheet.Cells(lngRow, 6).Value = "Lookup Values:" & strAttributes
        End If 'blnLookups
        
      End If 'Not LookupTable Is Nothing Then
      
      If blnExists Then
        rstDictionary.Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & lngField
        If Not rstDictionary.EOF Then oWorksheet.Cells(lngRow, 7).Value = rstDictionary("DESCRIPTION")
        rstDictionary.Filter = ""
      End If
      
    End If
    lngItem = lngItem + 1
    cptDataDictionary_frm.lblStatus.Caption = "Exporting Enterprise Custom Fields..." & lngItem & "/" & lngItems & " (" & Format(lngItem / lngItems, "0%") & ")"
    cptDataDictionary_frm.lblProgress.Width = (lngItem / lngItems) * cptDataDictionary_frm.lblStatus.Width
  Next lngField
    
  cptDataDictionary_frm.lblStatus.Caption = "Formatting..."
  
  'make it nice
  If blnLookups Then
    wsLookups.Activate
    oExcel.ActiveWindow.Zoom = 85
    wsLookups.[A2].Select
    oExcel.ActiveWindow.FreezePanes = True
    wsLookups.[A3].Select
    wsLookups.Rows(2).Hidden = True
  End If
  
  'convert to table / format it
  oWorksheet.Activate
  oExcel.ActiveWindow.ScrollRow = lngHeaderRow
  Set oRange = oWorksheet.Range(oWorksheet.Cells(lngHeaderRow, 1).End(-4161), oWorksheet.Cells(lngHeaderRow, 1).End(-4121))
  oWorksheet.ListObjects.Add(1, oRange, , 1).Name = "DATA_DICTIONARY"
  'autofit
  oWorksheet.Range("DATA_DICTIONARY[#All]").Select
  oRange.Columns.AutoFit
  oRange.Rows.AutoFit
  oRange.VerticalAlignment = xlCenter
  lngCol = oWorksheet.Rows(lngHeaderRow).Find("Attributes", lookat:=1).Column
  oWorksheet.Columns(lngCol).ColumnWidth = 100
  oWorksheet.Columns(lngCol).WrapText = True
  lngCol = oWorksheet.Rows(lngHeaderRow).Find("Description", lookat:=1).Column
  oWorksheet.Columns(lngCol).ColumnWidth = 100
  oWorksheet.Cells(lngHeaderRow + 1, 1).Select
  
  cptDataDictionary_frm.lblStatus.Caption = "Opening..."
  
exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  Set wsLookups = Nothing
  If rstDictionary.State Then rstDictionary.Close
  Set rstDictionary = Nothing
  Set oLookupTable = Nothing
  cptDataDictionary_frm.lblStatus.Caption = "Ready..."
  If Not oExcel Is Nothing Then oExcel.Visible = True
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
  
err_here:
  If Err.Number = 1101 Or Err.Number = 1004 Then
    Err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptExportCustomFields_bas", "cptExportDataDictionary", Err, Erl)
  End If
  
End Sub

Sub cptShowDataDictionary_frm()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  cptDataDictionary_frm.lboCustomFields.Clear
  Call cptRefreshDictionary
  cptDataDictionary_frm.txtFilter.SetFocus
  cptDataDictionary_frm.Show
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptShowDataDictionary_frm()", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshDictionary()
'objects
Dim dTypes As Object 'Scripting.Dictionary
Dim rstSaved As Object 'ADODB.Recordset
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
    cptDataDictionary_frm.txtFilter.Text = ""
    'cptDataDictionary_frm.txtDescription.Value = "" 'won't this erase an existing entry?
  End If
  
  'get unique id of the current project
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'if data file exists then use it else create it
  Set rstSaved = CreateObject("ADODB.Recordset")
  With rstSaved
    If Dir(cptDir & "\settings\cpt-data-dictionary.adtg") = vbNullString Then
      blnCreate = True
      .Fields.Append "PROJECT_ID", 200, 50 'adVarChar
      .Fields.Append "FIELD_ID", 3 'adInteger = Long
      .Fields.Append "FIELD_NAME", 200, 50
      .Fields.Append "CUSTOM_NAME", 200, 50
      .Fields.Append "DESCRIPTION", 203, 500 'adLongVarWChar
      .Open
    Else
      blnCreate = False
      .Open cptDir & "\settings\cpt-data-dictionary.adtg"
      .Filter = "PROJECT_ID='" & strGUID & "'"
    End If
    
    'get local custom fields
    'export local custom fields
    
    Set dTypes = CreateObject("Scripting.Dictionary")
    dTypes.Add "Cost", 10
    dTypes.Add "Date", 10
    dTypes.Add "Duration", 10
    dTypes.Add "Flag", 20
    dTypes.Add "Finish", 10
    dTypes.Add "Number", 20
    dTypes.Add "Start", 10
    dTypes.Add "Text", 30
    dTypes.Add "Outline Code", 10
        
    For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
      For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Number", "Start", "Text", "Outline Code")
        For intField = 1 To dTypes(vFieldType) 'lngMax
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
    .Save cptDir & "\settings\cpt-data-dictionary.adtg"
    
    'populate the list
    If Not .EOF Then
      .Filter = "PROJECT_ID='" & strGUID & "'"
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
  Set dTypes = Nothing
  If rstSaved.State = 1 Then rstSaved.Close
  Set rstSaved = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptRefreshDictionary", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportDataDictionary(Optional strFile As String)
'objects
Dim rstSaved As Object 'ADODB.Recordset
Dim oExcel As Object
Dim oWorkbook As Object
Dim oWorksheet As Object
Dim oRange As Object
'strings
Dim strMsg As String
Dim strNewDescription As String
Dim strCustomName As String
Dim strFieldsNotFound As String
Dim strGUID As String
Dim strSavedSettings As String
'longs
Dim lngNotFound As Long
Dim lngNew As Long
Dim lngTooLong As Long
Dim lngLastCol As Long
Dim lngRow As Long
Dim lngField As Long
Dim lngLastRow As Long
Dim lngDescriptionCol As Long
Dim lngNameCol As Long
Dim lngHeaderRow As Long
'integers
'doubles
'booleans
Dim blnClose As Boolean
'variants
Dim vFile As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'prompt user to select a file
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
    blnClose = True
  Else
    'oExcel.WindowState = xlMaximized
    'Application.ActivateMicrosoftApp pjMicrosoftExcel
    blnClose = False
  End If
  If Len(strFile) > 0 Then
    vFile = strFile
    GoTo skip_that
  Else
    vFile = oExcel.GetOpenFilename(FileFilter:="Microsoft Excel *.xls* (*.xls*),", _
                                  Title:="Select a Populated IMS Data Dictionary:", _
                                  ButtonText:="Import", MultiSelect:=False)
    If vFile = False Then GoTo exit_here
  End If
  oExcel.ActivateMicrosoftApp 6 'xlMicrosoftProject
  
skip_that:
  
  'get project uid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'validate the file by its headers
  On Error Resume Next
  oExcel.ScreenUpdating = False
  Set oWorkbook = oExcel.Workbooks(vFile)
  If oWorkbook Is Nothing Then Set oWorkbook = oExcel.Workbooks.Open(vFile)
  Set oWorksheet = oWorkbook.Sheets("Data Dictionary")
  If oWorksheet Is Nothing Then
    MsgBox strFile & " does not appear to be a valid IMS Data Dictionary workbook. The wheet named 'Data Dictionary' not found.", vbExclamation + vbOKOnly, "Invalid Workbook"
    GoTo exit_here
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  lngHeaderRow = oWorksheet.Columns(1).Find("Enterprise", lookat:=1).Row
  lngLastRow = oWorksheet.Cells(lngHeaderRow, 1).End(-4121).Row
  lngNameCol = oWorksheet.Rows(lngHeaderRow).Find("Custom Name", lookat:=1).Column
  lngLastCol = oWorksheet.Rows(lngHeaderRow).End(-4161).Column
  lngDescriptionCol = oWorksheet.Rows(lngHeaderRow).Find("Description", lookat:=1).Column
  
  'get saved dictionary settings
  strSavedSettings = cptDir & "\settings\cpt-data-dictionary.adtg"
  If Dir(strSavedSettings) = vbNullString Then Call cptRefreshDictionary
  
  Set rstSaved = CreateObject("ADODB.Recordset")
  With rstSaved
    .Open strSavedSettings
    .Filter = "PROJECT_ID='" & strGUID & "'"
    .MoveFirst
    For lngRow = lngHeaderRow + 1 To lngLastRow
      'reset validation on description cell
      oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Normal"
      oWorksheet.Cells(lngRow, lngDescriptionCol).HorizontalAlignment = xlCenter
      strCustomName = oWorksheet.Cells(lngRow, lngNameCol).Value
      strNewDescription = oWorksheet.Cells(lngRow, lngDescriptionCol).Value
      If Len(oWorksheet.Cells(lngRow, lngDescriptionCol).Value) > 500 Then
        lngTooLong = lngTooLong + 1
        oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Neutral"
      End If
      If strNewDescription = "<missing>" Then GoTo next_row
      .Find "CUSTOM_NAME='" & strCustomName & "'", , 1 'adSearchForward
      If Not .EOF Then
        If rstSaved("DESCRIPTION").Value <> strNewDescription Then
          oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Good"
          lngNew = lngNew + 1
        End If
      Else
        oWorksheet.Cells(lngRow, lngDescriptionCol).Style = "Bad"
        oWorksheet.Cells(lngRow, lngLastCol + 1).Value = "NOT FOUND"
        oWorksheet.Cells(lngRow, lngLastCol + 1).Style = "Bad"
        lngNotFound = lngNotFound + 1
      End If
next_row:
    Next lngRow
    
    'show the selected file
    If Not oExcel.Visible Then oExcel.Visible = True
    oExcel.ScreenUpdating = True
    oExcel.WindowState = -4140 'xlMinimized
    
    'notify of missing custom fields
    If lngNotFound > 0 Then
      strMsg = Format(lngNotFound, "#,##0") & " fields were found in the selected Excel Workbook but were not found in this Project File." & vbCrLf
      strMsg = strMsg & "These have been marked with cell style 'Bad'." & vbCrLf
      strMsg = strMsg & "Please correct the Excel workbook before importing if necessary. Otherwise, click 'Yes' to proceed." & vbCrLf & vbCrLf
      strMsg = strMsg & "Do you wish to proceed?"
      If MsgBox(strMsg, vbExclamation + vbYesNo, "Missing Custom Fields!") = vbNo Then GoTo exit_here
    End If
    
    'notify if character limit is exceeded
    If lngTooLong > 0 Then
      strMsg = Format(lngTooLong, "#,##0") & " descriptions exceed the 500 character limit; these have been marked with cell style 'Neutral'." & vbCrLf
      strMsg = strMsg & "If you proceed without correcting them, they will be concatenanted (though your spreadsheet will remain unchanged." & vbCrLf & vbCrLf
      strMsg = strMsg & "Do you wish to proceed?"
      If MsgBox(strMsg, vbExclamation + vbYesNo, "Character Limitation Exceeded") = vbNo Then GoTo exit_here
    End If
    
    'confirm updates
    If lngNew > 0 Then
      strMsg = Format(lngNew, "#,##0") & " updated IMS Data Dictionary entries found." & vbCrLf
      strMsg = strMsg & "Please take a moment to indicate unwanted changes in the Excel workbook by marking the cell style as 'Bad.'" & vbCrLf
      strMsg = strMsg & "When new entries have been validated, click 'Yes' below, or hit 'No' to cancel this import." & vbCrLf & vbCrLf
      strMsg = strMsg & "Do you wish to proceed?"
    Else
      MsgBox "No updated entries found.", vbInformation + vbOKOnly, "Import Skipped"
      'Workbook.Close False
      GoTo exit_here
    End If
    
    If MsgBox(strMsg, vbQuestion + vbYesNo, "Confirm Import") = vbNo Then
      'Workbook.Close True
      GoTo exit_here
    Else
      lngNew = 0 'reset the counter
      .MoveFirst
      For lngRow = lngHeaderRow + 1 To lngLastRow
        If oWorksheet.Cells(lngRow, lngDescriptionCol).Style <> "Bad" Then
          strCustomName = oWorksheet.Cells(lngRow, lngNameCol).Value
          strNewDescription = oWorksheet.Cells(lngRow, lngDescriptionCol).Value
          If strNewDescription = "<missing>" Then GoTo next_row2
          .Find "CUSTOM_NAME='" & strCustomName & "'", , 1 'adSearchForward
          If Not .EOF Then
            If rstSaved("DESCRIPTION").Value <> strNewDescription Then
              If rstSaved("DESCRIPTION") <> "<missing>" Then
                oWorksheet.Cells(lngRow, lngLastCol + 1).Value = "DESCRIPTION WAS: " & rstSaved("DESCRIPTION").Value
              End If
              rstSaved("DESCRIPTION").Value = strNewDescription
              .Update
              lngNew = lngNew + 1
            End If
          End If
        Else
          oWorksheet.Cells(lngRow, lngLastCol + 1).Value = oWorksheet.Cells(lngRow, lngLastCol + 1).Value & " - SKIPPED"
        End If
next_row2:
      Next
    End If
    .Filter = 0
    .Save strSavedSettings, adPersistADTG
    .Close
  End With
  
  If lngNew > 0 Then
    Call cptRefreshDictionary
    MsgBox Format(lngNew, "#,##0") & " entries updated.", vbInformation + vbOKOnly, "Import Complete"
  Else
    MsgBox "No updates found.", vbInformation + vbOKOnly, "Import Skipped"
  End If
  
exit_here:
  On Error Resume Next
  If rstSaved.State Then rstSaved.Close
  Set rstSaved = Nothing
  Set oRange = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  If Not oExcel.Visible Then oExcel.Visible = True
  oExcel.ScreenUpdating = True
  Set oExcel = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptImportDataDictionary", Err, Erl)
  Resume exit_here
End Sub
