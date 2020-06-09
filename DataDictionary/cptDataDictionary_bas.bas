Attribute VB_Name = "cptDataDictionary_bas"
'<cpt_version>v1.1.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptExportDataDictionary()
'objects
Dim ListObject As Object 'ListObject
Dim wsLookups As Object 'Worksheet
Dim aTypes As Object
Dim LookupTable As LookupTable
Dim rst As Object 'ADODB.Recordset
Dim xlApp As Object 'Excel.Application
Dim Workbook As Object 'Workbook
Dim Worksheet As Object 'Worksheet
Dim rng As Object 'Excel.Range
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
  
  blnLookups = MsgBox("Replicate Pick Lists in Excel?", vbQuestion + vbYesNo, "Data Dictionary") = vbYes
  
  'set up a workbook/worksheet
  cptDataDictionary_frm.lblStatus.Caption = "Creating Excel Workbook..."
  Set xlApp = CreateObject("Excel.Application")
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Worksheets(1)
  Worksheet.Name = "Data Dictionary"
  
  'create lookups worksheet (optional)
  If blnLookups Then
    Set wsLookups = Workbook.Sheets.Add(After:=Workbook.Sheets(1))
    wsLookups.Name = "LOOKUPS"
    wsLookups.Activate
    xlApp.ActiveWindow.Zoom = 83
    Worksheet.Activate
  End If
  
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
  
  blnExists = Dir(cptDir & "\settings\cpt-data-dictionary.adtg") <> vbNullString

  If blnExists Then
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open cptDir & "\settings\cpt-data-dictionary.adtg"
  End If
  
  'count of custom fields = local + enterprise
  lngItems = 260 + (188778000 - 188776000)
  
  Set aTypes = CreateObject("System.Collections.SortedList")
  aTypes.Add "Cost", 10
  aTypes.Add "Date", 10
  aTypes.Add "Duration", 10
  aTypes.Add "Flag", 20
  aTypes.Add "Finish", 10
  aTypes.Add "Outline Code", 10
  aTypes.Add "Number", 20
  aTypes.Add "Start", 10
  aTypes.Add "Text", 30
  
  'prep for data dump
  lngRow = lngHeaderRow
  'export local custom fields
  For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
    For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Outline Code", "Number", "Start", "Text")
      For intField = 1 To aTypes.Item(vFieldType)
        lngField = FieldNameToFieldConstant(vFieldType & intField, vFieldScope)
        strFieldName = CustomFieldGetName(lngField)
        If Len(strFieldName) > 0 Then
          lngRow = lngRow + 1
          Worksheet.Cells(lngRow, 1).Value = False
          Worksheet.Cells(lngRow, 2).Value = Choose(CInt(vFieldScope) + 1, "Task", "Resource", "Project")
          Worksheet.Cells(lngRow, 3).Value = CStr(vFieldType)
          Worksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
          Worksheet.Cells(lngRow, 5).Value = strFieldName
          'get attributes
          If Len(CustomFieldGetFormula(lngField)) > 0 Then
            Worksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
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
                Set LookupTable = ActiveProject.OutlineCodes(FieldConstantToFieldName(lngField))
                If Len(LookupTable(intListItem).Description) > 0 Then
                  If Left(LookupTable(intListItem).Description, Len(LookupTable(intListItem).FullName)) = LookupTable(intListItem).FullName Then
                    strAttributes = strAttributes & vbCrLf & LookupTable(intListItem).Description
                    If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = LookupTable(intListItem).Description
                  Else
                    strAttributes = strAttributes & vbCrLf & LookupTable(intListItem).FullName & " - " & LookupTable(intListItem).Description
                    If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = LookupTable(intListItem).FullName & " - " & LookupTable(intListItem).Description
                  End If
                Else
                  strAttributes = strAttributes & vbCrLf & LookupTable(intListItem).FullName
                  If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = LookupTable(intListItem).FullName
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
                  If blnLookups Then wsLookups.Cells(2, intListItem, lngLookupCol) = strValue
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
              With Worksheet.Cells(lngRow, 6).Validation
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
               Worksheet.Cells(lngRow, 6).Value = UCase(strFieldName) & " LOOKUP:"
            Else 'don't
              If Len(strAttributes) > 0 Then Worksheet.Cells(lngRow, 6).Value = "Lookup Values:" & strAttributes
            End If 'blnLookups
            
          End If 'Not LookupTable Is Nothing Then
          
        End If 'Len(strFieldName) > 0

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
      Worksheet.Cells(lngRow, 1).Value = True
      Worksheet.Cells(lngRow, 2).Value = "Enterprise"
      Worksheet.Cells(lngRow, 3).Value = "Enterprise"
      Worksheet.Cells(lngRow, 4).Value = FieldConstantToFieldName(lngField)
      Worksheet.Cells(lngRow, 5).Value = FieldConstantToFieldName(lngField)
      If Len(CustomFieldGetFormula(lngField)) > 0 Then
        Worksheet.Cells(lngRow, 6).Value = CustomFieldGetFormula(lngField)
      End If
      strAttributes = ""
      Set LookupTable = Nothing
      On Error Resume Next
      Set LookupTable = GlobalOutlineCodes(FieldConstantToFieldName(lngField)).LookupTable
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If Not LookupTable Is Nothing Then
        If blnLookups Then
          lngLookupCol = wsLookups.[XFD2].End(-4159).Column
          If wsLookups.Cells(1, lngLookupCol) <> "" Then lngLookupCol = lngLookupCol + 2
          wsLookups.Cells(1, lngLookupCol) = UCase(FieldConstantToFieldName(lngField))
          wsLookups.Cells(2, lngLookupCol) = UCase(FieldConstantToFieldName(lngField)) & " LOOKUP:"
        End If
        For intListItem = 1 To LookupTable.Count
          If Len(LookupTable(intListItem).Description) > 0 Then
            If Left(LookupTable(intListItem).Description, Len(LookupTable(intListItem).FullName)) = LookupTable(intListItem).FullName Then
              strAttributes = strAttributes & vbCrLf & LookupTable(intListItem).Description
              If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = LookupTable(intListItem).Description
            Else
              strAttributes = strAttributes & vbCrLf & LookupTable(intListItem).FullName & " - " & LookupTable(intListItem).Description
              If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = LookupTable(intListItem).FullName & " - " & LookupTable(intListItem).Description
            End If
          Else
            strAttributes = strAttributes & vbCrLf & LookupTable(intListItem).FullName
            If blnLookups Then wsLookups.Cells(2 + intListItem, lngLookupCol) = LookupTable(intListItem).FullName
          End If
        Next intListItem
        
        If blnLookups Then 'use validation
          'name the range
          wsLookups.ListObjects.Add(SourceType:=1, Source:=wsLookups.Range(wsLookups.Cells(1, lngLookupCol), wsLookups.Cells(2 + LookupTable.Count, lngLookupCol)).Address(True, True), xllistobjecthasheaders:=1).Name = UCase(Replace(FieldConstantToFieldName(lngField), " ", "_"))
          wsLookups.Columns(lngLookupCol).AutoFit
          wsLookups.Columns(lngLookupCol + 1).ColumnWidth = 2
          With Worksheet.Cells(lngRow, 6).Validation
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
          Worksheet.Cells(lngRow, 6).Value = UCase(FieldConstantToFieldName(lngField)) & " LOOKUP:"
        Else
          If Len(strAttributes) > 0 Then Worksheet.Cells(lngRow, 6).Value = "Lookup Values:" & strAttributes
        End If 'blnLookups
        
      End If 'Not LookupTable Is Nothing Then
      
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
  If blnLookups Then
    wsLookups.Activate
    xlApp.ActiveWindow.Zoom = 85
    wsLookups.[A2].Select
    xlApp.ActiveWindow.FreezePanes = True
    wsLookups.[A3].Select
    wsLookups.Rows(2).Hidden = True
  End If
  
  'convert to table / format it
  Worksheet.Activate
  xlApp.ActiveWindow.ScrollRow = lngHeaderRow
  Set rng = Worksheet.Range(Worksheet.Cells(lngHeaderRow, 1).End(-4161), Worksheet.Cells(lngHeaderRow, 1).End(-4121))
  Worksheet.ListObjects.Add(1, rng, , 1).Name = "DATA_DICTIONARY"
  'autofit
  Worksheet.Range("DATA_DICTIONARY[#All]").Select
  rng.Columns.AutoFit
  rng.Rows.AutoFit
  rng.VerticalAlignment = xlCenter
  lngCol = Worksheet.Rows(lngHeaderRow).Find("Attributes", lookat:=1).Column
  Worksheet.Columns(lngCol).ColumnWidth = 100
  Worksheet.Columns(lngCol).WrapText = True
  lngCol = Worksheet.Rows(lngHeaderRow).Find("Description", lookat:=1).Column
  Worksheet.Columns(lngCol).ColumnWidth = 100
  Worksheet.Cells(lngHeaderRow + 1, 1).Select
  
  cptDataDictionary_frm.lblStatus.Caption = "Opening..."
  
exit_here:
  On Error Resume Next
  Set ListObject = Nothing
  Set wsLookups = Nothing
  Set aTypes = Nothing
  Set LookupTable = Nothing
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
  If Err.Number = 1101 Or Err.Number = 1004 Then
    Err.Clear
    Resume next_field
  Else
    Call cptHandleErr("cptExportCustomFields_bas", "cptExportDataDictionary", Err, Erl)
  End If
  
End Sub

Sub ShowFrmCptDataDictionary()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  cptDataDictionary_frm.lboCustomFields.Clear
  Call cptRefreshDictionary
  cptDataDictionary_frm.txtFilter.SetFocus
  cptDataDictionary_frm.Show
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "ShowFrmCptDataDictionary()", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshDictionary()
'objects
Dim aTypes As Object
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
  Set rst = CreateObject("ADODB.Recordset")
  With rst
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
    
    Set aTypes = CreateObject("System.Collections.SortedList")
    aTypes.Add "Cost", 10
    aTypes.Add "Date", 10
    aTypes.Add "Duration", 10
    aTypes.Add "Flag", 20
    aTypes.Add "Finish", 10
    aTypes.Add "Number", 20
    aTypes.Add "Start", 10
    aTypes.Add "Text", 30
    aTypes.Add "Outline Code", 10
        
    For Each vFieldScope In Array(0, 1) '0 = pjTask; 1 = pjResource; 2 = pjProject
      For Each vFieldType In Array("Cost", "Date", "Duration", "Flag", "Finish", "Number", "Start", "Text", "Outline Code")
        For intField = 1 To aTypes.Item(vFieldType) 'lngMax
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
  Set aTypes = Nothing
  If rst.State = 1 Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptRefreshDictionary", Err, Erl)
  Resume exit_here
End Sub

Sub cptImportDataDictionary(Optional strFile As String)
'objects
Dim rst As Object 'ADODB.Recordset
Dim xlApp As Object
Dim Workbook As Object
Dim Worksheet As Object
Dim rng As Object
Dim ListObject As Object
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
  Set xlApp = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If xlApp Is Nothing Then
    Set xlApp = CreateObject("Excel.Application")
    blnClose = True
  Else
    'xlApp.WindowState = xlMaximized
    'Application.ActivateMicrosoftApp pjMicrosoftExcel
    blnClose = False
  End If
  If Len(strFile) > 0 Then
    vFile = strFile
    GoTo skip_that
  Else
    vFile = xlApp.GetOpenFilename(FileFilter:="Microsoft Excel *.xls* (*.xls*),", _
                                  Title:="Select a Populated IMS Data Dictionary:", _
                                  ButtonText:="Import", MultiSelect:=False)
    If vFile = False Then GoTo exit_here
  End If
  xlApp.ActivateMicrosoftApp 6 'xlMicrosoftProject
  
skip_that:
  
  'get project uid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'validate the file by its headers
  On Error Resume Next
  xlApp.ScreenUpdating = False
  Set Workbook = xlApp.Workbooks(vFile)
  If Workbook Is Nothing Then Set Workbook = xlApp.Workbooks.Open(vFile)
  Set Worksheet = Workbook.Sheets("Data Dictionary")
  If Worksheet Is Nothing Then
    MsgBox strFile & " does not appear to be a valid IMS Data Dictionary workbook. The wheet named 'Data Dictionary' not found.", vbExclamation + vbOKOnly, "Invalid Workbook"
    GoTo exit_here
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  lngHeaderRow = Worksheet.Columns(1).Find("Enterprise", lookat:=1).Row
  lngLastRow = Worksheet.Cells(lngHeaderRow, 1).End(-4121).Row
  lngNameCol = Worksheet.Rows(lngHeaderRow).Find("Custom Name", lookat:=1).Column
  lngLastCol = Worksheet.Rows(lngHeaderRow).End(-4161).Column
  lngDescriptionCol = Worksheet.Rows(lngHeaderRow).Find("Description", lookat:=1).Column
  
  'get saved dictionary settings
  strSavedSettings = cptDir & "\settings\cpt-data-dictionary.adtg"
  If Dir(strSavedSettings) = vbNullString Then Call cptRefreshDictionary
  
  Set rst = CreateObject("ADODB.Recordset")
  With rst
    .Open strSavedSettings
    .Filter = "PROJECT_ID='" & strGUID & "'"
    .MoveFirst
    For lngRow = lngHeaderRow + 1 To lngLastRow
      'reset validation on description cell
      Worksheet.Cells(lngRow, lngDescriptionCol).Style = "Normal"
      Worksheet.Cells(lngRow, lngDescriptionCol).HorizontalAlignment = xlCenter
      strCustomName = Worksheet.Cells(lngRow, lngNameCol).Value
      strNewDescription = Worksheet.Cells(lngRow, lngDescriptionCol).Value
      If Len(Worksheet.Cells(lngRow, lngDescriptionCol).Value) > 500 Then
        lngTooLong = lngTooLong + 1
        Worksheet.Cells(lngRow, lngDescriptionCol).Style = "Neutral"
      End If
      If strNewDescription = "<missing>" Then GoTo next_row
      .Find "CUSTOM_NAME='" & strCustomName & "'", , 1 'adSearchForward
      If Not .EOF Then
        If rst("DESCRIPTION").Value <> strNewDescription Then
          Worksheet.Cells(lngRow, lngDescriptionCol).Style = "Good"
          lngNew = lngNew + 1
        End If
      Else
        Worksheet.Cells(lngRow, lngDescriptionCol).Style = "Bad"
        Worksheet.Cells(lngRow, lngLastCol + 1).Value = "NOT FOUND"
        Worksheet.Cells(lngRow, lngLastCol + 1).Style = "Bad"
        lngNotFound = lngNotFound + 1
      End If
next_row:
    Next lngRow
    
    'show the selected file
    If Not xlApp.Visible Then xlApp.Visible = True
    xlApp.ScreenUpdating = True
    xlApp.WindowState = -4140 'xlMinimized
    
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
        If Worksheet.Cells(lngRow, lngDescriptionCol).Style <> "Bad" Then
          strCustomName = Worksheet.Cells(lngRow, lngNameCol).Value
          strNewDescription = Worksheet.Cells(lngRow, lngDescriptionCol).Value
          If strNewDescription = "<missing>" Then GoTo next_row2
          .Find "CUSTOM_NAME='" & strCustomName & "'", , 1 'adSearchForward
          If Not .EOF Then
            If rst("DESCRIPTION").Value <> strNewDescription Then
              If rst("DESCRIPTION") <> "<missing>" Then
                Worksheet.Cells(lngRow, lngLastCol + 1).Value = "DESCRIPTION WAS: " & rst("DESCRIPTION").Value
              End If
              rst("DESCRIPTION").Value = strNewDescription
              .Update
              lngNew = lngNew + 1
            End If
          End If
        Else
          Worksheet.Cells(lngRow, lngLastCol + 1).Value = Worksheet.Cells(lngRow, lngLastCol + 1).Value & " - SKIPPED"
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
  Set rst = Nothing
  If rst.State Then rst.Close
  Set rng = Nothing
  Set ListObject = Nothing
  Set Worksheet = Nothing
  Set Workbook = Nothing
  If Not xlApp.Visible Then xlApp.Visible = True
  xlApp.ScreenUpdating = True
  Set xlApp = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_bas", "cptImportDataDictionary", Err, Erl)
  Resume exit_here
End Sub
