Attribute VB_Name = "cptCostRateTables_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit

Sub cptShowCostRateTablesForm()
  'objects
  'strings
  Dim strOverwrite As String
  Dim strAddNew As String
  Dim strCustomFieldName As String
  'longs
  Dim lngCustomField As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  With cptCostRateTables_frm
    .Caption = "Cost Rate Tables (" & cptGetVersion("cptCostRateTables_frm") & ")"
    .lblProgress.Width = .lblStatus.Width
    .lblStatus.Caption = "Ready..."
    With .cboStatusField
      .Clear
      For lngItem = 1 To 30
        lngCustomField = FieldNameToFieldConstant("Text" & lngItem, pjResource)
        strCustomFieldName = CustomFieldGetName(lngCustomField)
        If Len(strCustomFieldName) > 0 Then
          .AddItem
          .List(lngItem - 1, 0) = lngCustomField
          .List(lngItem - 1, 1) = "Text" & lngItem & " (" & strCustomFieldName & ")"
        Else
          .AddItem
          .List(lngItem - 1, 0) = lngCustomField
          .List(lngItem - 1, 1) = "Text" & lngItem
        End If
      Next lngItem
    End With
    If ActiveProject.ResourceCount > 0 Then
      .tglExport = True
    Else
      .tglImport = True
    End If
    strOverwrite = cptGetSetting("CostRateTables", "chkOverwrite")
    If Len(strOverwrite) > 0 Then
      .chkOverwrite = CBool(strOverwrite)
    Else
      .chkOverwrite = True 'default
    End If
    strAddNew = cptGetSetting("CostRateTables", "chkAddNew")
    If Len(strAddNew) > 0 Then
      .chkAddNew = CBool(strAddNew)
    Else
      .chkAddNew = True 'default
    End If
    .Show
  End With

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptCostRateTables_bas", "cptShowCostRateTablesForm", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportCostRateTables(strCostRateTables As String)
  'objects
  Dim oPayRate As PayRate
  Dim oCostRateTable As CostRateTable
  Dim oResource As Resource
  Dim oExcel As Object 'Excel.Application
  Dim oWorkbook As Object 'Excel.Workbook
  Dim oWorksheet As Object 'Excel.Worksheet
  'strings
  Dim strType As String
  Dim strRateTable As String
  Dim strResource As String
  'longs
  Dim lngCostRateTable As Long
  Dim lngLastRow As Long
  Dim lngResource As Long
  Dim lngResourceCount As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vCostRateTable As Variant
  'dates
  
  cptCostRateTables_frm.lblStatus.Caption = "Getting Excel..."
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = xlCalculationManual
  oExcel.ScreenUpdating = False
  Set oWorksheet = oWorkbook.Sheets(1)
  cptCostRateTables_frm.lblStatus.Caption = "Creating Header..."
  oWorksheet.[A1:G1] = Split(("RESOURCE,TYPE,RATE TABLE,EFFECTIVE DATE,STANDARD RATE,OVERTIME RATE,COST PER USE"), ",")
  
  lngResourceCount = ActiveProject.ResourceCount
  lngResource = 0
  For Each oResource In ActiveProject.Resources
    lngResource = lngResource + 1
    strResource = oResource.Name
    For Each vCostRateTable In Split(strCostRateTables, ",")
      If vCostRateTable = "" Then GoTo next_cost_rate_table
      lngCostRateTable = Switch(vCostRateTable = "A", 1, vCostRateTable = "B", 2, vCostRateTable = "C", 3, vCostRateTable = "D", 4, vCostRateTable = "E", 5)
      Set oCostRateTable = oResource.CostRateTables(lngCostRateTable)
      strType = Choose(oResource.Type + 1, "WORK", "MATERIAL", "COST")
      For Each oPayRate In oCostRateTable.PayRates
        lngLastRow = oWorksheet.[A1048576].End(-4162).Row + 1 '-4162 = xlUp
        oWorksheet.Cells(lngLastRow, 1) = strResource
        oWorksheet.Cells(lngLastRow, 2) = strType
        oWorksheet.Cells(lngLastRow, 3) = CStr(vCostRateTable)
        oWorksheet.Cells(lngLastRow, 4) = FormatDateTime(oPayRate.EffectiveDate, vbShortDate)
        oWorksheet.Cells(lngLastRow, 5) = oPayRate.StandardRate
        oWorksheet.Cells(lngLastRow, 6) = oPayRate.OvertimeRate
        oWorksheet.Cells(lngLastRow, 7) = oPayRate.CostPerUse
      Next oPayRate
next_cost_rate_table:
    Next vCostRateTable
    Application.StatusBar = Format(lngResource, "#,##0") & "/" & Format(lngResourceCount, "#,##0") & "...(" & Format(lngResource / lngResourceCount, "0%") & ")"
    cptCostRateTables_frm.lblStatus.Caption = Format(lngResource, "#,##0") & "/" & Format(lngResourceCount, "#,##0") & "...(" & Format(lngResource / lngResourceCount, "0%") & ")"
    cptCostRateTables_frm.lblProgress.Width = (lngResource / lngResourceCount) * cptCostRateTables_frm.lblStatus.Width
    DoEvents
  Next oResource

  With cptCostRateTables_frm
    .lblProgress.Width = .lblStatus.Width
    .lblStatus = "Complete."
  End With
  Application.StatusBar = "Complete."
  
  oExcel.Visible = True
  With oExcel.ActiveWindow
    .Zoom = 85
    .SplitRow = 1
    .SplitColumn = 0
    .FreezePanes = True
  End With
  oWorksheet.Columns.AutoFit

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oPayRate = Nothing
  Set oCostRateTable = Nothing
  Set oResource = Nothing
  Set oWorksheet = Nothing
  oExcel.Visible = True
  oExcel.ScreenUpdating = True
  oExcel.Calculation = xlCalculationAutomatic
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  MsgBox Err.Number & ":" & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptImportCostRateTables(lngField As Long)
  'objects
  Dim oPayRate As MSProject.PayRate
  Dim oCostRateTable As MSProject.CostRateTable
  Dim oResource As MSProject.Resource
  Dim oExcel As Object 'Excel.Application
  Dim oWorkbook As Object 'Excel.Workbook
  Dim oWorksheet As Object 'Excel.Worksheet
  'strings
  Dim strCostRateTable As String
  Dim strType As String
  Dim strWorkbook As String
  'longs
  Dim lngCostRateTable As Long
  Dim lngRow As Long
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  Dim blnOverwrite As Boolean
  Dim blnAddResources As Boolean
  'variants
  Dim vCostPerUse As Variant
  Dim vOvtRate As Variant
  Dim vStdRate As Variant
  Dim vEffectiveDate As Variant
  'dates
  
  Application.ActiveWindow.TopPane.Activate
  ViewApply "Resource Sheet"
  FilterClear
  
  'clear out the field
  Application.ScreenUpdating = True
  Application.Calculation = pjAutomatic
  cptCostRateTables_frm.lblStatus.Caption = "Clearing Field..."
  For Each oResource In ActiveProject.Resources
    EditGoTo oResource.ID
    oResource.SetField lngField, ""
  Next oResource
  DoEvents
  
  cptCostRateTables_frm.lblStatus.Caption = "Getting Excel..."
  On Error Resume Next
  Set oExcel = GetObject(, "Excel.Application")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then
    Set oExcel = CreateObject("Excel.Application")
  End If
  With oExcel.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .ButtonName = "Import"
    .Title = "Import Cost Rate Tables:"
    .Filters.Clear
    .Filters.Add "Microsoft Excel", "*.xls*"
    .Show
    If .SelectedItems.Count > 0 Then
      strWorkbook = .SelectedItems(1)
      blnAddResources = MsgBox("Add Resources in Workbook but not in this project?", vbQuestion + vbYesNo, "Confirm Add New Resources") = vbYes
      blnOverwrite = MsgBox("Overwrite existing Cost Rate Tables?", vbQuestion + vbYesNo, "Confirm Overwrite Cost Rate Tables") = vbYes
    Else
      cptCostRateTables_frm.lblStatus.Caption = "Ready..."
      GoTo exit_here
    End If
  End With

  Application.Calculation = pjManual
  Application.ScreenUpdating = False
  cptCostRateTables_frm.lblStatus.Caption = "Opening Workbook..."
  Set oWorkbook = oExcel.Workbooks.Open(strWorkbook)
  Set oWorksheet = oWorkbook.Sheets(1)
      
  lngLastRow = oWorksheet.[A1048576].End(-4162).Row '-4162 = xlUp
  For lngRow = 2 To lngLastRow
    'get/add resource
    If lngRow = 2 Then
      On Error Resume Next
      Set oResource = ActiveProject.Resources(oWorksheet.Cells(lngRow, 1).Value)
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    Else
      If oResource.Name <> oWorksheet.Cells(lngRow, 1).Value Then
        Set oResource = Nothing
        On Error Resume Next
        Set oResource = ActiveProject.Resources(oWorksheet.Cells(lngRow, 1).Value)
        If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      Else
        GoTo cost_rate_tables
      End If
    End If
    If oResource Is Nothing Then
      If blnAddResources Then
        Set oResource = ActiveProject.Resources.Add(oWorksheet.Cells(lngRow, 1).Value)
        strType = oWorksheet.Cells(lngRow, 2).Value
        oResource.Type = Switch(strType = "WORK", pjResourceTypeWork, strType = "COST", pjResourceTypeCost, strType = "MATERIAL", pjResourceTypeMaterial)
        oResource.SetField lngField, "ADDED"
        GoTo cost_rate_tables
      End If
    Else
      oResource.SetField lngField, "UPDATED: "
    End If
        
    'get cost rate table
cost_rate_tables:
    strCostRateTable = oWorksheet.Cells(lngRow, 3).Value
    lngCostRateTable = Switch(strCostRateTable = "A", 1, strCostRateTable = "B", 2, strCostRateTable = "C", 3, strCostRateTable = "D", 4, strCostRateTable = "E", 5)
    Set oCostRateTable = oResource.CostRateTables(lngCostRateTable)
    If oResource.GetField(lngField) <> "ADDED" Then
      If InStr(oResource.GetField(lngField), strCostRateTable & ",") = 0 Then 'not wiped yet
        If blnOverwrite Then
          For Each oPayRate In oCostRateTable.PayRates
            If oPayRate.Index = 1 Then
              oPayRate.StandardRate = 0
              oPayRate.OvertimeRate = 0
              oPayRate.CostPerUse = 0
            Else
              oPayRate.Delete
            End If
          Next oPayRate
          oResource.SetField lngField, oResource.GetField(lngField) & strCostRateTable & IIf(strCostRateTable <> "E", ",", "")
        Else
          'todo: allow append vs overwrite?
        End If
      End If
    End If
    vEffectiveDate = oWorksheet.Cells(lngRow, 4).Value
    vStdRate = oWorksheet.Cells(lngRow, 5).Value
    vOvtRate = oWorksheet.Cells(lngRow, 6).Value
    vCostPerUse = oWorksheet.Cells(lngRow, 7).Value
    If CDate(vEffectiveDate) > #1/1/1984# Then
      oCostRateTable.PayRates.Add vEffectiveDate, vStdRate, vOvtRate, vCostPerUse
    Else
      oCostRateTable.PayRates(1).StandardRate = vStdRate
      If Not IsEmpty(vOvtRate) Then oCostRateTable.PayRates(1).OvertimeRate = vOvtRate
      oCostRateTable.PayRates(1).CostPerUse = vCostPerUse
    End If
    Application.StatusBar = Format(lngRow, "#,##0") & "/" & Format(lngLastRow, "#,##0") & "...(" & Format(lngRow / lngLastRow, "0%") & ")"
    cptCostRateTables_frm.lblStatus.Caption = Format(lngRow, "#,##0") & "/" & Format(lngLastRow, "#,##0") & "...(" & Format(lngRow / lngLastRow, "0%") & ")"
    cptCostRateTables_frm.lblProgress.Width = (lngRow / lngLastRow) * cptCostRateTables_frm.lblStatus.Width
    DoEvents
  Next lngRow
    
  With cptCostRateTables_frm
    .lblProgress.Width = .lblStatus.Width
    .lblStatus.Caption = "Complete."
  End With
  Application.StatusBar = "Complete."
  
  oWorkbook.Close False
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Application.ScreenUpdating = True
  Application.Calculation = pjAutomatic
  Set oPayRate = Nothing
  Set oCostRateTable = Nothing
  Set oResource = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  MsgBox Err.Number & ":" & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub
