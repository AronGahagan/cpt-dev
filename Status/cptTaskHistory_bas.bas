Attribute VB_Name = "cptTaskHistory_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit

Sub cptShowTaskHistoryFrm()
  Call cptUpdateTaskHistory
  cptTaskHistory_frm.Caption = "Task History (" & cptGetVersion("cptTaskHistory_bas") & ")"
  cptTaskHistory_frm.Show False
End Sub

Sub cptUpdateTaskHistory()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strNote As String
  Dim strFile As String
  Dim strProgramAcronym As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtFinish As Date
  Dim dtStart As Date
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  'ensure program acroynm
  strProgramAcronym = cptGetProgramAcronym
  'ensure file exists
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox "Please run 'capture week' for each week's file that you want recorded.", vbCritical + vbOKOnly, "File Not Found"
    GoTo exit_here
  End If
  
  With cptTaskHistory_frm
    .lboHeader.Clear
    .lboHeader.ColumnCount = 8
    .lboHeader.AddItem
    .lboHeader.List(0, 0) = "STATUS_DATE"
    .lboHeader.List(0, 1) = "STATUS DATE"
    .lboHeader.List(0, 2) = "START"
    .lboHeader.List(0, 3) = "DUR"
    .lboHeader.List(0, 4) = "FINISH"
    .lboHeader.List(0, 5) = "RDur"
    .lboHeader.List(0, 6) = "RWork"
    .lboHeader.List(0, 7) = "STATUS NOTE"
    .lboTaskHistory.Clear
    .lboTaskHistory.ColumnCount = 8
    .lblUID.Caption = "-"
    .txtVariance = ""
    .lblWarning.Visible = False
    If ActiveSelection.Tasks.Count <> 1 Then
      If ActiveSelection.Tasks.Count > 1 Then
        .lboTaskHistory.ColumnCount = 2
        .lboTaskHistory.AddItem
        .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = "One at a time, please..."
      Else
        .lboTaskHistory.Clear
      End If
    Else
      If ActiveSelection.Tasks(1).Summary Then
        .lblWarning.Caption = "History not captured on summary tasks."
        .lblWarning.Visible = True
        GoTo exit_here
      ElseIf Not ActiveSelection.Tasks(1).Active Then
        .lblWarning.Caption = "History not captured on inactive tasks."
        .lblWarning.Visible = True
        GoTo exit_here
      ElseIf ActiveSelection.Tasks(1).ExternalTask Then
        .lblWarning.Caption = "History not captured on external tasks."
        .lblWarning.Visible = True
        GoTo exit_here
      End If
      lngUID = ActiveSelection.Tasks(1).UniqueID
      .lblUID.Caption = lngUID
      Set oRecordset = CreateObject("ADODB.Recordset")
      oRecordset.Open strFile
      If oRecordset.RecordCount > 0 Then
        oRecordset.Filter = "PROJECT='" & strProgramAcronym & "' AND TASK_UID=" & CInt(lngUID)
        oRecordset.Sort = "STATUS_DATE desc"
        If oRecordset.EOF Then
          .lblWarning.Caption = "No history for UID " & lngUID & "."
          .lblWarning.Visible = True
          GoTo exit_here
        Else
          oRecordset.MoveFirst
          Do While Not oRecordset.EOF
            .lboTaskHistory.AddItem
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 0) = oRecordset("STATUS_DATE")
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = FormatDateTime(oRecordset("STATUS_DATE"), vbShortDate) & IIf(Len(oRecordset("NOTE")) > 0, "*", "")
            If oRecordset("TASK_AS") > 0 Then
              dtStart = oRecordset("TASK_AS")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 2) = "[" & FormatDateTime(dtStart, vbShortDate) & "]"
            Else
              dtStart = oRecordset("TASK_START")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 2) = FormatDateTime(dtStart, vbShortDate)
            End If
          
            If oRecordset("TASK_AF") > 0 Then
              dtFinish = oRecordset("TASK_AF")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 4) = "[" & FormatDateTime(dtFinish, vbShortDate) & "]"
            Else
              dtFinish = oRecordset("TASK_FINISH")
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 4) = FormatDateTime(dtFinish, vbShortDate)
            End If
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 3) = Application.DateDifference(dtStart, dtFinish) / 480 & "d"
            .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 5) = oRecordset("TASK_RD") & "d"
            'todo: Remaining Work
            'todo: EV%
            'todo: remove NOTE?
            strNote = oRecordset("NOTE")
            If Len(strNote) = 0 Then
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 7) = ""
            ElseIf Len(strNote) > 10 Then
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 7) = Left(strNote, 7) & "..."
            Else
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 7) = strNote
            End If
            If Len(strNote) > 0 Then
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = FormatDateTime(oRecordset("STATUS_DATE"), vbShortDate) & "*"
            Else
              .lboTaskHistory.List(.lboTaskHistory.ListCount - 1, 1) = FormatDateTime(oRecordset("STATUS_DATE"), vbShortDate)
            End If
            oRecordset.MoveNext
          Loop
          oRecordset.Filter = 0
          oRecordset.Close
        End If
      Else
        .lblWarning.Caption = "No records found."
        .lblWarning.Visible = True
      End If
    End If
  End With

exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptUpdateTaskHistory", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateTaskHistoryNote(lngUID As Long, dtStatus As Date, strVariance As String)
  'objects
  Dim oRecordsetNew As ADODB.Recordset
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  Dim blnFieldExists As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'ensure NOTE field exists
  strFile = cptDir & "\settings\cpt-cei.adtg"
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strFile
  blnFieldExists = False
  For lngItem = 0 To oRecordset.Fields.Count - 1
    If oRecordset.Fields(lngItem).Name = "NOTE" Then
      blnFieldExists = True
      Exit For
    End If
  Next lngItem
  
  'add new field if necessary
  If Not blnFieldExists Then
    Set oRecordsetNew = CreateObject("ADODB.Recordset")
    'recreate the field structure
    For lngItem = 0 To oRecordset.Fields.Count - 1
      If oRecordset.Fields(lngItem).Type = adVarChar Then
        oRecordsetNew.Fields.Append CStr(oRecordset.Fields(lngItem).Name), oRecordset.Fields(lngItem).Type, oRecordset.Fields(lngItem).DefinedSize
      Else
        oRecordsetNew.Fields.Append CStr(oRecordset.Fields(lngItem).Name), oRecordset.Fields(lngItem).Type
      End If
    Next lngItem
    oRecordsetNew.Fields.Append "NOTE", adVarChar, 255
    oRecordsetNew.Open
    'copy recordset in todo: there must be a faster way
    oRecordset.MoveFirst
    Do While Not oRecordset.EOF
      oRecordsetNew.AddNew
      For lngItem = 0 To oRecordset.Fields.Count - 1
        oRecordsetNew(lngItem) = oRecordset(lngItem)
      Next
      oRecordset.MoveNext
    Loop
    oRecordset.Save Replace(strFile, ".adtg", "-backup.adtg"), adPersistADTG
    oRecordset.Close
    Kill strFile
    oRecordsetNew.Save strFile, adPersistADTG
    Set oRecordset = New ADODB.Recordset
    oRecordset.Open strFile
  End If
  
  oRecordset.MoveFirst
  oRecordset.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND TASK_UID=" & CInt(lngUID) & " AND STATUS_DATE=#" & FormatDateTime(dtStatus, vbGeneralDate) & "#"
  If Not oRecordset.EOF Then
    oRecordset.Update Array("NOTE"), Array(strVariance)
  End If
  oRecordset.Filter = 0
  oRecordset.Save strFile, adPersistADTG
  oRecordset.Close
  
  With cptTaskHistory_frm
    If Len(strVariance) = 0 Then
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 7) = ""
    ElseIf Len(strVariance) > 10 Then
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 7) = Left(strVariance, 7) & "..."
    Else
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 7) = strVariance
    End If
    If Len(strVariance) > 0 Then
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 1) = FormatDateTime(dtStatus, vbShortDate) & "*"
    Else
      .lboTaskHistory.List(.lboTaskHistory.ListIndex, 1) = FormatDateTime(dtStatus, vbShortDate)
    End If
  End With
  
exit_here:
  On Error Resume Next
  If oRecordsetNew.State Then
    oRecordsetNew.Filter = 0
    oRecordsetNew.Save strFile, adPersistADTG
    oRecordsetNew.Close
  End If
  Set oRecordsetNew = Nothing
  If oRecordset.State Then
    oRecordset.Filter = 0
    oRecordset.Save strFile, adPersistADTG
    oRecordset.Close
  End If
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptUpdateTaskHistoryNote", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetTaskHistoryNote(dtStatus As Date, lngUID As Long)
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) = vbNullString Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  With cptTaskHistory_frm
    .txtVariance = ""
    Set oRecordset = CreateObject("ADODB.Recordset")
    oRecordset.Open strFile
    oRecordset.MoveFirst
    oRecordset.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND TASK_UID=" & lngUID & " AND STATUS_DATE=#" & FormatDateTime(dtStatus, vbGeneralDate) & "#"
    If Not oRecordset.EOF Then
      .txtVariance = oRecordset("NOTE")
    End If
  End With
    
exit_here:
  On Error Resume Next
  oRecordset.Filter = 0
  oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHIstory_bas", "cptGetTaskHistoryNote", Err, Erl)
  Resume exit_here
End Sub

Sub cptExportTaskHistory(lngUID As Long)
  'objects
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  Dim oRecordset As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vRow As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) = vbNullString Then Exit Sub
  Application.StatusBar = "Getting data..."
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strFile, , adOpenKeyset
  oRecordset.MoveFirst
  Application.StatusBar = "Filtering for UID " & lngUID & "..."
  oRecordset.Filter = "PROJECT='" & cptGetProgramAcronym & "' AND TASK_UID=" & lngUID
  Application.StatusBar = "Sorting by STATUS_DATE descending..."
  oRecordset.Sort = "STATUS_DATE desc"
  If Not oRecordset.EOF Then
    Application.StatusBar = "Getting Excel..."
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    oExcel.ScreenUpdating = False
    Application.StatusBar = "Setting up workbook..."
    Set oWorkbook = oExcel.Workbooks.Add
    oExcel.Calculation = xlCalculationManual
    Set oWorksheet = oWorkbook.Sheets(1)
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Name = "Task History"
    oWorksheet.[A1].Value = ActiveProject.Name
    oWorksheet.[A1:E1].Merge
    Application.StatusBar = "Building header..."
    vRow = oWorksheet.Range(oWorksheet.[A3], oWorksheet.Cells(3, oRecordset.Fields.Count)).Value
    For lngItem = 0 To oRecordset.Fields.Count - 1
      vRow(1, lngItem + 1) = oRecordset.Fields(lngItem).Name
    Next lngItem
    oWorksheet.Range(oWorksheet.[A3], oWorksheet.Cells(3, oRecordset.Fields.Count)) = vRow
    Application.StatusBar = "Importing data..."
    oWorksheet.[A4].CopyFromRecordset cptConvertFilteredRecordset(oRecordset)
    Application.StatusBar = "Formatting..."
    oWorksheet.[A3].AutoFilter
    oWorksheet.Range(oWorksheet.[A3], oWorksheet.[A3].End(xlToRight)).Font.Bold = True
    oWorksheet.Columns.AutoFit
    oWorksheet.[A4].Select
    oExcel.Visible = True
    oExcel.ScreenUpdating = True
    oExcel.ActiveWindow.FreezePanes = True
    Application.StatusBar = "...done."
  End If
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oWorksheet = Nothing
  oExcel.Calculation = xlCalculationAutomatic
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  oRecordset.Filter = 0
  oRecordset.Close
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptTaskHistory_bas", "cptExportTaskHistory", Err, Erl)
  Resume exit_here
End Sub

Function cptConvertFilteredRecordset(oRecordset As ADODB.Recordset) As Recordset
  Dim oStream As ADODB.Stream
  Set oStream = New ADODB.Stream
  oRecordset.Save oStream, adPersistXML
  Set cptConvertFilteredRecordset = New ADODB.Recordset
  cptConvertFilteredRecordset.Open oStream
End Function
