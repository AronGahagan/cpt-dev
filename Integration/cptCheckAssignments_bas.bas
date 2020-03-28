Attribute VB_Name = "cptCheckAssignments_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptCheckAssignments()
'objects
Dim oSubProject As SubProject
Dim oComment As Excel.Comment
Dim oListObject As ListObject
Dim oWorksheet As Excel.Worksheet
Dim oWorkbook As Excel.Workbook
Dim xlApp As Excel.Application
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim oAssignment As Assignment
Dim oTask As Task
'strings
Dim strMsg As String
'longs
Dim lngLastRow As Long
Dim lngTask As Long
Dim lngTasks As Long
Dim lngCount As Long
Dim lngSigDig As Long
'integers
'doubles
Dim dblTRW As Double
Dim dblTBLW As Double
Dim dblTRC As Double
Dim dblTBLC As Double
Dim dblTRW_T As Double
Dim dblTBLW_T As Double
Dim dblTRC_T As Double
Dim dblTBLC_T As Double
Dim dblARW As Double
Dim dblABLW As Double
Dim dblARC As Double
Dim dblABLC As Double
Dim dblARW_T As Double
Dim dblABLW_T As Double
Dim dblARC_T As Double
Dim dblABLC_T As Double
'booleans
Dim blnBaselined As Boolean
'variants
Dim vCol As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'user input: significant digits
  lngSigDig = Val(InputBox("How many significant digits:", "How precise?", 3))
    
  'create workbook
  Set xlApp = CreateObject("Excel.Application")
  Set oWorkbook = xlApp.Workbooks.Add
  Set oWorksheet = oWorkbook.Sheets(1)
  'todo: perhaps change to UID,CLASS,TASK,TASK_T,ASSIGNMENT,ASSIGNMENT_T,NOTE
  'todo: ...in order to isolate 1 discrepancy per report row
  oWorksheet.[A1:R1] = Array("UID", "TRW", "ARW", "TRW_T", "ARW_T", "TRC", "ARC", "TRC_T", "ARC_T", "TBLW", "ABLW", "TBLW_T", "ABLW_T", "TBLC", "ABLC", "TBLC_T", "ABLC_T", "RESULT")
  
  'get task count
  For Each oTask In ActiveProject.Tasks
    lngTasks = lngTasks + 1
  Next
  
  'account for when no baseline
  blnBaselined = IsDate(ActiveProject.BaselineSavedDate(pjBaseline))
  'todo: compare parent/child dates also summary vs. task; task vs. assignment
  'todo: compare summary rollups
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    oTask.Marked = False
    'capture task totals
    dblTRW = oTask.RemainingWork / 60
    dblTRC = oTask.RemainingCost
    dblTBLW = oTask.BaselineWork / 60
    dblTBLC = Val(oTask.BaselineCost)
    'get task timephased work
    dblTRW_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledWork, pjTimescaleYears)
    For Each TSV In TSVS
      dblTRW_T = dblTRW_T + (Val(TSV.Value) / 60)
      'subtract timephased actual work
      dblTRW_T = dblTRW_T - (Val(oTask.TimeScaleData(TSV.StartDate, TSV.EndDate, pjTaskTimescaledActualWork, pjTimescaleYears).Item(1)) / 60)
    Next TSV
    'get task timephased baseline work
    dblTBLW_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.BaselineStart, oTask.BaselineFinish, pjTaskTimescaledBaselineWork, pjTimescaleYears)
    For Each TSV In TSVS
      dblTBLW_T = dblTBLW_T + (Val(TSV.Value) / 60)
    Next TSV
    'get task timephased remainig cost
    dblTRC_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledCost, pjTimescaleYears)
    For Each TSV In TSVS
      dblTRC_T = dblTRC_T + Val(TSV.Value)
      'subtract timephased actual cost
      dblTRC_T = dblTRC_T - Val(oTask.TimeScaleData(TSV.StartDate, TSV.EndDate, pjTaskTimescaledActualCost, pjTimescaleYears).Item(1))
    Next TSV
    'get task timephased baseline cost
    dblTBLC_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.BaselineStart, oTask.BaselineFinish, pjTaskTimescaledBaselineCost, pjTimescaleYears)
    For Each TSV In TSVS
      dblTBLC_T = dblTBLC_T + Val(TSV.Value)
    Next TSV
    'clear assignment total variables
    dblARW = 0
    dblABLW = 0
    dblARC = 0
    dblABLC = 0
    dblARW_T = 0
    dblABLW_T = 0
    dblARC_T = 0
    dblABLC_T = 0
    'summarize assignment values
    For Each oAssignment In oTask.Assignments
      'capture assignment totals
      dblARW = dblARW + (oAssignment.RemainingWork / 60)
      dblABLW = dblABLW + (Val(oAssignment.BaselineWork) / 60)
      dblARC = dblARC + oAssignment.RemainingCost
      dblABLC = dblABLC + Val(oAssignment.BaselineCost)
      'get timephased remaining assignment work
      Set TSVS = oAssignment.TimeScaleData(oAssignment.Start, oAssignment.Finish, pjAssignmentTimescaledWork, pjTimescaleYears)
      For Each TSV In TSVS
        dblARW_T = dblARW_T + (Val(TSV.Value) / 60)
        'subtract actuals
        dblARW_T = dblARW_T - (Val(oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleYears).Item(1)) / 60)
      Next TSV
      'get timephased assignment baseline work
      Set TSVS = oAssignment.TimeScaleData(oAssignment.BaselineStart, oAssignment.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleYears)
      For Each TSV In TSVS
        dblABLW_T = dblABLW_T + (Val(TSV.Value) / 60)
      Next TSV
      'get timephased assignment remaining cost
      Set TSVS = oAssignment.TimeScaleData(oAssignment.Start, oAssignment.Finish, pjAssignmentTimescaledCost, pjTimescaleYears)
      For Each TSV In TSVS
        dblARC_T = dblARC_T + Val(TSV.Value)
        'subtract actuals
        dblARC_T = dblARC_T - (Val(oAssignment.TimeScaleData(TSV.StartDate, TSV.EndDate, pjAssignmentTimescaledActualCost, pjTimescaleYears).Item(1)))
      Next TSV
      'get timephased assignment baseline cost
      Set TSVS = oAssignment.TimeScaleData(oAssignment.BaselineStart, oAssignment.BaselineFinish, pjAssignmentTimescaledBaselineCost, pjTimescaleYears)
      For Each TSV In TSVS
        dblABLC_T = dblABLC_T + Val(TSV.Value)
      Next TSV
    Next oAssignment
    strMsg = ""
    'TRW,ARW,TRW_T,ARW_T,TRC,ARC,TRC_T,ARC_T,TBLW,ABLW,TBLW_T,ABLW_T,TBLC,ABLC,TBLC_T,ABLC_T
    If Round(dblTRW, lngSigDig) <> Round(dblARW, lngSigDig) Then
      strMsg = strMsg & "Task Remaining Work does not match Assignment Remaining Work." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTBLW, lngSigDig) <> Round(dblABLW, lngSigDig) Then
      strMsg = strMsg & "Task Baseline Work does not match Assignment Baseline Work." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTRC, lngSigDig) <> Round(dblARC, lngSigDig) Then
      strMsg = strMsg & "Task Remaining Cost does not match Assignment Remaining Cost." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTBLC, lngSigDig) <> Round(dblABLC, lngSigDig) Then
      strMsg = strMsg & "Task Baseline Cost does not match Assignment Baseline Cost." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTRW_T, lngSigDig) <> Round(dblARW_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Remaining Work does not match Assignment Timephased Remaining Work." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTBLW_T, lngSigDig) <> Round(dblABLW_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Baseline Work does not match Assignment Timephased Baseline Work." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTRC_T, lngSigDig) <> Round(dblARC_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Remaining Cost does not match Assignment Timephased Remaining Cost." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Round(dblTBLC_T, lngSigDig) <> Round(dblABLC_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Baseline Cost does not match Assignment Timephased Baseline Cost." & Chr(10)
      lngCount = lngCount + 1
    End If
    If Len(strMsg) > 0 Then
      oTask.Marked = True
      'TRW,ARW,TRW_T,ARW_T,TRC,ARC,TRC_T,ARC_T,TBLW,ABLW,TBLW_T,ABLW_T,TBLC,ABLC,TBLC_T,ABLC_T
      lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
      'hack off the last crlf
      strMsg = Left(strMsg, Len(strMsg) - 1)
      oWorksheet.Range(oWorksheet.Cells(lngLastRow, 1), oWorksheet.Cells(lngLastRow, 18)) = Array(oTask.UniqueID, dblTRW, dblARW, dblTRW_T, dblARW_T, dblTRC, dblARC, dblTRC_T, dblARC_T, dblTBLW, dblABLW, dblTBLW_T, dblABLW_T, dblTBLC, dblABLC, dblTBLC_T, dblABLC_T, strMsg)
    End If
next_task:
    'provide user feedback
    lngTask = lngTask + 1
    Application.StatusBar = Format(lngTask, "#,##0") & " of " & Format(lngTasks, "#,##0") & " (" & Format(lngTask / lngTasks, "0%") & ")  |  " & Format(lngCount, "#,##0") & IIf(lngCount = 1, " discrepancy", " discrepancies")
  Next

  Close #lngFile

  If lngCount > 0 Then
    SetAutoFilter "Marked", pjAutoFilterFlagYes
    MsgBox Format(lngCount, "#,##0") & IIf(lngCount = 1, " discrepancy", " discrepancies") & " found.", vbExclamation + vbOKOnly, "cptCheckAssignments"
    Application.StatusBar = "Opening Discrepancy Report..."
    Application.StatusBar = "Formatting Discrepancy Report..."
    xlApp.ActiveWindow.Zoom = 85
    xlApp.ActiveWindow.DisplayGridLines = False
    oWorksheet.[B2].Select
    xlApp.ActiveWindow.FreezePanes = True
    oWorksheet.[B:Q].NumberFormat = "_(* #,##0." & String(lngSigDig, "0") & "_);_(* (#,##0." & String(lngSigDig, "0") & ");_(* ""-""??_);_(@_)"
    Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).Address(True, True), , xlYes)
    oListObject.TableStyle = ""
    cptAddBorders oListObject.DataBodyRange
    'add 4 validation columns =1=2=3=4
    'todo: do not include baseline values when there is no baseline
    oWorksheet.Columns(6).Insert
    oWorksheet.Cells(1, 6).Value = "RW_MATCH"
    oWorksheet.Cells(2, 6).FormulaR1C1 = "=AND(ROUND([@TRW]," & lngSigDig & ")=ROUND([@ARW]," & lngSigDig & "),ROUND([@ARW]," & lngSigDig & ")=ROUND([@[TRW_T]]," & lngSigDig & "),ROUND([@[TRW_T]]," & lngSigDig & ")=ROUND([@[ARW_T]]," & lngSigDig & "))"
    oWorksheet.Columns(11).Insert
    oWorksheet.Cells(1, 11).Value = "RC_MATCH"
    oWorksheet.Cells(2, 11).FormulaR1C1 = "=AND(ROUND([@TRC]," & lngSigDig & ")=ROUND([@ARC]," & lngSigDig & "),ROUND([@ARC]," & lngSigDig & ")=ROUND([@[TRC_T]]," & lngSigDig & "),ROUND([@[TRC_T]]," & lngSigDig & ")=ROUND([@[ARC_T]]," & lngSigDig & "))"
    oWorksheet.Columns(16).Insert
    oWorksheet.Cells(1, 16).Value = "BLW_MATCH"
    oWorksheet.Cells(2, 16).FormulaR1C1 = "=AND(ROUND([@TBLW]," & lngSigDig & ")=ROUND([@ABLW]," & lngSigDig & "),ROUND([@ABLW]," & lngSigDig & ")=ROUND([@[TBLW_T]]," & lngSigDig & "),ROUND([@[TBLW_T]]," & lngSigDig & ")=ROUND([@[ABLW_T]]," & lngSigDig & "))"
    oWorksheet.Columns(21).Insert
    oWorksheet.Cells(1, 21).Value = "BLC_MATCH"
    oWorksheet.Cells(2, 21).FormulaR1C1 = "=AND(ROUND([@TBLC]," & lngSigDig & ")=ROUND([@ABLC]," & lngSigDig & "),ROUND([@ABLC]," & lngSigDig & ")=ROUND([@[TBLC_T]]," & lngSigDig & "),ROUND([@[TBLC_T]]," & lngSigDig & ")=ROUND([@[ABLC_T]]," & lngSigDig & "))"
    oWorksheet.[B3].Select
    oWorksheet.Columns.WrapText = False
    oWorksheet.Columns.AutoFit
    oWorksheet.Columns(22).WrapText = True
    oWorksheet.Columns(22).AutoFit
    oListObject.DataBodyRange.VerticalAlignment = xlCenter
    'add conditional formatting
    For Each vCol In Array(6, 11, 16, 21)
      With oListObject.ListColumns(vCol).DataBodyRange
          .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=FALSE"
          .FormatConditions(.FormatConditions.Count).SetFirstPriority
          .FormatConditions(1).Font.Color = -16383844
          .FormatConditions(1).Font.TintAndShade = 0
        With .FormatConditions(1).Interior
          .PatternColorIndex = xlAutomatic
          .Color = 13551615
          .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
      End With
      With oListObject.ListColumns(vCol).DataBodyRange
          .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=TRUE"
          .FormatConditions(.FormatConditions.Count).SetFirstPriority
          .FormatConditions(1).Font.Color = -16752384
          .FormatConditions(1).Font.TintAndShade = 0
        With .FormatConditions(1).Interior
          .PatternColorIndex = xlAutomatic
          .Color = 13561798
          .TintAndShade = 0
        End With
        .FormatConditions(1).StopIfTrue = False
      End With
      cptAddBorders oWorksheet.Range(oWorksheet.Cells(1, vCol), oWorksheet.Cells(1, vCol).End(xlDown).Offset(0, -4))
    Next vCol
    oListObject.Range.Columns.AutoFit
    'add comments (or entry note) to headers
    With oWorksheet
      Set oComment = .Cells(1, 1).AddComment("Task Unique ID")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 2).AddComment("Task Remaining Work")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 3).AddComment("Assignment Remaining Work")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 4).AddComment("Task Remaining Work (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 5).AddComment("Assignment Remaining Work (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 7).AddComment("Task Remaining Cost")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 8).AddComment("Assignment Remaining Cost")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 9).AddComment("Task Remaining Cost (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 10).AddComment("Assignment Remaining Cost (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 12).AddComment("Task Baseline Work")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 13).AddComment("Assignment Baseline Work")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 14).AddComment("Task Baseline Work (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 15).AddComment("Assignment Baseline Work (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 17).AddComment("Task Baseline Cost")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 18).AddComment("Assignment Baseline Cost")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 19).AddComment("Task Baseline Cost (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
      Set oComment = .Cells(1, 20).AddComment("Assignment Baseline Cost (Timephased)")
      oComment.Shape.Height = .Cells(1, 1).Height * 4
    End With
    'pretty up the header and borders
    oListObject.HeaderRowRange.Font.Bold = True
    oWorksheet.Rows(1).Insert
    With oWorksheet.[B1:F1]
      .Merge True
      .HorizontalAlignment = xlCenter
      .Font.Bold = True
      .Value = "REMAINING WORK"
    End With
    cptAddBorders oWorksheet.[B1:F1]
    With oWorksheet.[G1:K1]
      .Merge True
      .HorizontalAlignment = xlCenter
      .Font.Bold = True
      .Value = "REMAINING COST"
    End With
    cptAddBorders oWorksheet.[G1:K1]
    With oWorksheet.[L1:P1]
      .Merge True
      .HorizontalAlignment = xlCenter
      .Font.Bold = True
      .Value = "BASELINE WORK"
    End With
    cptAddBorders oWorksheet.[L1:P1]
    With oWorksheet.[Q1:U1]
      .Merge True
      .HorizontalAlignment = xlCenter
      .Font.Bold = True
      .Value = "BASELINE COST"
    End With
    cptAddBorders oWorksheet.[Q1:U1]
    'borders around headers
    cptAddBorders oListObject.HeaderRowRange
    'todo: crlf in comments
    '-- replace "." with "."&chr(10)
    '-- reapply autofit column
    'todo: autofit rows
    'todo: vertical align rows
  Else
    oWorkbook.Close False
    xlApp.Quit
    MsgBox "No discrepancies found.", vbInformation + vbOKOnly, "CheckAssignments"
  End If

  Application.StatusBar = "Report complete."

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oSubProject = Nothing
  Set oComment = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  If Not xlApp Is Nothing Then xlApp.Visible = True
  If Not xlApp Is Nothing Then xlApp.WindowState = xlMaximized
  Set xlApp = Nothing
  Set TSV = Nothing
  Set TSVS = Nothing
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCheckAssignments_bas", "cptCheckAssignments", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptAddBorders(ByRef rng As Excel.Range)
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.499984740745262
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
    rng.Borders(xlInsideHorizontal).LineStyle = xlNone

End Sub
