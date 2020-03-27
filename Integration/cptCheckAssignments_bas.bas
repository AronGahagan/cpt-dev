Attribute VB_Name = "cptCheckAssignments_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptCheckAssignments()
'objects
Dim oListObject As ListObject
Dim oWorksheet As Excel.Worksheet
Dim oWorkbook As Excel.Workbook
Dim xlApp As Excel.Application
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim oAssignment As Assignment
Dim oTask As Task
'strings
Dim strFileName As String
Dim strMsg As String
'longs
Dim lngFile As Long
Dim lngCount As Long
Dim lngSigDig As Long
'integers
'doubles
Dim dblTW As Double
Dim dblTBLW As Double
Dim dblTC As Double
Dim dblTBLC As Double
Dim dblTW_T As Double
Dim dblTBLW_T As Double
Dim dblTC_T As Double
Dim dblTBLC_T As Double
Dim dblAW As Double
Dim dblABLW As Double
Dim dblAC As Double
Dim dblABLC As Double
Dim dblAW_T As Double
Dim dblABLW_T As Double
Dim dblAC_T As Double
Dim dblABLC_T As Double
'booleans
'variants
Dim vCol As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'user input: significant digits
  lngSigDig = Val(InputBox("How many significant digits:", "How precise?", 3))
  'create CSV
  strFileName = Environ("USERPROFILE") & "\cptCheckAssignments_" & Format(Now, "yyyy-mm-dd-hh-nn-ss") & ".csv"
  lngFile = FreeFile
  Open strFileName For Output As #lngFile
  'print header
  Print #lngFile, "UID,TW,AW,TW_T,AW_T,TC,AC,TC_T,AC_T,TBLW,ABLW,TBLW_T,ABLW_T,TBLC,ABLC,TBLC_T,ABLC_T,RESULT"
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    oTask.Marked = False
    'capture task totals
    dblTW = oTask.Work / 60
    dblTBLW = oTask.BaselineWork / 60
    dblTC = oTask.Cost
    dblTBLC = Val(oTask.BaselineCost)
    'capture timephased task totals
    dblTW_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledWork, pjTimescaleYears)
    For Each TSV In TSVS
      dblTW_T = dblTW_T + (Val(TSV.Value) / 60)
    Next TSV
    dblTBLW_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledBaselineWork, pjTimescaleYears)
    For Each TSV In TSVS
      dblTBLW_T = dblTBLW_T + (Val(TSV.Value) / 60)
    Next TSV
    dblTC_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledCost, pjTimescaleYears)
    For Each TSV In TSVS
      dblTC_T = dblTC_T + Val(TSV.Value)
    Next TSV
    dblTBLC_T = 0
    Set TSVS = oTask.TimeScaleData(oTask.Start, oTask.Finish, pjTaskTimescaledBaselineCost, pjTimescaleYears)
    For Each TSV In TSVS
      dblTBLC_T = dblTBLC_T + Val(TSV.Value)
    Next TSV
    'clear assignment total variables
    dblAW = 0
    dblABLW = 0
    dblAC = 0
    dblABLC = 0
    dblAW_T = 0
    dblABLW_T = 0
    dblAC_T = 0
    dblABLC_T = 0
    'summarize assignment values
    For Each oAssignment In oTask.Assignments
      dblAW = dblAW + (oAssignment.Work / 60)
      dblABLW = dblABLW + (oAssignment.BaselineWork / 60)
      dblAC = dblAC + oAssignment.Cost
      dblABLC = dblABLC + Val(oAssignment.BaselineCost)
      Set TSVS = oAssignment.TimeScaleData(oTask.Start, oTask.Finish, pjAssignmentTimescaledWork, pjTimescaleYears)
      For Each TSV In TSVS
        dblAW_T = dblAW_T + (Val(TSV.Value) / 60)
      Next TSV
      Set TSVS = oAssignment.TimeScaleData(oTask.Start, oTask.Finish, pjAssignmentTimescaledBaselineWork, pjTimescaleYears)
      For Each TSV In TSVS
        dblABLW_T = dblABLW_T + (Val(TSV.Value) / 60)
      Next TSV
      Set TSVS = oAssignment.TimeScaleData(oTask.Start, oTask.Finish, pjAssignmentTimescaledCost, pjTimescaleYears)
      For Each TSV In TSVS
        dblAC_T = dblAC_T + Val(TSV.Value)
      Next TSV
      Set TSVS = oAssignment.TimeScaleData(oTask.Start, oTask.Finish, pjAssignmentTimescaledBaselineCost, pjTimescaleYears)
      For Each TSV In TSVS
        dblABLC_T = dblABLC_T + Val(TSV.Value)
      Next TSV
    Next oAssignment
    strMsg = ""
    'TW,AW,TW_T,AW_T,TC,AC,TC_T,AC_T,TBLW,ABLW,TBLW_T,ABLW_T,TBLC,ABLC,TBLC_T,ABLC_T
    If Round(dblTW, lngSigDig) <> Round(dblAW, lngSigDig) Then
      strMsg = strMsg & "Task Work does not match Assignment Work." & Chr(10)
    End If
    If Round(dblTBLW, lngSigDig) <> Round(dblABLW, lngSigDig) Then
      strMsg = strMsg & "Task Baseline Work does not match Assignment Baseline Work." & Chr(10)
    End If
    If Round(dblTC, lngSigDig) <> Round(dblAC, lngSigDig) Then
      strMsg = strMsg & "Task Cost does not match Assignment Cost." & Chr(10)
    End If
    If Round(dblTBLC, lngSigDig) <> Round(dblABLC, lngSigDig) Then
      strMsg = strMsg & "Task Baseline Cost does not match Assignment Baseline Cost." & Chr(10)
    End If
    If Round(dblTW_T, lngSigDig) <> Round(dblAW_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Work does not match Assignment Timephased Work." & Chr(10)
    End If
    If Round(dblTBLW_T, lngSigDig) <> Round(dblABLW_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Baseline Work does not match Assignment Timephased Baseline Work." & Chr(10)
    End If
    If Round(dblTC_T, lngSigDig) <> Round(dblAC_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephase Cost does not match Assignment Timephased Cost." & Chr(10)
    End If
    If Round(dblTBLC_T, lngSigDig) <> Round(dblABLC_T, lngSigDig) Then
      strMsg = strMsg & "Task Timephased Baseline Cost does not match Assignment Timephased Baseline Cost." & Chr(10)
    End If
    If Len(strMsg) > 0 Then
      oTask.Marked = True
      lngCount = lngCount + 1
      'TW,AW,TW_T,AW_T,TC,AC,TC_T,AC_T,TBLW,ABLW,TBLW_T,ABLW_T,TBLC,ABLC,TBLC_T,ABLC_T
      Print #lngFile, oTask.UniqueID & "," & dblTW & "," & dblAW & "," & dblTW_T & "," & dblAW_T & "," & dblTC & "," & dblAC & "," & dblTC_T & "," & dblAC_T & "," & dblTBLW & "," & dblABLW & "," & dblTBLW_T & "," & dblABLW_T & "," & dblTBLC & "," & dblABLC & "," & dblTBLC_T & "," & dblABLC_T & "," & strMsg
    End If
next_task:
    'todo: need a progress thingy on the status bar
    'x of y (%) | count discrepancies
  Next

  Close #lngFile

  If lngCount > 0 Then
    SetAutoFilter "Marked", pjAutoFilterFlagYes
    MsgBox Format(lngCount, "#,##0") & IIf(lngCount = 1, " discrepancy", " discrepancies") & " found.", vbExclamation + vbOKOnly, "CheckAssignments"
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set oWorkbook = xlApp.Workbooks.Open(strFileName)
    Set oWorksheet = oWorkbook.Sheets(1)
    xlApp.ActiveWindow.Zoom = 85
    xlApp.WindowState = xlMaximized
    oWorksheet.[B2].Select
    xlApp.ActiveWindow.FreezePanes = True
    oWorksheet.[B:Q].NumberFormat = "_(* #,##0." & String(lngSigDig, "0") & "_);_(* (#,##0." & String(lngSigDig, "0") & ");_(* ""-""??_);_(@_)"
    Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).Address(True, True), , xlYes)
    oListObject.TableStyle = ""
    'add 4 validation columns =1=2=3=4 using EXACT
    oWorksheet.Columns(6).Insert
    oWorksheet.Cells(1, 6).Value = "W_MATCH"
    oWorksheet.Cells(2, 6).FormulaR1C1 = "=EXACT(Round([@TW]," & lngSigDig & ")=Round([@AW]," & lngSigDig & "),Round([@[TW_T]]," & lngSigDig & ")=Round([@[AW_T]]," & lngSigDig & "))"
    oWorksheet.Columns(11).Insert
    oWorksheet.Cells(1, 11).Value = "C_MATCH"
    oWorksheet.Cells(2, 11).FormulaR1C1 = "=EXACT(Round([@TC]," & lngSigDig & ")=Round([@AC]," & lngSigDig & "),Round([@[TC_T]]," & lngSigDig & ")=Round([@[AC_T]]," & lngSigDig & "))"
    oWorksheet.Columns(16).Insert
    oWorksheet.Cells(1, 16).Value = "BLW_MATCH"
    oWorksheet.Cells(2, 16).FormulaR1C1 = "=EXACT(Round([@TBLW]," & lngSigDig & ")=Round([@ABLW]," & lngSigDig & "),Round([@[TBLW_T]]," & lngSigDig & ")=Round([@[ABLW_T]]," & lngSigDig & "))"
    oWorksheet.Columns(21).Insert
    oWorksheet.Cells(1, 21).Value = "BLC_MATCH"
    oWorksheet.Cells(2, 21).FormulaR1C1 = "=EXACT(Round([@TBLC]," & lngSigDig & ")=Round([@ABLC]," & lngSigDig & "),Round([@[TBLC_T]]," & lngSigDig & ")=Round([@[ABLC_T]]," & lngSigDig & "))"
    oWorksheet.Columns.AutoFit
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
      'todo: borders around groups
    Next vCol
    'todo: add comments (or entry note) to headers
  Else
    MsgBox "No discrepancies found.", vbInformation + vbOKOnly, "CheckAssignments"
  End If

exit_here:
  On Error Resume Next
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
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
