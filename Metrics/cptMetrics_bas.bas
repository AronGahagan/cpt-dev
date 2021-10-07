Attribute VB_Name = "cptMetrics_bas"
'<cpt_version>v1.1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptGetBAC()
  MsgBox Format(cptGetMetric("bac"), "#,##0.00h"), vbInformation + vbOKOnly, "Budget at Complete (BAC) - hours"
End Sub

Sub cptGetETC()
  MsgBox Format(cptGetMetric("etc"), "#,##0.00h"), vbInformation + vbOKOnly, "Estimate to Complete (ETC) - hours"
End Sub

Sub cptGetBCWS()

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcws"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Scheduled (BCWS) - hours"
  End If

End Sub

Sub cptGetBCWP()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    MsgBox Format(cptGetMetric("bcwp"), "#,##0.00"), vbInformation + vbOKOnly, "Budgeted Cost of Work Performed (BCWP) - hours"
  End If
  
End Sub

Sub cptGetSPI()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("SPI")
  End If
  
End Sub

Sub cptGetBEI()

  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("BEI")
  End If
  
End Sub

Sub cptGetCEI()
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("CEI")
  End If
  
End Sub

Sub cptGetSV()
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    Exit Sub
  Else
    Call cptGET("SV")
  End If

End Sub

Sub cptGetCPLI()
'objects
Dim oTasks As Tasks
Dim oPred As Task
Dim oTask As Task
'strings
Dim strProgram  As String
Dim strMsg As String
Dim strTitle As String
'longs
Dim lngConstraintType As Long
Dim lngTS As Long
Dim lngCPL As Long
'integers
'doubles
Dim dblCPLI As Double
'booleans
'variants
'dates
Dim dtStart As Date, dtFinish As Date
Dim dtConstraintDate As Date

  strTitle = "Critical Path Length Index (CPLI)"

  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then
    MsgBox "No Target Task selected.", vbExclamation + vbOKOnly, "Oops"
    GoTo exit_here
  End If

  'confirm a single, target oTask is selected
  If oTasks.Count <> 1 Then
    MsgBox "Please select a single, active, and non-summary target Task.", vbExclamation + vbOKOnly, strTitle
    GoTo exit_here
  End If
  
  Set oTask = oTasks(1)
  
  If oTask.Summary Or Not oTask.Active Or oTask.ExternalTask Then
    MsgBox "Please select a single, active, and non-summary target Task.", vbExclamation + vbOKOnly, strTitle
    GoTo exit_here
  End If
  
  strMsg = "TARGET TASK:" & vbCrLf & "UID " & oTask.UniqueID & " - " & oTask.Name & vbCrLf & vbCrLf
  
  'use MFO or MSO constraint
  If oTask.ConstraintType <> pjMFO And oTask.ConstraintType <> pjMSO Then
    strMsg = strMsg & "No MSO/MFO constraint found; temporarily using Deadline..." & vbCrLf
    'if no MFO then use deadline as MFO
    If IsDate(oTask.Deadline) Then
      If IsDate(oTask.ConstraintDate) Then dtConstraintDate = oTask.ConstraintDate
      lngConstraintType = oTask.ConstraintType
      oTask.ConstraintDate = oTask.Deadline
      oTask.ConstraintType = pjMFO
      lngTS = oTask.TotalSlack
      dtFinish = oTask.Finish
      If CLng(dtConstraintDate) > 0 Then oTask.ConstraintDate = dtConstraintDate
      oTask.ConstraintType = lngConstraintType
    Else
      strMsg = strMsg & "No Deadline found; temporarily using Baseline Finish..." & vbCrLf
      If Not IsDate(oTask.BaselineFinish) Then
        strMsg = strMsg & "No Baseline Finish found." & vbCrLf & vbCrLf
        strMsg = strMsg & "In order to calculate the CPLI, the target Task should be (at least temporarily) constrained with a MFO or Deadline." & vbCrLf & vbCrLf
        strMsg = strMsg & "Please constrain the Target Task and try again."
        MsgBox strMsg, vbExclamation + vbOKOnly, strTitle
        GoTo exit_here
      Else
        If IsDate(oTask.ConstraintDate) Then dtConstraintDate = oTask.ConstraintDate
        lngConstraintType = oTask.ConstraintType
        oTask.ConstraintDate = oTask.BaselineFinish
        oTask.ConstraintType = pjMFO
        lngTS = oTask.TotalSlack
        dtFinish = oTask.Finish
        If CLng(dtConstraintDate) > 0 Then oTask.ConstraintDate = dtConstraintDate
        oTask.ConstraintType = lngConstraintType
      End If
    End If
  Else
    lngTS = oTask.TotalSlack
    dtFinish = oTask.Finish
  End If
      
  'use status date if exists
  If IsDate(ActiveProject.StatusDate) Then
    dtStart = ActiveProject.StatusDate
  Else
    dtStart = FormatDateTime(Now(), vbShortDate) & " 08:00 AM"
  End If
  
  If ActiveWindow.ActivePane.View.Screen <> pjGantt Then
    If MsgBox("Cannot use this screen: OK to switch?", vbExclamation + vbYesNo, strTitle) = vbYes Then
      ActiveWindow.TopPane.Activate
      ViewApply "Gantt Chart"
      FilterClear
      GroupClear
      Application.Sort "ID", , , , , , , True
      OptionsViewEx displaysummarytasks:=True, displaynameindent:=True, displayoutlinesymbols:=True
      OutlineShowAllTasks
      EditGoTo oTask.ID
    Else
      GoTo exit_here
    End If
  End If
  
  'use earliest start date
  If oTask Is Nothing Then GoTo exit_here
  If oTask.Summary Then GoTo exit_here
  If Not oTask.Active Then GoTo exit_here
  HighlightDrivingPredecessors Set:=True
  For Each oPred In ActiveProject.Tasks
    If oPred.PathDrivingPredecessor Then
      If IsDate(oPred.ActualStart) Then
        If oPred.Stop < dtStart Then dtStart = oPred.Stop
      Else
        If oPred.Start < dtStart Then dtStart = oPred.Start
      End If
    End If
  Next oPred
  'calculate the CPL
  lngCPL = Application.DateDifference(dtStart, dtFinish)
  'convert values to days
  lngCPL = lngCPL / 480
  lngTS = lngTS / 480
  'notify user
  strMsg = strMsg & vbCrLf & "CPL = Critical Path Length" & vbCrLf
'  strMsg = strMsg & "CPL = Target Finish - Timenow (or CP start)" & vbCrLf
'  strMsg = strMsg & "CPL = " & FormatDateTime(dtFinish, vbShortDate) & " - " & FormatDateTime(dtStart, vbShortDate) & vbCrLf
'  strMsg = strMsg & "CPL = " & lngCPL & vbCrLf
  strMsg = strMsg & "TS = Total Slack" & vbCrLf & vbCrLf
  strMsg = strMsg & "CPLI = ( CPL + TS ) / CPL" & vbCrLf
  strMsg = strMsg & "CPLI = ( " & lngCPL & " + " & lngTS & " ) / " & lngCPL & vbCrLf & vbCrLf
  strMsg = strMsg & "CPLI = " & Round((lngCPL + lngTS) / lngCPL, 3) & vbCrLf & vbCrLf
  strMsg = strMsg & "Note: your CPL may include SCHEDULE MARGIN."
  
  MsgBox strMsg, vbInformation + vbOKOnly, strTitle
  
  dblCPLI = Round((lngCPL + lngTS) / lngCPL, 2)
  strProgram = cptGetProgramAcronym
  cptCaptureMetric strProgram, dtStart, "CPLI", dblCPLI
    
exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Set oPred = Nothing
  Application.CloseUndoTransaction
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetCPLI", Err, Erl)
  Resume exit_here
End Sub

Sub cptGET(strWhat As String)
'objects
Dim oRecordset As ADODB.Recordset
'strings
Dim strMsg As String, strProgram As String
'longs
Dim lngBEI_AF As Long
Dim lngBEI_BF As Long
'integers
'doubles
Dim dblBCWS As Double
Dim dblBCWP As Double
Dim dblResult As Double
'booleans
'variants
'dates
Dim dtStatus As Date, dtPrevious As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'validate tasks exist
  If ActiveProject.Tasks.Count = 0 Then
    MsgBox "This project has no tasks.", vbExclamation + vbOKOnly, "No Tasks"
    GoTo exit_here
  End If
  
  strProgram = cptGetProgramAcronym
  If Len(strProgram) = 0 Then GoTo exit_here
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project requires a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
    
  Select Case strWhat
    Case "BEI"
      lngBEI_BF = CLng(cptGetMetric("bei_bf"))
      If lngBEI_BF = 0 Then
        MsgBox "No baseline finishes found.", vbExclamation + vbOKOnly, "No BEI"
        GoTo exit_here
      End If
      lngBEI_AF = CLng(cptGetMetric("bei_af"))
      strMsg = "BEI = # Actual Finishes / # Planned Finishes" & vbCrLf
      strMsg = strMsg & "BEI = " & Format(lngBEI_AF, "#,##0") & " / " & Format(lngBEI_BF, "#,##0") & vbCrLf & vbCrLf
      strMsg = strMsg & "BEI = " & Round(lngBEI_AF / lngBEI_BF, 2)
      cptCaptureMetric strProgram, dtStatus, "BEI", Round(lngBEI_AF / lngBEI_BF, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Baseline Execution Index (BEI)"
      
    Case "CEI"
      'does cpt-cei.adtg exist?
      If Dir(cptDir & "\settings\cpt-cei.adtg") = vbNullString Then
        MsgBox "No data file found. You must 'Capture Week' on previous period's file before you can run CEI on current period's statused IMS.", vbExclamation + vbOKOnly, "File Not Found"
        GoTo exit_here
      End If
      'get program acronym
      strProgram = cptGetProgramAcronym
      If Len(strProgram) = 0 Then GoTo exit_here
      'connect to data source
      Set oRecordset = CreateObject("ADODB.Recordset")
      'get list of tasks & count
      oRecordset.Open cptDir & "\settings\cpt-cei.adtg"
      dtStatus = ActiveProject.StatusDate
      With oRecordset
        .MoveFirst
        'get most previous week_ending
        dtPrevious = .Fields("STATUS_DATE")
        Do While Not .EOF
          If .Fields("PROJECT") = strProgram Then
            If .Fields("STATUS_DATE") > dtPrevious And .Fields("STATUS_DATE") < dtStatus Then
              dtPrevious = .Fields("STATUS_DATE")
            End If
          End If
          .MoveNext
        Loop
        'test each one to see if complete and get count
        .MoveFirst
        Do While Not .EOF
          If CBool(.Fields("IS_LOE")) Then GoTo next_record
          If .Fields("PROJECT") = strProgram And .Fields("STATUS_DATE") = dtPrevious Then
            If .Fields("TASK_FINISH") > dtPrevious And .Fields("TASK_FINISH") <= dtStatus Then
              Dim lngFF As Long
              lngFF = lngFF + 1
              On Error Resume Next
              Dim oTask As Task
              If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
              Set oTask = ActiveProject.Tasks.UniqueID(.Fields(1))
              If Not oTask Is Nothing Then
                If IsDate(oTask.ActualFinish) Then
                  Dim lngAF As Long
                  lngAF = lngAF + 1
                End If
              End If
            End If
          End If
next_record:
          .MoveNext
        Loop
        'todo: notify user; prompt for list of FFs?
        strMsg = "CEI = Tasks completed in current period / Tasks forecasted to complete in current period" & vbCrLf & vbCrLf
        strMsg = strMsg & "CEI = " & lngAF & " / " & lngFF & vbCrLf
        strMsg = strMsg & "CEI = " & Round(lngAF / lngFF, 2) & vbCrLf & vbCrLf
        strMsg = strMsg & "- Does not include LOE tasks." & vbCrLf
        strMsg = strMsg & "- Does not include tasks completed in current period but not forecasted to complete in current period." & vbCrLf
        strMsg = strMsg & "- See NDIA Predictive Measures Guide for more information."
        Call cptCaptureMetric(strProgram, dtStatus, "CEI", Round(lngAF / lngFF, 2))
        MsgBox strMsg, vbInformation + vbOKOnly, "Current Execution Index"
        .Close
      End With
      
    Case "SPI"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      If dblBCWS = 0 Then
        MsgBox "No BCWS found.", vbExclamation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
        GoTo exit_here
      End If
      strMsg = "SPI = BCWP / BCWS" & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP, "#,##0h") & " / " & Format(dblBCWS, "#,##0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "SPI = ~" & Round(dblBCWP / dblBCWS, 2) '& vbCrLf & vbCrLf
      'strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      cptCaptureMetric strProgram, dtStatus, "SPI", Round(dblBCWP / dblBCWS, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Performance Index (SPI) - Hours"
      
    Case "SV"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      If dblBCWS = 0 Then
        MsgBox "No BCWS found.", vbExclamation + vbOKOnly, "Schedule Variance (SV) - Hours"
        GoTo exit_here
      End If
      strMsg = strMsg & "Schedule Variance (SV)" & vbCrLf
      strMsg = strMsg & "SV = BCWP - BCWS" & vbCrLf
      strMsg = strMsg & "SV = " & Format(dblBCWP, "#,##0h") & " - " & Format(dblBCWS, "#,##0h") & vbCrLf
      strMsg = strMsg & "SV = ~" & Format(dblBCWP - dblBCWS, "#,##0.0h") & vbCrLf & vbCrLf
      strMsg = strMsg & "Schedule Variance % (SV%)" & vbCrLf
      strMsg = strMsg & "SV% = ( SV / BCWS ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = ( " & Format((dblBCWP - dblBCWS), "#,##0.0h") & " / " & Format(dblBCWS, "#,##0.0h") & " ) * 100" & vbCrLf
      strMsg = strMsg & "SV% = " & Format(((dblBCWP - dblBCWS) / dblBCWS), "0.00%") '& vbCrLf & vbCrLf
      'strMsg = strMsg & "(Assumes EV% in Physical % Complete.)"
      cptCaptureMetric strProgram, dtStatus, "SV", Round((dblBCWP - dblBCWS) / dblBCWS, 2)
      MsgBox strMsg, vbInformation + vbOKOnly, "Schedule Variance (SV) - Hours"
      
    Case "es" 'earned schedule
      cptGetEarnedSchedule
    
  End Select
  
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMerics_Bas", "cptGet", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetHitTask()
'objects
Dim oTask As Task
'strings
Dim strMsg As String
'longs
Dim lngAF As Long
Dim lngBLF As Long
'integers
'doubles
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  'find it
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If IsDate(oTask.BaselineFinish) Then
      'was task baselined to finish before status date?
      If oTask.BaselineFinish <= dtStatus Then
        lngBLF = lngBLF + 1
        'did it?
        If IsDate(oTask.ActualFinish) Then
          If oTask.ActualFinish <= oTask.BaselineFinish Then
            lngAF = lngAF + 1
          End If
        End If
      End If
    End If
next_task:
  Next oTask

  strMsg = "BF = # Tasks Baselined to Finish ON or before Status Date" & vbCrLf
  strMsg = strMsg & "AF = # BF that Actually Finished ON or before Baseline Finish" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = (AF / BF) / 100" & vbCrLf
  strMsg = strMsg & "Hit Task % = (" & Format(lngAF, "#,##0") & " / " & Format(lngBLF, "#,##0") & ") / 100" & vbCrLf & vbCrLf
  strMsg = strMsg & "Hit Task % = " & Format((lngAF / lngBLF), "0%")
  MsgBox strMsg, vbInformation + vbOKOnly, "Hit Task %"

exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetHitTask", Err, Erl)
  Resume exit_here
End Sub

Function cptGetMetric(strGet As String) As Double
'todo: no screen changes!
'objects
Dim oAssignment As Assignment
Dim tsv As TimeScaleValue
Dim tsvs As TimeScaleValues
Dim oTasks As Tasks
Dim oTask As Task
'strings
Dim strLOE As String
'longs
Dim lngLOEField As Long
Dim lngEVP As Long
Dim lngYears As Long
'integers
'doubles
Dim dblResult As Double
'booleans
'variants
'dates
Dim dtStatus As Date

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  lngYears = Year(ActiveProject.ProjectFinish) - Year(ActiveProject.ProjectStart) + 1
  
  'confirm status date
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
    GoTo exit_here
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  cptSpeed True
  ActiveWindow.TopPane.Activate
  FilterClear
  GroupClear
  OptionsViewEx displaysummarytasks:=True, displaynameindent:=True
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    Sort "ID", , , , , , False, True
    OutlineShowAllTasks
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  SelectAll
  Set oTasks = ActiveSelection.Tasks
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.Summary Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.BaselineWork > 0 Then 'idea here was to limit tasks to PMB tasks only
                                  'but won't work for non-resource loaded schedules
      Select Case strGet
        Case "bac"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + (oAssignment.BaselineWork / 60)
            End If
          Next oAssignment
          
        Case "etc"
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              dblResult = dblResult + (oAssignment.RemainingWork / 60)
            End If
          Next oAssignment

          
        Case "bcws"
          If oTask.BaselineStart < dtStatus Then
            For Each oAssignment In oTask.Assignments
              If oAssignment.ResourceType = pjResourceTypeWork Then
                Set tsvs = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                For Each tsv In tsvs
                  dblResult = dblResult + (IIf(tsv.Value = "", 0, tsv.Value) / 60)
                Next
              End If
            Next oAssignment
          End If
          
        Case "bcwp"
          
          If Not cptMetricsSettingsExist Then
            cptShowMetricsSettings_frm True
            If Not cptMetricsSettingsExit Then
              cptGetMetric = 0
              GoTo exit_here
            End If
          End If
        
          lngEVP = CLng(cptGetSetting("Metrics", "cboEVP"))
          lngLOEField = CLng(cptGetSetting("Metrics", "cboLOEField"))
          strLOE = cptGetSetting("Metrics", "txtLOE")
          
          For Each oAssignment In oTask.Assignments
            If oAssignment.ResourceType = pjResourceTypeWork Then
              If oTask.GetField(lngLOEField) = strLOE Then
                If oTask.BaselineStart < dtStatus Then
                  Set tsvs = oAssignment.TimeScaleData(oTask.BaselineStart, dtStatus, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
                  For Each tsv In tsvs
                    dblResult = dblResult + (IIf(tsv.Value = "", 0, tsv.Value) / 60)
                  Next
                End If
              Else
                dblResult = dblResult + ((oAssignment.BaselineWork / 60) * (CLng(cptRegEx(oTask.GetField(lngEVP), "[0-9]*")) / 100))
              End If
            End If
          Next oAssignment
      End Select
    End If 'bac>0
    Select Case strGet
    
      Case "bei_bf"
        dblResult = dblResult + IIf(oTask.BaselineFinish <= dtStatus, 1, 0)
          
      Case "bei_af"
        dblResult = dblResult + IIf(oTask.ActualFinish <= dtStatus, 1, 0)

    End Select
next_task:
    Application.StatusBar = "Calculating " & UCase(strGet) & "..."
  Next

  cptGetMetric = dblResult

exit_here:
  On Error Resume Next
  Set oAssignment = Nothing
  Application.StatusBar = ""
  cptSpeed False
  Set tsv = Nothing
  Set tsvs = Nothing
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Function
err_here:
  'Debug.Print Task.UniqueID & ": " & Task.Name
  Call cptHandleErr("cptMetrics_bas", "cptGetMetric", Err, Erl)
  Resume exit_here

End Function

Sub cptShowMetricsSettings_frm(Optional blnModal As Boolean = False)
  'objects
  'strings
  Dim strCustomName As String
  Dim strLOE As String
  Dim strLOEField As String
  Dim strEVP As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptMetricsSettings_frm
  
    .cboEVP.Clear
    .cboEVP.AddItem
    .cboEVP.List(.cboEVP.ListCount - 1, 0) = FieldNameToFieldConstant("Physical % Complete")
    .cboEVP.List(.cboEVP.ListCount - 1, 1) = "Physical % Complete"
    For lngItem = 1 To 20
      .cboEVP.AddItem
      .cboEVP.List(.cboEVP.ListCount - 1, 0) = FieldNameToFieldConstant("Number" & lngItem)
      .cboEVP.List(.cboEVP.ListCount - 1, 1) = "Number" & lngItem
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant("Number" & lngItem))
      If Len(strCustomName) > 0 Then
        .cboEVP.List(.cboEVP.ListCount - 1, 1) = strCustomName & " (Number" & lngItem & ")"
      End If
    Next lngItem
    
    .cboLOEField.Clear
    For lngItem = 1 To 30
      .cboLOEField.AddItem
      .cboLOEField.List(.cboLOEField.ListCount - 1, 0) = FieldNameToFieldConstant("Text" & lngItem)
      .cboLOEField.List(.cboLOEField.ListCount - 1, 1) = "Text" & lngItem
      strCustomName = CustomFieldGetName(FieldNameToFieldConstant("Text" & lngItem))
      If Len(strCustomName) > 0 Then
        .cboLOEField.List(.cboLOEField.ListCount - 1, 1) = strCustomName & " (Text" & lngItem & ")"
      End If
    Next lngItem
    
    strEVP = cptGetSetting("Metrics", "cboEVP")
    If Len(strEVP) > 0 Then .cboEVP.Value = CLng(strEVP)
    strLOEField = cptGetSetting("Metrics", "cboLOEField")
    If Len(strLOEField) > 0 Then .cboLOEField.Value = CLng(strLOEField)
    strLOE = cptGetSetting("Metrics", "txtLOE")
    If Len(strLOE) > 0 Then .txtLOE = strLOE
    If blnModal Then
      .Show
    Else
      .Show False
    End If
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  'Call HandleErr("cptMetrics_bas", "cptShowMetricsSettings_frm", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Function cptMetricsSettingsExist() As Boolean
  'objects
  'strings
  Dim strLOE As String
  Dim strLOEField As String
  Dim strEVP As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
    
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strEVP = cptGetSetting("Metrics", "cboEVP")
  strLOEField = cptGetSetting("Metrics", "cboLOEField")
  strLOE = cptGetSetting("Metrics", "txtLOE")
  
  If Len(strEVP) = 0 Or Len(strLOEField) = 0 Or Len(strLOE) = 0 Then
    cptMetricsSettingsExist = False
  Else
    cptMetricsSettingsExist = True
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptMetricsSettingsExist", Err, Erl)
  Resume exit_here
End Function


Sub cptCaptureWeek()
  'objects
  Dim oTasks As Tasks
  Dim oTask As Task
  Dim rst As ADODB.Recordset
  'strings
  Dim strLOE As String
  Dim strEVT As String
  Dim strProject As String
  Dim strFile As String
  Dim strDir As String
  'longs
  Dim lngEVT As Long
  Dim lngTasks As Long
  Dim lngTask As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'ensure program acronym
  strProject = cptGetProgramAcronym
  If Len(strProject) = 0 Then
    MsgBox "Program Acronym is required for this feature.", vbExclamation + vbOKOnly, "Program Acronym Needed"
    GoTo exit_here
  End If
    
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings required."
      Exit Sub
    End If
  End If
    
  Set rst = CreateObject("ADODB.Recordset")
  strFile = cptDir & "\settings\cpt-cei.adtg"
  If Dir(strFile) = vbNullString Then
    rst.Fields.Append "PROJECT", adVarChar, 25      '0
    rst.Fields.Append "TASK_UID", adInteger         '1
    rst.Fields.Append "TASK_NAME", adVarChar, 255   '2
    rst.Fields.Append "IS_LOE", adInteger           '3
    rst.Fields.Append "TASK_BLS", adDate            '4
    rst.Fields.Append "TASK_BLD", adInteger         '5
    rst.Fields.Append "TASK_BLF", adDate            '6
    rst.Fields.Append "TASK_AS", adDate             '7
    rst.Fields.Append "TASK_AD", adInteger          '8
    rst.Fields.Append "TASK_AF", adDate             '9
    rst.Fields.Append "TASK_START", adDate          '10
    rst.Fields.Append "TASK_RD", adInteger          '11
    rst.Fields.Append "TASK_FINISH", adDate         '12
    rst.Fields.Append "STATUS_DATE", adDate         '13
    rst.Open
  Else
    rst.Open strFile
  End If
  
  dtStatus = ActiveProject.StatusDate
  If rst.RecordCount > 0 Then
    rst.MoveFirst
    rst.Filter = "STATUS_DATE=#" & FormatDateTime(dtStatus, vbGeneralDate) & "# AND PROJECT='" & strProject & "'"
    If Not rst.EOF Then
      If MsgBox("Status Already Imported for WE " & FormatDateTime(dtStatus, vbShortDate) & "." & vbCrLf & vbCrLf & "Overwrite it?", vbExclamation + vbYesNo, "Overwrite?") = vbYes Then
        rst.MoveFirst
        Do While Not rst.EOF
          If rst("PROJECT") = strProject And rst("STATUS_DATE") = FormatDateTime(dtStatus, vbGeneralDate) Then rst.Delete adAffectCurrent
          rst.MoveNext
        Loop
      End If
    End If
    rst.Filter = 0
  End If
  
  strEVT = cptGetSetting("Metrics", "cboLOEField")
  If Len(strEVT) > 0 Then
    lngEVT = CLng(strEVT)
  Else
    MsgBox "Error retrieving setting for Metrics.cboLOEField. Cannot proceed.", vbExclamation + vbOKOnly, "Error"
    GoTo exit_here
  End If
  strLOE = cptGetSetting("Metrics", "txtLOE")
  If Len(strLOE) = 0 Then
    MsgBox "Error retrieving setting for Metrics.txtLOE. Cannot proceed.", vbExclamation + vbOKOnly, "Error"
    GoTo exit_here
  End If
  
  Set oTasks = ActiveProject.Tasks
  lngTasks = oTasks.Count
  'include all discrete, LOE, milestones, and all SVTs
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task 'skip blank lines
    If Not oTask.Active Then GoTo next_task 'skip inactive
    If oTask.ExternalTask Then GoTo next_task 'skip external
    If oTask.Summary Then GoTo next_task 'skip summaries
    'If oTask.Milestone Then GoTo next_task 'skip milestones
    If oTask.Resources.Count > 0 Or InStr(oTask.Name, "SVT") > 0 Then
      rst.AddNew
      rst(0) = strProject
      rst(1) = oTask.UniqueID
      rst(2) = oTask.Name
      rst(3) = IIf(oTask.GetField(lngEVT) = strLOE, 1, 0)
      If IsDate(oTask.BaselineStart) Then
        rst(4) = FormatDateTime(oTask.BaselineStart, vbGeneralDate)
        rst(5) = Round(oTask.BaselineDuration / (60 * 8), 0)
      End If
      If IsDate(oTask.BaselineFinish) Then
        rst(6) = FormatDateTime(oTask.BaselineFinish, vbGeneralDate)
      End If
      If IsDate(oTask.ActualStart) Then
        rst(7) = FormatDateTime(oTask.ActualStart, vbGeneralDate)
        rst(8) = Round(oTask.ActualDuration / (60 * 8), 0)
      End If
      If IsDate(oTask.ActualFinish) Then
        rst(9) = FormatDateTime(oTask.ActualFinish, vbGeneralDate)
      End If
      rst(10) = FormatDateTime(oTask.Start, vbGeneralDate)
      rst(11) = Round(oTask.RemainingDuration / (60 * 8), 0)
      rst(12) = FormatDateTime(oTask.Finish, vbGeneralDate)
      rst(13) = FormatDateTime(ActiveProject.StatusDate, vbGeneralDate)
      rst.Update
    End If
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = lngTask & " / " & lngTasks & " (" & Format(lngTask / lngTasks, "0%")
  Next oTask
  
  rst.Save strFile, adPersistADTG
  rst.Close
  Application.StatusBar = "Complete."
  MsgBox "Current Schedule as of " & FormatDateTime(ActiveProject.StatusDate, vbShortDate) & " captured.", vbInformation + vbOKOnly, "Complete"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oTasks = Nothing
  Set oTask = Nothing
  If rst.State = 1 Then rst.Close
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("focptMetrics_bas", "cptCaptureWeek", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptLateStartsFinishes()
  'objects
  Dim oSeries As Excel.Series
  Dim oChart As Excel.ChartObject
  Dim oShape As Excel.Shape
  Dim oOutlook As Outlook.Application
  Dim oMailItem As Outlook.MailItem
  Dim oDocument As Word.Document
  Dim oWord As Word.Application
  Dim oSelection As Word.Selection
  Dim oEmailTemplate As Word.Template
  Dim oWorksheet As Excel.Worksheet
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oListObject As Excel.ListObject
  Dim oRange As Excel.Range
  Dim oCell As Excel.Range
  Dim oAssignment As MSProject.Assignment
  Dim oTask As Task
  'strings
  Dim strSummary As String
  Dim strHeaders As String
  Dim strMyHeaders As String
  Dim strLOE As String
  Dim strLOEField As String
  Dim strCC As String
  Dim strTo As String
  Dim strProject As String
  Dim strFile As String
  Dim strDir As String
  'longs
  Dim lngMyHeaders As Long
  Dim lngLastCol As Long
  Dim lngResponse As Long
  Dim lngLOEField As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngForecastCount As Long
  Dim lngBaselineCount As Long
  Dim lngLastRow As Long
  Dim lngCol As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  Dim vRow As Variant
  Dim vMyHeader As Variant
  Dim vCol As Variant
  'dates
  Dim dtDate As Date
  Dim dtMax As Date
  Dim dtMin As Date
  Dim dtStatus As Date
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  dtStatus = ActiveProject.StatusDate
    
  strProject = cptGetProgramAcronym
  
  If Not cptMetricsSettingsExist Then
    Call cptShowMetricsSettings_frm(True)
    If Not cptMetricsSettingsExist Then
      MsgBox "No settings saved. Cannot proceed.", vbExclamation + vbOKOnly, "Settings Required"
      GoTo exit_here
    End If
  End If
  
try_again:
  'get other fields
  strMyHeaders = cptGetSetting("Metrics", "strMyHeaders")
  If Len(strMyHeaders) = 0 Then strMyHeaders = "CAM,WPCN,WPM,"
  vResponse = InputBox("At least one custom field is required." & vbCrLf & vbCrLf & "Enter a comma-separated list (BEI will be grouped by first item):", "Late Starts and Finishes", strMyHeaders)
  If StrPtr(vResponse) = 0 Then 'user hit cancel
    GoTo exit_here
  ElseIf vResponse = "" Or Len(Replace(vResponse, ",", "")) = 0 Then 'user entered zer-value
    'nothing selected
    If MsgBox("You must select at least one custom field. Try again?", vbQuestion + vbYesNo, "Field Required") = vbYes Then
      GoTo try_again
    Else
      GoTo exit_here
    End If
  ElseIf Len(vResponse) > 0 Then
    strMyHeaders = CStr(vResponse)
  End If
  
  If Right(Trim(strMyHeaders), 1) <> "," And Len(strMyHeaders) > 0 Then strMyHeaders = Trim(strMyHeaders) & ","
  'validate strMyHeaders
  On Error Resume Next
  For Each vMyHeader In Split(strMyHeaders, ",")
    If UBound(Split(stryMyHeaders, ",")) = -1 And vMyHeader = "" Then
      If MsgBox("You must select at least one custom field. Try again?", vbQuestion + vbYesNo, "Field Required") = vbYes Then
        GoTo try_again
      Else
        GoTo exit_here
      End If
    End If
    If vMyHeader = "" Then Exit For
    Debug.Print FieldNameToFieldConstant(vMyHeader)
    If Err.Number > 0 Then
      lngResponse = MsgBox("Custom Field '" & vMyHeader & "' not found!" & vbCrLf & vbCrLf & "OK = skip; Cancel = try again", vbExclamation + vbOKCancel, "Invalid Field")
      If lngResponse = vbCancel Then
        Err.Clear
        GoTo try_again
      Else
        Err.Clear
        strMyHeaders = Replace(strMyHeaders, vMyHeader & ",", "")
      End If
    End If
  Next vMyHeader
  
  'get excel
  On Error Resume Next
  'Set oExcel = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
  oExcel.Visible = False
  Set oWorkbook = oExcel.Workbooks.Add
  oExcel.Calculation = xlCalculationManual
  oExcel.ScreenUpdating = False
  Set oWorksheet = oWorkbook.Sheets(1)
  oWorksheet.Name = "DETAILS"
  cptSaveSetting "Metrics", "strMyHeaders", strMyHeaders
  
  strLOEField = cptGetSetting("Metrics", "cboLOEField")
  If Len(strLOEField) > 0 Then
    lngLOEField = CLng(strLOEField)
  Else
    MsgBox "Error retrieving setting Metrics.cboLOEField. Cannot proceed.", vbExclamation + vbOKOnly, "Settings"
    GoTo exit_here
  End If
  strLOE = cptGetSetting("Metrics", "txtLOE")
  If Len(strLOE) = 0 Then
    MsgBox "Error retrieving setting Metrics.strLOE. Cannot proceed.", vbExclamation + vbOKOnly, "Settings"
    GoTo exit_here
  End If
  
  strHeaders = "UID,"
  strHeaders = strHeaders & strMyHeaders
  strHeaders = strHeaders & "NAME,TOTAL SLACK,REMAINING DURATION,REMAINING WORK,BASELINE START,START VARIANCE,ACTUAL START,START,BASELINE FINISH,FINISH VARIANCE,ACTUAL FINISH,FINISH"
  
  oWorksheet.Range(oWorksheet.Cells(1, 1), oWorksheet.Cells(1, 1).Offset(0, UBound(Split(strHeaders, ",")))) = Split(strHeaders, ",")
  lngLastCol = oWorksheet.[A1].End(xlToRight).Column
  lngTasks = ActiveProject.Tasks.Count
  
  For Each oTask In ActiveProject.Tasks
    If Not oTask Is Nothing Then
      oTask.Marked = False
      'skip inactive tasks
      If Not oTask.Active Then GoTo next_task
      'skip summaries
      If oTask.Summary Then GoTo next_task
      'only check for tasks with assignments
      If oTask.Resources.Count = 0 Then GoTo next_task
      'only check for discrete tasks
      If oTask.GetField(lngLOEField) = "A" Then GoTo next_task
      'skip unassigned (currently material/odc/tvl)
      'If oTask.GetField(FieldNameToFieldConstant("WPM")) = "" Then GoTo next_task
      'only report early/late starts/finishes
      If oTask.StartVariance <> 0 Or oTask.FinishVariance <> 0 Then
        lngLastRow = oWorksheet.Cells(1048576, 1).End(xlUp).Row + 1
        vRow = oWorksheet.Range(oWorksheet.Cells(lngLastRow, 1), oWorksheet.Cells(lngLastRow, lngLastCol))
        
        vRow(1, 1) = oTask.UniqueID
        lngCol = 1
        For Each vMyHeader In Split(strMyHeaders, ",")
          If vMyHeader = "" Then Exit For
          lngCol = lngCol + 1
          vRow(1, lngCol) = oTask.GetField(FieldNameToFieldConstant(vMyHeader))
        Next vMyHeader
        lngMyHeaders = UBound(Split(strMyHeaders, ","))
        vRow(1, 2 + lngMyHeaders) = oTask.Name
        vRow(1, 3 + lngMyHeaders) = Round(oTask.TotalSlack / (8 * 60), 0)
        vRow(1, 4 + lngMyHeaders) = oTask.RemainingDuration / (8 * 60)
        vRow(1, 5 + lngMyHeaders) = Round(oTask.RemainingWork / 60, 0)
        
        vRow(1, 6 + lngMyHeaders) = FormatDateTime(oTask.BaselineStart, vbShortDate)
        vRow(1, 7 + lngMyHeaders) = Round(oTask.StartVariance / (8 * 60), 0)
        If IsDate(oTask.ActualStart) Then
          vRow(1, 8 + lngMyHeaders) = FormatDateTime(oTask.ActualStart, vbShortDate)
        End If
        vRow(1, 9 + lngMyHeaders) = FormatDateTime(oTask.Start, vbShortDate)
        
        vRow(1, 10 + lngMyHeaders) = FormatDateTime(oTask.BaselineFinish, vbShortDate)
        vRow(1, 11 + lngMyHeaders) = Round(oTask.FinishVariance / (8 * 60), 0)
        If IsDate(oTask.ActualFinish) Then
          vRow(1, 12 + lngMyHeaders) = FormatDateTime(oTask.ActualFinish, vbShortDate)
        End If
        vRow(1, 13 + lngMyHeaders) = FormatDateTime(oTask.Finish, vbShortDate)
        
        oWorksheet.Range(oWorksheet.Cells(lngLastRow, 1), oWorksheet.Cells(lngLastRow, lngLastCol)) = vRow
      End If
    End If
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Exporting BEI... " & Format(lngTask, "#,##0") & " / " & Format(lngTasks, "#,##0") & " (" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask

  Application.StatusBar = "Analyzing..."

  oWorksheet.Cells(1, oWorksheet.Rows(1).Find("START", lookat:=xlWhole).Column).Value = "CURRENT START"
  oWorksheet.Cells(1, oWorksheet.Rows(1).Find("FINISH", lookat:=xlWhole).Column).Value = "CURRENT FINISH"
  oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True

  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True

  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)).Address, , xlYes)
  
  oListObject.HeaderRowRange.WrapText = True
  oListObject.TableStyle = ""
  oWorksheet.Columns.AutoFit
  oListObject.ListColumns("TOTAL SLACK").Range.ColumnWidth = 10
  For lngCol = 4 + lngMyHeaders To lngLastCol
    oListObject.ListColumns(lngCol).Range.ColumnWidth = 12
  Next lngCol
  oListObject.HeaderRowRange.EntireRow.AutoFit
  
  'create summary worksheet
  Set oWorksheet = oWorkbook.Sheets.Add(oWorkbook.Sheets(1))
  oWorksheet.Name = "SUMMARY"
  oWorksheet.[A1] = strProject & " IMS - Early/Late Starts/Finishes"
  oWorksheet.[A1].Font.Bold = True
  oWorksheet.[A1].Font.Size = 14
  oWorksheet.[A2].Value = FormatDateTime(dtStatus, vbShortDate)
  oWorksheet.Names.Add "STATUS_DATE", oWorksheet.[A2].Address
  
  'get field to summarize by
  strSummary = Split(strMyHeaders, ",")(0)
  oListObject.ListColumns(strSummary).Range.Copy oWorksheet.[A5]
  oWorksheet.Range(oWorksheet.[A6], oWorksheet.[A1048576]).RemoveDuplicates Columns:=1, Header:=xlNo
  oWorksheet.Sort.SortFields.Clear
  oWorksheet.Sort.SortFields.Add2 key:=oWorksheet.Range(oWorksheet.[A6], oWorksheet.[A6].End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With oWorksheet.Sort
    .SetRange oWorksheet.Range(oWorksheet.[A6], oWorksheet.[A1048576].End(xlUp))
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
  End With
  
  'todo: nuance for Critical/Driving?
  oWorksheet.[B4].Value = "ACTUAL"
  oWorksheet.[B4:H4].Merge True
  oWorksheet.[B4:H4].HorizontalAlignment = xlCenter
  oWorksheet.[B4:H4].Font.Bold = True
  'todo: interior
  oWorksheet.[B5:H5] = Array("ES", "EF", "LS", "LF", "# BLF", "# AF", "BEI (Finishes)")
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A5].End(xlToRight), oWorksheet.[A5].End(xlDown)), , xlYes)
  oListObject.TableStyle = "TableStyleMedium2"
  oListObject.Name = "BEI"
  
  'ACTUAL
  oListObject.ListColumns("ES").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[START VARIANCE],""<0"",Table1[ACTUAL START],""<>"")"
  oListObject.ListColumns("EF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[FINISH VARIANCE],""<0"",Table1[ACTUAL FINISH],""<>"")"
  oListObject.ListColumns("LS").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[START VARIANCE],"">0"",Table1[ACTUAL START],""<>"")"
  oListObject.ListColumns("LF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[FINISH VARIANCE],"">0"",Table1[ACTUAL FINISH],""<>"")"
  oListObject.ListColumns("# BLF").DataBodyRange.FormulaR1C1 = "=COUNTIFS(Table1[" & strSummary & "],[@" & strSummary & "],Table1[BASELINE FINISH],""<=" & Format(dtStatus, "mm/dd/yyyy") & """)"
  oListObject.ListColumns("# AF").DataBodyRange.FormulaR1C1 = "=COUNTIFS(Table1[" & strSummary & "],[@" & strSummary & "],Table1[ACTUAL FINISH],""<>"")"
  oListObject.ListColumns("BEI (Finishes)").DataBodyRange.FormulaR1C1 = "=[@['# AF]]/IF([@['# BLF]]=0,1,[@['# BLF]])"
  oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Style = "Comma"
  oListObject.ShowTotals = True
  oListObject.ListColumns("ES").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("EF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("LS").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("LF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("# BLF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.ListColumns("# AF").TotalsCalculation = xlTotalsCalculationSum
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Column).FormulaR1C1 = "=BEI[[#Totals],['# AF]]/BEI[[#Totals],['# BLF]]"
  oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Style = "Comma"
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("BEI (Finishes)").DataBodyRange.Column).Style = "Comma"

  'PROJECTED
  lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 2
  oWorksheet.Cells(lngLastRow, 2).Value = "PROJECTED"
  oWorksheet.Range(oWorksheet.Cells(lngLastRow, 2), oWorksheet.Cells(lngLastRow, 2).Offset(0, 5)).Merge True
  oWorksheet.Cells(lngLastRow, 2).HorizontalAlignment = xlCenter
  oWorksheet.Cells(lngLastRow, 2).Font.Bold = True
  oListObject.Range.Copy oWorksheet.Cells(oWorksheet.[A1048576].End(xlUp).Row + 3, 1)
  Set oListObject = oWorksheet.ListObjects(2)
  oListObject.Name = "PROJECTED"
  oListObject.ListColumns("ES").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[START VARIANCE],""<0"",Table1[ACTUAL START],""="")"
  oListObject.ListColumns("EF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[FINISH VARIANCE],""<0"",Table1[ACTUAL FINISH],""="")"
  oListObject.ListColumns("LS").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[START VARIANCE],"">0"",Table1[ACTUAL START],""="")"
  oListObject.ListColumns("LF").DataBodyRange.Formula2R1C1 = "=COUNTIFS(Table1[" & strSummary & "],RC1,Table1[FINISH VARIANCE],"">0"",Table1[ACTUAL FINISH],""="")"
  oListObject.ListColumns("# BLF").Name = "# TOTAL"
  oListObject.ListColumns("# TOTAL").DataBodyRange.FormulaR1C1 = "=COUNTIFS(Table1[" & strSummary & "],[@" & strSummary & "])"
  oListObject.ListColumns("# AF").Name = "% TOTAL"
  oListObject.ListColumns("% TOTAL").DataBodyRange.FormulaR1C1 = "=[@[LF]]/IF([@['# TOTAL]]=0,1,[@['# TOTAL]])"
  oListObject.ListColumns("% TOTAL").DataBodyRange.Style = "Comma"
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("% TOTAL").DataBodyRange.Column).FormulaR1C1 = "=PROJECTED[[#Totals],[LF]]/IF(PROJECTED[[#Totals],['# TOTAL]]=0,1,PROJECTED[[#Totals],['# TOTAL]])"
  oListObject.TotalsRowRange(ColumnIndex:=oListObject.ListColumns("% TOTAL").DataBodyRange.Column).Style = "Comma"
  oListObject.ListColumns("BEI (Finishes)").Delete
    
  oExcel.ActiveWindow.DisplayGridLines = False
  oExcel.ActiveWindow.Zoom = 85
  oListObject.Range.Columns.AutoFit
  
  'week,BLF,AF,CF (BEI/S-chart)
  'get earliest start and latest finish
  Set oListObject = oWorkbook.Sheets("DETAILS").ListObjects(1)
  dtMin = oExcel.WorksheetFunction.Min(oListObject.ListColumns("Baseline Start").DataBodyRange)
  dtMin = oExcel.WorksheetFunction.Min(dtMin, oListObject.ListColumns("Actual Start").DataBodyRange)
  dtMin = oExcel.WorksheetFunction.Min(dtMin, oListObject.ListColumns("Current Start").DataBodyRange)
  'convert to WE Friday
  dtMin = DateAdd("d", 6 - Weekday(dtMin), dtMin)
  dtMax = oExcel.WorksheetFunction.Max(oListObject.ListColumns("Baseline Finish").DataBodyRange)
  dtMax = oExcel.WorksheetFunction.Max(dtMax, oListObject.ListColumns("Actual Finish").DataBodyRange)
  dtMax = oExcel.WorksheetFunction.Max(dtMax, oListObject.ListColumns("Current Finish").DataBodyRange)
  dtMax = DateAdd("d", 6 - Weekday(dtMax), dtMax)
  
  Set oWorksheet = oWorkbook.Sheets.Add(After:=oWorkbook.Sheets(oWorkbook.Sheets.Count))
  oWorksheet.Name = "Chart_Data"
  oWorksheet.[A1:D1] = Array("WEEK", "BLF", "AF", "FF")
  lngLastRow = 2
  dtDate = dtMin & " 5:00 PM"
  oWorksheet.Cells(lngLastRow, 1) = dtDate
  Do While dtDate <= dtMax
    dtDate = DateAdd("d", 7, dtDate)
    lngLastRow = oWorksheet.[A1048576].End(xlUp).Row + 1
    oWorksheet.Cells(lngLastRow, 1) = dtDate
  Loop
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown)))
  oListObject.Name = "ChartData"
  oListObject.ListColumns("BLF").DataBodyRange.Formula2R1C1 = "=SUMPRODUCT((--Table1[BASELINE FINISH]<=[@WEEK])*--(Table1[BASELINE FINISH]>R[-1]C[-1])*1)"
  oListObject.ListColumns("AF").DataBodyRange.Formula2R1C1 = "=SUMPRODUCT((--Table1[ACTUAL FINISH]<=[@WEEK])*--(Table1[ACTUAL FINISH]>R[-1]C[-2])*1)"
  oListObject.ListColumns("FF").DataBodyRange.Formula2R1C1 = "=SUMPRODUCT((--Table1[CURRENT FINISH]<=[@WEEK])*--(Table1[CURRENT FINISH]>R[-1]C[-3])*--(Table1[ACTUAL FINISH]="""")*1)"
  oWorksheet.[I1] = dtStatus
  oWorksheet.[E1] = "BLF_CUM"
  oListObject.ListColumns("BLF_CUM").DataBodyRange.FormulaR1C1 = "=IF(ROW(R[-1]C)=1,[@BLF],R[-1]C+[@BLF])"
  oWorksheet.[F1] = "AF_CUM"
  oListObject.ListColumns("AF_CUM").DataBodyRange.FormulaR1C1 = "=IF(ROW(R[-1]C)=1,[@AF],IF([@WEEK]<=R1C9,R[-1]C+[@AF],""""))"
  oWorksheet.[G1] = "FF_CUM"
  oListObject.ListColumns("FF_CUM").DataBodyRange.FormulaR1C1 = "=IF([@WEEK]=R1C9,[@[AF_CUM]],IF([@WEEK]>R1C9,R[-1]C+[@FF],""""))"
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.SplitRow = 1
  oExcel.ActiveWindow.SplitColumn = 0
  oExcel.ActiveWindow.FreezePanes = True
  oListObject.Range.Columns.AutoFit
  oListObject.DataBodyRange.Copy
  oListObject.DataBodyRange.PasteSpecial xlPasteValuesAndNumberFormats
  lngLastRow = oWorksheet.Columns(1).Find(dtStatus).Row
  oWorksheet.Range(oWorksheet.Cells(2, 7), oWorksheet.Cells(lngLastRow - 1, 7)).ClearContents
  oWorksheet.Range(oWorksheet.Cells(lngLastRow + 1, 6), oWorksheet.Cells(1048576, 6)).ClearContents
  oWorksheet.[I1].Select
  oWorksheet.Shapes.AddChart2 227, xlLine
  Set oChart = oWorksheet.ChartObjects(oWorksheet.ChartObjects.Count)
  oChart.Chart.FullSeriesCollection(1).Delete
  oChart.Chart.SeriesCollection.NewSeries
  oChart.Chart.FullSeriesCollection(1).Name = "=Chart_Data!$E$1"
  oChart.Chart.FullSeriesCollection(1).Values = "=Chart_Data!" & oListObject.ListColumns("BLF_CUM").DataBodyRange.Address(True)
  oChart.Chart.FullSeriesCollection(1).XValues = "=Chart_Data!" & oListObject.ListColumns("WEEK").DataBodyRange.Address(True)
  oChart.Chart.SeriesCollection.NewSeries
  oChart.Chart.FullSeriesCollection(2).Name = "=Chart_Data!$F$1"
  oChart.Chart.FullSeriesCollection(2).Values = "=Chart_Data!" & oListObject.ListColumns("AF_CUM").DataBodyRange.Address(True)
  oChart.Chart.SeriesCollection.NewSeries
  oChart.Chart.FullSeriesCollection(3).Name = "=Chart_Data!$G$1"
  oChart.Chart.FullSeriesCollection(3).Values = "=Chart_Data!" & oListObject.ListColumns("FF_CUM").DataBodyRange.Address(True)
  oChart.Chart.SetElement (msoElementChartTitleAboveChart)
  oChart.Chart.SetElement (msoElementLegendBottom)
  oChart.Chart.ChartTitle.Text = strProject & " IMS - Task Completion" & Chr(10) & Format(dtStatus, "mm/dd/yyyy")
  oChart.Chart.ChartTitle.Characters(1, 25).Font.Bold = True
  oChart.Chart.Location Where:=xlLocationAsObject, Name:="SUMMARY"
  'must reset the object after move
  oWorksheet.Visible = xlSheetHidden
  Set oWorksheet = oWorkbook.Sheets("SUMMARY")
  Set oShape = oWorksheet.Shapes(oWorksheet.Shapes.Count)
  oShape.Top = oWorksheet.[J5].Top
  oShape.Left = oWorksheet.[J5].Left
  oShape.ScaleWidth 1.6663381968, msoFalse, msoScaleFromTopLeft
  oShape.ScaleHeight 1.8082112132, msoFalse, msoScaleFromTopLeft
  Set oChart = oWorksheet.ChartObjects(1)
  Set oSeries = oChart.Chart.SeriesCollection(1)
  With oSeries.Format.Line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
  End With
  Set oSeries = oChart.Chart.FullSeriesCollection(3)
  With oSeries.Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 112, 192)
    .Transparency = 0
    .DashStyle = msoLineDash
  End With
  oChart.Chart.Axes(xlCategory).CategoryType = xlTimeScale
  oChart.Chart.Axes(xlCategory).TickLabels.NumberFormat = "m/d/yyyy"

  Set oWorksheet = oWorkbook.Worksheets("Chart_Data")
  lngBaselineCount = oWorksheet.[E1048576].End(xlUp).Value
  lngForecastCount = oWorksheet.[G1048576].End(xlUp).Value
  If lngForecastCount < lngBaselineCount Then
    oWorkbook.Sheets("Summary").[J31] = "There are " & lngBaselineCount - lngForecastCount & " unstatused tasks in the current IMS."
    oWorkbook.Sheets("Summary").[J31].Font.Italic = True
    With oWorkbook.Sheets("Summary").[J31].Font
      .Color = -16777024
      .TintAndShade = 0
    End With
  End If
  
  oWorkbook.Sheets("Summary").Activate
  oWorkbook.Sheets("Summary").[A2].Select
  
  'save the file
  'todo: user-defined locations for metrics output
'  strDir = ActiveProject.Path & "\Metrics\"
'  strDir = strDir & Format(dtStatus, "yyyy-mm-dd") & "\"
'  If Dir(strDir, vbDirectory) = vbNullString Then MkDir strDir
'  strFile = strDir & Replace(strProject, " ", "_") & "_IMS_EarlyLateStartsFinishes_" & Format(ActiveProject.StatusDate, "yyyy-mm-dd") & ".xlsx"
'  If Dir(strFile) <> vbNullString Then Kill strFile
  oExcel.Calculation = xlCalculationAutomatic
  oExcel.ScreenUpdating = True
  'oWorkbook.SaveAs strFile, 51
  'oWorkbook.Close True
  
  'capture BEI
  Set oWorksheet = oWorkbook.Sheets("SUMMARY")
  cptCaptureMetric strProject, dtStatus, "BEI", Round(oWorksheet.Range("BEI[[#Totals],[BEI (Finishes)]]").Value, 2)
  Application.StatusBar = "Complete."
  DoEvents
  
'  If MsgBox("Complete. Open for review?", vbInformation + vbYesNo, "Late Starts and Finishes") = vbYes Then
'    oExcel.Workbooks.Open strFile
    oExcel.Visible = True
'    Application.ActivateMicrosoftApp pjMicrosoftExcel
'  End If

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oSeries = Nothing
  Set oChart = Nothing
  Set oShape = Nothing
  Set oOutlook = Nothing
  Set oMailItem = Nothing
  Set oDocument = Nothing
  Set oWord = Nothing
  Set oSelection = Nothing
  Set oEmailTemplate = Nothing
  Set oWorksheet = Nothing
  Set oCell = Nothing
  Set oRange = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing
  Set oTask = Nothing
  Set oShape = Nothing
  Set oChart = Nothing
  
  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics", "cptLateStartsFinishes", Err, Erl)
  Resume exit_here
End Sub

Sub cptCaptureMetric(strProgram As String, dtStatus As Date, strMetric As String, vMetric As Variant)
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
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set oRecordset = CreateObject("ADODB.Recordset")
  strFile = cptDir & "\settings\cpt-metrics.adtg"
  With oRecordset
    If Dir(strFile) = vbNullString Then
      'create it
      .Fields.Append "PROGRAM", adVarChar, 50
      .Fields.Append "STATUS_DATE", adDate
      .Fields.Append "SPI", adDouble
      .Fields.Append "SV", adDouble
      .Fields.Append "BEI", adDouble
      .Fields.Append "CPLI", adDouble
      .Fields.Append "CEI", adDouble
      .Fields.Append "TFCI", adDouble
      'others needed for ES?
      .Fields.Append "ES", adDate
      .Open
    Else
      .Open strFile
    End If
    .Filter = "PROGRAM='" & strProgram & "' AND STATUS_DATE=#" & dtStatus & "#"
    If Not .EOF Then
      .MoveFirst
      .Update Array(strMetric), Array(CDbl(vMetric))
    Else
      .AddNew Array("PROGRAM", "STATUS_DATE", strMetric), Array(strProgram, dtStatus, CDbl(vMetric))
    End If
    .Filter = ""
    .Save strFile, adPersistADTG
    .Close
  End With


exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptCaptureMetric", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetTrend_SPI()
  cptGetTrend "SPI"
End Sub

Sub cptGetTrend_BEI()
  cptGetTrend "BEI"
End Sub

Sub cptGetTrend_CPLI()
  cptGetTrend "CPLI"
End Sub

Sub cptGetTrend_CEI()
  cptGetTrend "CEI"
End Sub

Sub cptGetTrend(strMetric As String, Optional dtStatus As Date)
  'objects
  Dim oLegendEntry As Excel.LegendEntry
  Dim oChart As Excel.Chart
  Dim oChartObject As Excel.ChartObject
  Dim oListObject As Excel.ListObject
  Dim oRecordset As ADODB.Recordset
  Dim oWorksheet As Excel.Worksheet
  Dim oWorkbook As Excel.Workbook
  Dim oExcel As Excel.Application
  'strings
  Dim strHeader As String
  Dim strBanding As String
  Dim strProgram As String
  Dim strFile As String
  'longs
  Dim lngLastRow As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vBorder As Variant
  Dim vHeader As Variant
  Dim vBanding As Variant
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  strFile = cptDir & "\settings\cpt-metrics.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox strFile & " not found.", vbExclamation + vbOKOnly, "File Not Found"
    GoTo exit_here
  Else
    'get program
    strProgram = cptGetProgramAcronym
    If Len(strProgram) = 0 Then GoTo exit_here
    'get status date
    If dtStatus = 0 Then
      If Not IsDate(ActiveProject.StatusDate) Then
        MsgBox "This project requires a Status Date.", vbExclamation + vbOKOnly, "Invalid Status Date"
        Application.ChangeStatusDate
        If Not IsDate(ActiveProject.StatusDate) Then GoTo exit_here
      End If
      dtStatus = ActiveProject.StatusDate
    End If
    Set oRecordset = CreateObject("ADODB.Recordset")
    With oRecordset
      .Open strFile
      If .RecordCount = 0 Then
        MsgBox "No records found!", vbExclamation + vbOKOnly, "Trend Data: " & strMetric
        GoTo exit_here
      End If
      .Sort = "STATUS_DATE"
      .Filter = "PROGRAM='" & strProgram & "' AND STATUS_DATE<=#" & dtStatus & "#"
      If .EOF Then
        MsgBox "No records found for program '" & strProgram & "'!", vbExclamation + vbOKOnly, "Trend Data: " & strMetric
        GoTo exit_here
      End If
      .MoveFirst
      
      'get excel
      On Error Resume Next
      'Set oExcel = GetObject(, "Excel.Application")
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If oExcel Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
      End If
      'oExcel.WindowState = xlMinimized
      Set oWorkbook = oExcel.Workbooks.Add
      Set oWorksheet = oWorkbook.Sheets(1)
      oWorksheet.Name = strMetric & " TREND"
      oWorksheet.Cells(1, 1) = strProgram & " IMS - " & strMetric & IIf(strMetric = "SPI", "* ", " ") & "Trend"
      oWorksheet.Cells(1, 1).Font.Bold = True
      oWorksheet.Cells(1, 1).Font.Size = 16
      
      oWorksheet.Cells(2, 1) = dtStatus
      oWorksheet.Cells(2, 1).NumberFormat = "m/d/yyyy"
      oWorksheet.Cells(2, 1).HorizontalAlignment = xlCenter
      
      For Each vBorder In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        With oWorksheet.Cells(2, 1).Borders
          .LineStyle = xlContinuous
          .ThemeColor = 1
          .TintAndShade = -0.249946592608417
          .Weight = xlThin
        End With
      Next vBorder
      For Each vBorder In Array(xlDiagonalDown, xlDiagonalUp, xlInsideVertical, xlInsideHorizontal)
        oWorksheet.Cells(2, 1).Borders(vBorder).LineStyle = xlNone
      Next
      With oWorksheet.Cells(2, 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
        .PatternTintAndShade = 0
      End With
      
      If strMetric = "SPI" Then
        oWorksheet.Cells(3, 1) = "*SPI based in hours"
        oWorksheet.Cells(3, 1).Font.Italic = True
      End If
      
      oWorksheet.Cells(4, 1).Value = "STATUS DATE"
      oWorksheet.Cells(4, 2).Value = strMetric
      
      'handle headers
      Select Case strMetric
        Case "CEI"
          'todo: CEI Trend
        Case "TFCI"
          'todo: TFCI Trend
        Case Else 'SPI,BEI,CPLI
          strHeader = "CHANGE,CLEAR,< 0.95,0.95 - 0.99,1.00 - 1.05,> 1.05"
          oWorksheet.[C4:H4] = Split(strHeader, ",")
          vBanding = Array(0#, 0.94, 0.05, 0.06, 0.45)
          
      End Select
            
      'banding
      Do While Not .EOF
        lngLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, 1).End(xlUp).Row + 1
        oWorksheet.Cells(lngLastRow, 1) = FormatDateTime(CDate(.Fields("STATUS_DATE")), vbShortDate)
        oWorksheet.Cells(lngLastRow, 2) = .Fields(strMetric)
        Select Case strMetric
          Case "CEI"
            'todo: CEI Trend
          Case "TFCI"
            'todo: TFCI Trend
          Case Else 'SPI,BEI,CPLI
            oWorksheet.Range(oWorksheet.Cells(lngLastRow, 4), oWorksheet.Cells(lngLastRow, 8)).Style = "Comma"
            oWorksheet.Range(oWorksheet.Cells(lngLastRow, 4), oWorksheet.Cells(lngLastRow, 8)) = vBanding
            If lngLastRow = 5 Then
              oWorksheet.[C5] = 0
            Else
              oWorksheet.Cells(lngLastRow, 3).FormulaR1C1 = "=RC[-1]-R[-1]C[-1]"
            End If
        End Select
        .MoveNext
      Loop
      .Close
    End With
  End If
  
  'make it nice
  Set oListObject = oWorksheet.ListObjects.Add(xlSrcRange, oWorksheet.Range(oWorksheet.[A4].End(xlToRight), oWorksheet.[A4].End(xlDown)).Address, , xlYes)
  oListObject.ListColumns("CHANGE").DataBodyRange.Style = "Comma"
  oListObject.Range.Select
  oWorksheet.Shapes.AddChart2 332, xlLineMarkers, oWorksheet.[A4].End(xlToRight).Offset(0, 2).Left, oListObject.Range.Top
  Set oChartObject = oWorksheet.ChartObjects(1)
  Set oChart = oChartObject.Chart
  oChart.SetSourceData Source:=oWorksheet.Range(oListObject.Range.Address)
  oChart.Axes(xlCategory).CategoryType = xlTimeScale
  oChart.Axes(xlCategory).MajorUnit = 7
  oChart.Axes(xlCategory).MinorUnit = 7
  'oChart.Axes(xlValue).MinimumScale = 0
  'oChart.Axes(xlValue).MaximumScale = 2
  
  oChart.FullSeriesCollection("CHANGE").Delete

  oChart.FullSeriesCollection("CLEAR").ChartType = xlAreaStacked
  oChart.FullSeriesCollection("CLEAR").Format.Fill.Visible = msoFalse
  oChart.FullSeriesCollection("CLEAR").Format.Line.Visible = msoFalse
  
  'red
  oChart.FullSeriesCollection("< 0.95").ChartType = xlAreaStacked
  With oChart.FullSeriesCollection("< 0.95").Format.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 0, 0)
    .Transparency = 0
    .Solid
  End With
  With oChart.FullSeriesCollection("< 0.95").Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 0, 0)
    .Transparency = 0
  End With
  
  'yellow
  oChart.FullSeriesCollection("0.95 - 0.99").ChartType = xlAreaStacked
  With oChart.FullSeriesCollection("0.95 - 0.99").Format.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 255, 0)
    .Transparency = 0
    .Solid
  End With
  With oChart.FullSeriesCollection("0.95 - 0.99").Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(255, 255, 0)
    .Transparency = 0
  End With
  
  'green
  oChart.FullSeriesCollection("1.00 - 1.05").ChartType = xlAreaStacked
  With oChart.FullSeriesCollection("1.00 - 1.05").Format.Fill
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 176, 80)
    .Transparency = 0
    .Solid
  End With
  With oChart.FullSeriesCollection("1.00 - 1.05").Format.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(0, 176, 80)
    .Transparency = 0
  End With
  
  'blue
  oChart.FullSeriesCollection("> 1.05").ChartType = xlAreaStacked
  With oChart.FullSeriesCollection("> 1.05").Format.Fill
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorAccent5
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
    .Solid
  End With
  With oChart.FullSeriesCollection("> 1.05").Format.Line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorAccent5
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
  End With
  
  'strMetric
  With oChart.FullSeriesCollection(strMetric).Format.Line
    .Visible = msoTrue
    .ForeColor.ObjectThemeColor = msoThemeColorText1
    .ForeColor.TintAndShade = 0
    .ForeColor.Brightness = 0
    .Transparency = 0
  End With
  oChart.FullSeriesCollection(strMetric).Select
  oChart.SetElement (msoElementDataLabelBottom)
  oChart.FullSeriesCollection(strMetric).DataLabels.Select
  oExcel.Selection.Format.TextFrame2.TextRange.Font.Size = 11
  
  oChart.ChartTitle.Text = strProgram & " IMS - " & strMetric & " Trend" & vbLf & FormatDateTime(dtStatus, vbShortDate)
  oChart.SetElement (msoElementLegendRight)
  oChart.ChartArea.Format.TextFrame2.TextRange.Font.Size = 11
  oChart.ChartTitle.Characters(1, Len(strProgram & " IMS - " & strMetric & " Trend")).Font.Size = 14
  oChart.ChartTitle.Characters(1, Len(strProgram & " IMS - " & strMetric & " Trend")).Font.Bold = True
  
  oChartObject.Width = 792.173
  oChartObject.Height = 489.6
    
  For Each oLegendEntry In oChart.Legend.LegendEntries
    If oLegendEntry.LegendKey.Format.Fill.Visible = msoFalse Then
      oLegendEntry.Delete
      Exit For
    End If
  Next oLegendEntry
  
  oExcel.ActiveWindow.Zoom = 85
  oExcel.ActiveWindow.DisplayGridLines = False
  oWorksheet.Columns.AutoFit
  oWorksheet.Columns(1).ColumnWidth = 16
  oExcel.ActiveWindow.ScrollRow = 1
  oWorksheet.[A2].Select
  oExcel.Visible = True
  oExcel.WindowState = xlMaximized

exit_here:
  On Error Resume Next
  Set oLegendEntry = Nothing
  Set oChart = Nothing
  Set oChartObject = Nothing
  Set oListObject = Nothing
  Set oRecordset = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetTrend", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub cptCaptureAllMetrics()
  cptGET "SPI"
  cptGET "SV"
  cptGET "BEI"
  MsgBox "CPLI must be run manually.", vbInformation + vbOKOnly, "Capture All Metrics"
  cptGET "CEI"
  'cptGET "TFCI"
  cptGetEarnedSchedule
End Sub

Sub cptExportMetricsData()
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  'strings
  Dim strFile As String
  Dim strProgram As String
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  Application.StatusBar = "Exporting..."
  DoEvents
  
  strFile = cptDir & "\settings\cpt-metrics.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox strFile & " not found!", vbCritical + vbOKOnly, "File Not Found"
    GoTo exit_here
  End If
  
  strProgram = cptGetProgramAcronym
    
  Set oRecordset = CreateObject("ADODB.Recordset")
  oRecordset.Open strFile
  oRecordset.Filter = "PROGRAM='" & strProgram & "'"
  If oRecordset.RecordCount > 0 Then
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If oExcel Is Nothing Then
      Set oExcel = CreateObject("Excel.Application")
    End If
    'oExcel.Visible = True
    Set oWorkbook = oExcel.Workbooks.Add
    Set oWorksheet = oWorkbook.Sheets(1)
    oWorksheet.Name = strProgram
    For lngField = 0 To oRecordset.Fields.Count - 1
      oWorksheet.Cells(1, lngField + 1) = oRecordset.Fields(lngField).Name
    Next lngField
    oWorksheet.[A2].CopyFromRecordset oRecordset
    oWorksheet.Columns(2).NumberFormat = "[$-en-US]m/d/yy h:mm AM/PM;@"
    oExcel.ActiveWindow.Zoom = 85
    oWorksheet.Columns.AutoFit
    oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight)).Font.Bold = True
    Application.StatusBar = "Complete"
    DoEvents
    oExcel.Visible = True
    Application.ActivateMicrosoftApp pjMicrosoftExcel
    oExcel.WindowState = xlMaximized
  Else
    MsgBox "No records found for program '" & strProgram & "'", vbExclamation + vbOKOnly, "No Data"
  End If
  oRecordset.Filter = ""
  oRecordset.Close
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set oRecordset = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptExportMetricsData", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetEarnedSchedule()
  'objects
  Dim oAssignment As Assignment
  Dim oRecordset As ADODB.Recordset
  Dim oTasks As Tasks
  Dim oTask As Task
  Dim oTSV As TimeScaleValue
  Dim oTSVS As TimeScaleValues
  Dim oExcel As Excel.Application
  Dim oWorkbook As Excel.Workbook
  Dim oWorksheet As Excel.Worksheet
  Dim oListObject As Excel.ListObject
  Dim oRange As Excel.Range
  Dim oCell As Excel.Range
  'strings
  Dim strFormula As String
  Dim strLOEField As String
  Dim strEVP As String
  Dim strLOE As String
  Dim strProgram As String
  Dim strCon As String
  Dim strDir As String
  Dim strSQL As String
  Dim strFile As String
  'longs
  Dim lngRow As Long
  Dim lngAD As Long
  Dim lngDuration As Long
  Dim lngES As Long
  Dim lngLastRow As Long
  Dim lngEVP As Long
  Dim lngLOEField As Long
  Dim lngFile As Long
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngWork As Long
  'integers
  'doubles
  Dim dblBCWP As Double
  Dim dblBCWS As Double
  'booleans
  'variants
  Dim vBorder As Variant
  'dates
  Dim dtLatestFinish As Date
  Dim dtStart As Date
  Dim dtStatus As Date
  
  On Error Resume Next
  Set oTasks = ActiveProject.Tasks
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oTasks Is Nothing Then GoTo exit_here
  
  strProgram = cptGetProgramAcronym
  
  If Not IsDate(ActiveProject.StatusDate) Then
    MsgBox "Status Date required.", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  
  dtStatus = FormatDateTime(ActiveProject.StatusDate, vbShortDate) 'todo: format?
  
  If ActiveProject.ResourceCount = 0 Then
    MsgBox "Project must have resources.", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  
  If Not IsDate(ActiveProject.BaselineSavedDate(pjBaseline)) Then
    MsgBox "Project must be baselined.", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  
  If Not cptMetricsSettingsExist Then
    cptShowMetricsSettings_frm True
    If Not cptMetricsSettingsExist Then
      MsgBox "Metrics Settings required.", vbExclamation + vbOKOnly, "Earned Schedule"
      GoTo exit_here
    End If
  End If
  
  strFile = cptDir & "\Schema.ini"
  lngFile = FreeFile
  Open strFile For Output As #lngFile
  Print #lngFile, "[bcws.csv]"
  Print #lngFile, "Format=CSVDelimited"
  Print #lngFile, "ColNameHeaders=True"
  Print #lngFile, "Col1=WEEK_ENDING DateTime"
  Print #lngFile, "Col2=BCWS Double"
  Print #lngFile, "Col3=ETC Double"
  Close #lngFile
  
  strFile = cptDir & "\EarnedSchedule.csv"
  Open strFile For Output As #lngFile
  Print #lngFile, "WEEK_ENDING,BCWS,ETC,"
  
  strEVP = cptGetSetting("Metrics", "cboEVP")
  If Len(strEVP) = 0 Then
    MsgBox "Error obtaining setting Metrics.cboEVP", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  lngEVP = CLng(strEVP)
  strLOEField = cptGetSetting("Metrics", "cboLOEField")
  If Len(strLOEField) > 0 Then
    lngLOEField = CLng(strLOEField)
  Else
    MsgBox "Error obtaining setting Metrics.cboLOEField", vbExclamation + vbOKOnly, "Earned Schedule"
    GoTo exit_here
  End If
  strLOE = cptGetSetting("Metrics", "txtLOE")
  If Len(strLOE) = 0 Then
    MsgBox "Error obtaining setting Metrics.txtLOE", vbExclamation + vbOKOnly, "Earned Schedule"
  End If
  
  lngTasks = ActiveProject.Tasks.Count
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If oTask.GetField(lngLOEField) = strLOE Then GoTo next_task
    If oTask.Assignments.Count = 0 Then GoTo next_task
    For Each oAssignment In oTask.Assignments
      If oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment
      'todo: handle when EVP is between 0 and 1 vs. between 0 and 100
      'todo: if >0 <1 then don't divide it; if > 1 then divide it
      dblBCWP = dblBCWP + ((oAssignment.BaselineWork / 60) * (CLng(cptRegEx(oTask.GetField(lngEVP), "[0-9]{1,}")) / 100))
      Set oTSVS = oAssignment.TimeScaleData(oAssignment.BaselineStart, oAssignment.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks, 1)
      For Each oTSV In oTSVS
        Print #lngFile, DateAdd("d", -2, oTSV.EndDate) & "," & (Val(oTSV.Value) / 60) & ",0,"
      Next oTSV
      If oTask.RemainingDuration > 0 Then
        If oTask.Finish > dtLatestFinish Then dtLatestFinish = oTask.Finish
        Set oTSVS = oAssignment.TimeScaleData(oAssignment.Start, oAssignment.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks, 1)
        For Each oTSV In oTSVS
          If oTSV.StartDate > dtStatus Then
            Print #lngFile, DateAdd("d", -2, oTSV.EndDate) & ",0," & (Val(oTSV.Value) / 60) & ","
          End If
        Next oTSV
      End If
next_assignment:
    Next oAssignment
next_task:
    lngTask = lngTask + 1
    Application.StatusBar = "Analyzing...(" & Format(lngTask / lngTasks, "0%") & ")"
    DoEvents
  Next oTask

  Close #lngFile
  
  Application.StatusBar = "Building report..."
  DoEvents
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & cptDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  strSQL = "SELECT WEEK_ENDING,SUM(BCWS) AS BCWS,SUM(ETC) AS ETC "
  strSQL = strSQL & "FROM EarnedSchedule.csv "
  strSQL = strSQL & "GROUP BY WEEK_ENDING "
  strSQL = strSQL & "ORDER BY WEEK_ENDING"
  With oRecordset
    .Open strSQL, strCon, 1, 1
    If .RecordCount > 0 Then
      On Error Resume Next
      Set oExcel = GetObject(, "Excel.Application")
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If oExcel Is Nothing Then
        Set oExcel = CreateObject("Excel.Application")
        'oExcel.Visible = True
      End If
      Set oWorkbook = oExcel.Workbooks.Add
      Set oWorksheet = oWorkbook.Sheets(1)
      oWorksheet.Name = "Earned Schedule"
      oWorksheet.[A1:C1] = Array("WEEK_ENDING", "BCWS", "ETC")
      oWorksheet.[A2].CopyFromRecordset oRecordset
      oWorksheet.Columns(2).Style = "Comma"
      oWorksheet.Columns(3).Style = "Comma"
      oWorksheet.Columns.AutoFit 'todo: do this later
      oExcel.ActiveWindow.Zoom = 85
      oExcel.ActiveWindow.SplitRow = 1
      oExcel.ActiveWindow.SplitColumn = 0
      oExcel.ActiveWindow.FreezePanes = True
      'make cumulative column
      oWorksheet.Columns(3).Insert Shift:=xlRight
      oWorksheet.[C1].Value = "BCWS_CUM"
      oWorksheet.[C2].FormulaR1C1 = "=RC[-1]"
      lngLastRow = oWorksheet.[A1].End(xlDown).Row
      oWorksheet.Range(oWorksheet.Cells(3, 3), oWorksheet.Cells(lngLastRow, 3)).FormulaR1C1 = "=R[-1]C+RC[-1]"
      'find the status date/AD
      lngAD = CLng(oExcel.WorksheetFunction.Match(CLng(dtStatus), oWorksheet.[A:A], 1)) - 1
      oWorksheet.Cells(lngAD + 1, 1).Style = "Neutral"
      'find BCWP
      lngES = CLng(oExcel.WorksheetFunction.Match(CLng(dblBCWP), oWorksheet.[C:C], 1)) - 1
      oWorksheet.Cells(lngES + 1, 3).Style = "Neutral"
      oExcel.ActiveWindow.ScrollRow = lngES
      oWorksheet.Cells(lngES, 6) = "Status Date"
      oWorksheet.Cells(lngES, 7) = dtStatus
      oWorksheet.Names.Add "SD", oWorksheet.Cells(lngES, 7)
      oWorksheet.Cells(lngES, 7).NumberFormat = "m/d/yyyy"
      oWorksheet.Cells(lngES + 1, 6) = "BCWP"
      oWorksheet.Cells(lngES + 1, 7) = dblBCWP
      oWorksheet.Cells(lngES + 1, 7).Style = "Comma"
      oWorksheet.Cells(lngES + 1, 8) = "discrete only, in hours"
      'ES = duration planned to hit bcwp
      oWorksheet.Cells(lngES + 3, 6) = "Earned Schedule"
      oWorksheet.Cells(lngES + 3, 7) = lngES
      oWorksheet.Names.Add "ES", oWorksheet.Cells(lngES + 3, 7)
      oWorksheet.Cells(lngES + 3, 8) = "weeks"
      'AD = duration consumed to hit bcwp
      oWorksheet.Cells(lngES + 4, 6) = "Actual Duration"
      oWorksheet.Cells(lngES + 4, 7) = lngAD
      oWorksheet.Names.Add "AD", oWorksheet.Cells(lngES + 4, 7)
      oWorksheet.Cells(lngES + 4, 8) = "weeks"
      'SPI(t) = ES/ED
      oWorksheet.Cells(lngES + 5, 6) = "SPI(t)"
      oWorksheet.Cells(lngES + 5, 7) = "=ES/AD"
      oWorksheet.Names.Add "SPI_t", oWorksheet.Cells(lngES + 5, 7)
      oWorksheet.Cells(lngES + 5, 7).Style = "Comma"
      If oWorksheet.Cells(lngES + 5, 7) > 1.05 Then
        oWorksheet.Cells(lngES + 5, 7).Style = "Accent1"
        oWorksheet.Cells(lngES + 5, 8) = "too good?"
      ElseIf oWorksheet.Cells(lngES + 5, 7) >= 1 Then
        oWorksheet.Cells(lngES + 5, 7).Style = "Good"
        oWorksheet.Cells(lngES + 5, 8) = "on track"
      ElseIf oWorksheet.Cells(lngES + 5, 7) >= 0.95 Then
        oWorksheet.Cells(lngES + 5, 7).Style = "Neutral"
        oWorksheet.Cells(lngES + 5, 8) = "caution"
      ElseIf oWorksheet.Cells(lngES + 5, 7) < 0.95 Then
        oWorksheet.Cells(lngES + 5, 7).Style = "Bad"
        oWorksheet.Cells(lngES + 5, 8) = "warning"
      End If
      'PDWR = TD - ES
      oWorksheet.Cells(lngES + 7, 6) = "Planned Duration Work Remaining"
      'find max week ending where BCWS > 0
      strFormula = "=MATCH(MAXIFS("
      strFormula = strFormula & oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].End(xlDown)).AddressLocal(ReferenceStyle:=xlR1C1) & ","
      strFormula = strFormula & oWorksheet.Range(oWorksheet.[B2], oWorksheet.[B2].End(xlDown)).AddressLocal(ReferenceStyle:=xlR1C1) & ","
      strFormula = strFormula & """>0""),"
      strFormula = strFormula & oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A2].End(xlDown)).AddressLocal(ReferenceStyle:=xlR1C1) & ",0)"
      strFormula = strFormula & "-1-ES"
      oWorksheet.Cells(lngES + 7, 7).FormulaR1C1 = strFormula
      oWorksheet.Names.Add "PDWR", oWorksheet.Cells(lngES + 7, 7)
      oWorksheet.Cells(lngES + 7, 8) = "weeks"
      'RD = ETC DUR - AD
      oWorksheet.Cells(lngES + 8, 6) = "Remaining Duration"
      'find max week ending where ETC >0
      strFormula = "=MATCH(MAXIFS("
      strFormula = strFormula & oWorksheet.Range(oWorksheet.[A2], oWorksheet.[A2].End(xlDown)).AddressLocal(ReferenceStyle:=xlR1C1) & ","
      strFormula = strFormula & oWorksheet.Range(oWorksheet.[D2], oWorksheet.[D2].End(xlDown)).AddressLocal(ReferenceStyle:=xlR1C1) & ","
      strFormula = strFormula & """>0""),"
      strFormula = strFormula & oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A2].End(xlDown)).AddressLocal(ReferenceStyle:=xlR1C1) & ",0)"
      strFormula = strFormula & "-1-AD"
      oWorksheet.Cells(lngES + 8, 7) = strFormula
      oWorksheet.Names.Add "RD", oWorksheet.Cells(lngES + 8, 7)
      oWorksheet.Cells(lngES + 8, 8) = "weeks"
      'TSPI(ed) = PDWR / RD (ETC)
      oWorksheet.Cells(lngES + 9, 6) = "TSPI(ed)"
      oWorksheet.Cells(lngES + 9, 7) = "=PDWR/RD"
      oWorksheet.Names.Add "TSPI_ed", oWorksheet.Cells(lngES + 9, 7)
      oWorksheet.Cells(lngES + 9, 7).Style = "Comma"
      
      'compare SPI(t) vs. TSPI(ed)
      oWorksheet.Cells(lngES + 11, 6) = "|SPI(t)-TSPI(ed)|"
      oWorksheet.Cells(lngES + 11, 7).FormulaR1C1 = "=ABS(SPI_t-TSPI_ed)"
      oWorksheet.Cells(lngES + 11, 7).Style = "Comma"
      If oWorksheet.Cells(lngES + 11, 7) < 0.1 Then
        oWorksheet.Cells(lngES + 11, 7).Style = "Good"
      Else
        oWorksheet.Cells(lngES + 11, 7).Style = "Bad"
      End If
      oWorksheet.Cells(lngES + 12, 6) = "SPI(t)-TSPI(ed)"
      oWorksheet.Cells(lngES + 12, 7).FormulaR1C1 = "=SPI_t-TSPI_ed"
      oWorksheet.Cells(lngES + 12, 7).Style = "Comma"
      If oWorksheet.Cells(lngES + 12, 7) > 0.1 Then
        oWorksheet.Cells(lngES + 12, 8) = "overly pessimistic"
        oWorksheet.Cells(lngES + 12, 7).Style = "Bad"
      ElseIf oWorksheet.Cells(lngES + 12, 7) < -0.1 Then
        oWorksheet.Cells(lngES + 12, 8) = "overly optimistic"
        oWorksheet.Cells(lngES + 12, 7).Style = "Bad"
      Else
        oWorksheet.Cells(lngES + 12, 7).Style = "Good"
      End If
      'PDWR in days
      oWorksheet.Cells(lngES + 14, 6) = "PDWR"
      lngDuration = Application.DateDifference(dtStatus, dtLatestFinish)
      oWorksheet.Cells(lngES + 14, 7).FormulaR1C1 = lngDuration / (60 * 8)
      oWorksheet.Cells(lngES + 14, 7).Style = "Comma"
      oWorksheet.Cells(lngES + 14, 8) = "work days"
      
      'PDWR factored
      oWorksheet.Cells(lngES + 15, 6) = "PDWR/SPI(t)"
      lngDuration = lngDuration / oWorksheet.Range("SPI_t")
      oWorksheet.Cells(lngES + 15, 7).FormulaR1C1 = "=R[-1]C/SPI_t"
      oWorksheet.Cells(lngES + 15, 7).Style = "Comma"
      oWorksheet.Cells(lngES + 15, 8) = "work days"
      
      'IECD(es)
      oWorksheet.Cells(lngES + 16, 6) = "IECD(es)"
      oWorksheet.Cells(lngES + 16, 7).FormulaR1C1 = Application.DateAdd(dtStatus, lngDuration)
      oWorksheet.Cells(lngES + 16, 7).NumberFormat = "m/d/yyyy"
      oWorksheet.Cells(lngES + 16, 8) = "Using Calendar '" & ActiveProject.Calendar.Name & "'"
      
      'record the metric
      cptCaptureMetric strProgram, CDate(dtStatus & " 05:00 PM"), "ES", CDate(FormatDateTime(oWorksheet.Cells(lngES + 16, 7).Value, vbShortDate))
      'format the columns
      oWorksheet.Columns.AutoFit
      Set oRange = oWorksheet.Range(oWorksheet.[A1].End(xlToRight), oWorksheet.[A1].End(xlDown))
      For Each vBorder In Array(xlEdgeTop, xlEdgeLeft, xlEdgeRight, xlEdgeBottom)
        With oRange.Borders(vBorder)
          .LineStyle = xlContinuous
          .ColorIndex = xlAutomatic
          .TintAndShade = 0
          .Weight = xlThin
        End With
      Next vBorder
      For Each vBorder In Array(xlInsideVertical, xlInsideHorizontal)
        With oRange.Borders(vBorder)
          .LineStyle = xlContinuous
          .ThemeColor = 1
          .TintAndShade = -0.249946594869248
          .Weight = xlThin
        End With
      Next vBorder
      Set oRange = oWorksheet.Range(oWorksheet.[A1], oWorksheet.[A1].End(xlToRight))
      oRange.Font.Bold = True
      oRange.HorizontalAlignment = xlCenter
      With oRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998478032896
        .PatternTintAndShade = 0
      End With
    Else
      MsgBox "No records found.", vbExclamation + vbOKOnly, "Earned Schedule"
    End If
    .Close
  End With
  
  Application.StatusBar = "Complete."
  DoEvents
    
  oWorksheet.Cells(lngES + 16, 7).Select
  oExcel.Visible = True
  oExcel.WindowState = xlMaximized
  Application.ActivateMicrosoftApp pjMicrosoftExcel
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  DoEvents
  Kill cptDir & "\Schema.ini"
  Kill cptDir & "\EarnedSchedule.csv"
  Set oAssignment = Nothing
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Set oTasks = Nothing
  Set oTask = Nothing
  Set oTSV = Nothing
  Set oTSVS = Nothing
  Set oCell = Nothing
  Set oRange = Nothing
  Set oListObject = Nothing
  Set oWorksheet = Nothing
  Set oWorkbook = Nothing
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptGetEarnedSchedule", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowMetricsData_frm()
  'objects
  Dim oRecordset As ADODB.Recordset 'Object
  'strings
  Dim strPrograms As String
  Dim strProgram As String
  Dim strFile As String
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'ensure file exists
  strFile = cptDir & "\settings\cpt-metrics.adtg"
  If Dir(strFile) = vbNullString Then
    MsgBox strFile & " does not exist.", vbExclamation + vbOKOnly, "File Not Found"
    GoTo exit_here
  End If
  
  'ensure program name
  strProgram = cptGetProgramAcronym
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open strFile
    If .RecordCount = 0 Then
      MsgBox "No records found.", vbExclamation + vbOKOnly, "No Data"
      GoTo exit_here
    End If
    'gather unique programs
    .Sort = "PROGRAM"
    .MoveFirst
    Do While Not .EOF
      If InStr(strPrograms, .Fields("PROGRAM")) = 0 Then
        strPrograms = .Fields("PROGRAM") & ","
        cptMetricsData_frm.cboProgram.AddItem .Fields("PROGRAM")
      End If
      .MoveNext
    Loop
    cptMetricsData_frm.cboProgram.Value = strProgram
    cptMetricsData_frm.cboProgram.Locked = True 'todo: make cboProgram dynamic
    cptMetricsData_frm.cboProgram.Enabled = False 'todo: make cboProgram dynamic
    .MoveFirst
    .Sort = "STATUS_DATE DESC"
    .Filter = "PROGRAM='" & strProgram & "'"
    If Not .EOF Then
      'populate headers
      cptMetricsData_frm.lboHeader.AddItem
      For lngItem = 0 To .Fields.Count - 1
        cptMetricsData_frm.lboHeader.List(cptMetricsData_frm.lboHeader.ListCount - 1, lngItem) = .Fields(lngItem).Name
      Next lngItem
      'populate data
      .MoveFirst
      Do While Not .EOF
        cptMetricsData_frm.lboMetricsData.AddItem
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 0) = .Fields(0)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 1) = .Fields(1)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 2) = .Fields(2)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 3) = .Fields(3)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 4) = .Fields(4)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 5) = .Fields(5)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 6) = .Fields(6)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 7) = .Fields(7)
        cptMetricsData_frm.lboMetricsData.List(cptMetricsData_frm.lboMetricsData.ListCount - 1, 8) = IIf(CLng(.Fields(8)) = 0, "-", .Fields(8))
        .MoveNext
      Loop
      cptMetricsData_frm.lboMetricsData.Top = cptMetricsData_frm.lboHeader.Top + cptMetricsData_frm.lboHeader.Height
      cptMetricsData_frm.Show
    Else
      MsgBox "No records found for Program '" & strProgram & "'", vbExclamation + vbOKOnly, "No Records Found"
      GoTo exit_here
    End If
  End With
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptShowMetricsData_frm", Err, Erl)
  Resume exit_here
End Sub
