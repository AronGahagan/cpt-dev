Attribute VB_Name = "cptMetrics_bas"
'cpt-pre-release
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

'add disclaimer: unburdened hours - not meant to be precise - generally within +/- 1%

Sub cptExportMetricsExcel()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  MsgBox "Stay tuned...", vbInformation + vbOKOnly, "Under Construction..."

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMetrics_bas", "cptExportMetricsExcel", err, Erl)
  Resume exit_here
End Sub

Sub cptGetBAC()
  MsgBox cptGetMetric("bac")
End Sub

Sub cptGetETC()
  MsgBox cptGetMetric("etc")
End Sub

Sub cptGetBCWS()
  MsgBox cptGetMetric("bcws")
End Sub

Sub cptGetBCWP()
  MsgBox cptGetMetric("bcwp")
End Sub

Sub cptGetSPI()
  Call cptGET("SPI")
End Sub

Sub cptGetBEI()
  Call cptGET("BEI")
End Sub

Sub cptGET(strWhat As String)
'todo: need to store weekly bcwp, etc data somewhere
'objects
'strings
Dim strMsg As String
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

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Select Case strWhat
    Case "SPI"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      strMsg = "SPI = BCWP / BCWS" & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP, "#,##0") & " / " & Format(dblBCWS, "#,##0") & vbCrLf & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP / dblBCWS, "0%")
      MsgBox strMsg, vbInformation + vbOKOnly, "SPI (Hours)"
      
    Case "BEI"
      lngBEI_BF = CLng(cptGetMetric("bei_bf"))
      If lngBEI_BF = 0 Then
        MsgBox "No baseline finishes found.", vbExclamation + vbOKOnly, "No BEI"
        GoTo exit_here
      End If
      lngBEI_AF = CLng(cptGetMetric("bei_af"))
      strMsg = "BEI = # Actual Finishes / # Planned Finishes" & vbCrLf
      strMsg = strMsg & "BEI = " & Format(lngBEI_AF, "#,##0") & " / " & Format(lngBEI_BF, "#,##0") & vbCrLf & vbCrLf
      strMsg = strMsg & "BEI = " & Format((lngBEI_AF / lngBEI_BF), "#.#0")
      
    Case "cei"
      'todo: need to track previous week's plan
    
    Case "es" 'earned schedule
    
  End Select
  
  MsgBox strMsg, vbInformation + vbOKOnly, "cpt:Metrics"
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptMerics_Bas", "cptGet", err, Erl)
  Resume exit_here
End Sub

Function cptGetMetric(strGet As String) As Double
'todo: no screen changes!
'objects
Dim TSV As Object 'TimeScaleValue
Dim TSVS As Object 'TimeScaleValues
Dim Tasks As Object 'Tasks
Dim Task As Object 'Task
'strings
'longs
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
  
  If ActiveProject.StatusDate = "" Then
    MsgBox "This project has no status date. Please update and try again.", vbExclamation + vbOKOnly, "Metrics"
  Else
    dtStatus = ActiveProject.StatusDate
  End If
  
  cptSpeed True
  Call cptResetAll
  SelectAll
  Set Tasks = ActiveSelection.Tasks
  For Each Task In Tasks
    If Not Task Is Nothing Then
      If Task.ExternalTask Then GoTo next_task
      If Task.Summary Then GoTo next_task
      If Not Task.Active Then GoTo next_task
      If Task.BaselineWork > 0 Then
        Select Case strGet
          Case "bac"
            dblResult = dblResult + (Task.BaselineWork / 60)
            
          Case "etc"
            dblResult = dblResult + (Task.RemainingWork / 60)
            
          Case "bcws"
            If Task.Start < dtStatus Then
              Set TSVS = Task.TimeScaleData(Task.Start, dtStatus, pjTaskTimescaledBaselineWork, pjTimescaleWeeks)
              For Each TSV In TSVS
                dblResult = dblResult + IIf(TSV.Value = "", 0, TSV.Value) / 60
              Next
            End If
            
          Case "bcwp"
            dblResult = dblResult + ((Task.BaselineWork / 60) * (Task.PhysicalPercentComplete / 100))
            
          Case "bei_bf"
            dblResult = dblResult + IIf(Task.BaselineFinish <= dtStatus, 1, 0)
            
          Case "bei_af"
            dblResult = dblResult + IIf(Task.ActualFinish <= dtStatus, 1, 0)
            
        End Select
      End If 'bac>0
    End If 'not nothing
next_task:
  Next

  cptGetMetric = dblResult

exit_here:
  On Error Resume Next
  cptSpeed False
  Set TSV = Nothing
  Set TSVS = Nothing
  Set Tasks = Nothing
  Set Task = Nothing

  Exit Function
err_here:
  'Debug.Print Task.UniqueID & ": " & Task.Name
  Call cptHandleErr("cptMetrics_bas", "cptGetMetric", err, Erl)
  Resume exit_here

End Function

