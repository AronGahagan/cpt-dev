Attribute VB_Name = "cptMetrics"
'cpt-pre-release
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

'add disclaimer: unburdened hours - not meant to be precise - generally within +/- 1%

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

Sub cptGET(strWhat As String)
'objects
'strings
Dim strMsg As String
'longs
'integers
'doubles
Dim dblBCWS As Double
Dim dblBCWP As Double
Dim dblResult As Double
'booleans
'variants
'dates

  'todo: need to store weekly bcwp, etc data somewhere
  'todo:

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Select Case strWhat
    Case "SPI"
      dblBCWP = cptGetMetric("bcwp")
      dblBCWS = cptGetMetric("bcws")
      strMsg = "SPI = BCWP / BCWS" & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP, "#,##0") & " / " & Format(dblBCWS, "#,##0") & vbCrLf
      strMsg = strMsg & "SPI = " & Format(dblBCWP / dblBCWS, "0%")
      MsgBox strMsg, vbInformation + vbOKOnly, "SPI (Hours)"
      
    Case "bei"
      
    
    Case "cei"
      'todo: need to track previous week's plan
    
    Case "es" 'earned schedule
    
  End Select
  
  MsgBox strMsg, vbOKOnly, "cpt:Metrics"
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptMerics_Bas", "cptGet", err)
  Resume exit_here
End Sub

Function cptGetMetric(strGet As String) As Double
'objects
Dim TSV As TimeScaleValue
Dim TSVS As TimeScaleValues
Dim Tasks As Tasks 'object
Dim Task As Task 'object
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
  Debug.Print Task.UniqueID & ": " & Task.Name
  Call HandleErr("cptMetrics_bas", "cptGetMetric", err)
  Resume exit_here

End Function

