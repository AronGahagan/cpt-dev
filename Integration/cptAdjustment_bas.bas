Attribute VB_Name = "cptAdjustment_bas"
'<cpt_version>v0.0.3</cpt_version>
Option Explicit

Sub cptShowAdjustment_frm()
  'objects
  Dim oResources As Object 'Scripting.Dictionary
  Dim oResource As MSProject.Resource
  'strings
  Dim strIgnoreTaskType As String
  Dim strResources As String
  'variants
  Dim vResources As Variant
  Dim vResource As Variant
 
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If ActiveProject.ResourceCount = 0 Then
    MsgBox "This project has no resources.", vbExclamation + vbOKOnly, "No Resources"
    GoTo exit_here
  End If
  Set oResources = CreateObject("Scripting.Dictionary")
  For Each oResource In ActiveProject.Resources
    If oResource.Type = pjResourceTypeWork Then
      strResources = strResources & oResource.Name & ","
      oResources.Add oResource.Name, oResource.UniqueID
    End If
  Next
  If Len(strResources) = 0 Then
    MsgBox "No Labor [Resource Type=Work] resources found.", vbExclamation + vbOKOnly, "Adjustments"
    GoTo exit_here
  End If
  strResources = Left(strResources, Len(strResources) - 1)
  vResources = Split(strResources, ",")
  
  Call cptQuickSort(vResources, 0, UBound(vResources))
  
  'vResources = Split("All Resources," & Join(vResources, ","), ",")
  
  cptStartEvents
  
  With cptAdjustment_frm
    .Caption = "ETC Adjustment (" & cptGetVersion("cptAdjustment_frm") & ")"
    .lboHeader.Clear
    .lboHeader.AddItem
    .lboHeader.List(.lboHeader.ListCount - 1, 0) = "UID"
    .lboHeader.List(.lboHeader.ListCount - 1, 1) = "RESOURCE"
    .lboHeader.List(.lboHeader.ListCount - 1, 2) = "ETC"
    .lboHeader.List(.lboHeader.ListCount - 1, 3) = "NEW ETC"
    .lboTotal.Clear
    .lboTotal.AddItem
    .lboTotal.List(.lboTotal.ListCount - 1, 1) = "TOTAL"
    .cboResources.Clear
    .cboResources.AddItem
    .cboResources.List(0, 0) = "0"
    .cboResources.List(0, 1) = "All Resources"
    For Each vResource In vResources
      .cboResources.AddItem
      .cboResources.List(.cboResources.ListCount - 1, 0) = CStr(oResources(vResource))
      .cboResources.List(.cboResources.ListCount - 1, 1) = vResource
    Next vResource
    .cboResources.Value = "0" '"All Resources"
    .optTarget = True
    strIgnoreTaskType = cptGetSetting("ETCAdjustment", "chkIgnoreTaskType")
    If Len(strIgnoreTaskType) > 0 Then
      .chkIgnoreTaskType = CBool(strIgnoreTaskType)
    Else
      .chkIgnoreTaskType = True 'defaults to true
    End If
    cptRefreshAdjustment
    .Show False
  End With
  
exit_here:
  On Error Resume Next
  Set oResources = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptShowAdjustment_frm", Err)
  Resume exit_here

End Sub

Sub cptApplyAdjustment()
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  'strings
  Dim strStatusBar As String
  'longs
  Dim lngItem As Long
  Dim lngType As Long
  'integers
  'doubles
  'booleans
  Dim blnRefreshStatusBar As Boolean
  Dim blnIgnoreTaskType As Boolean
  Dim blnEffortDriven As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptSpeed True
  Application.OpenUndoTransaction "cptAdjustment"
  
  blnRefreshStatusBar = False
  If InStr(Application.StatusBar, "Remaining Work") > 0 Then
    blnRefreshStatusBar = True
    strStatusBar = Replace(Application.StatusBar, cptAdjustment_frm.lboTotal.List(0, 2), cptAdjustment_frm.lboTotal.List(0, 3))
  End If
  
  With cptAdjustment_frm
    blnIgnoreTaskType = .chkIgnoreTaskType
    cptSaveSetting "ETCAdjustment", "chkIgnoreTaskType", IIf(.chkIgnoreTaskType, "1", "0")
    If Not blnIgnoreTaskType Then
      If MsgBox("Adjustment may alter remaining duration on selected tasks." & vbCrLf & vbCrLf & "Continue?", vbExclamation + vbOKCancel, "Please Confirm") = vbCancel Then GoTo exit_here
    End If
    With .lboAdjustmentPreview
      If .ListCount = 0 Then GoTo exit_here
      For lngItem = 0 To .ListCount - 1
        Set oTask = ActiveProject.Tasks.UniqueID(.List(lngItem, 0))
        For Each oAssignment In oTask.Assignments
          If oAssignment.ResourceUniqueID = CLng(.List(lngItem, 1)) Then
            If blnIgnoreTaskType Then 'capture settings
              lngType = oTask.Type
              blnEffortDriven = oTask.EffortDriven
              'change type to apply
              oTask.Type = pjFixedDuration
              oTask.EffortDriven = False
            End If
            'apply adjustment
            oAssignment.RemainingWork = CDbl(.List(lngItem, 4)) * 60
            If blnIgnoreTaskType Then
              'restore settings
              oTask.Type = lngType
              If lngType <> pjFixedWork Then oTask.EffortDriven = blnEffortDriven
            End If
          End If
        Next oAssignment
      Next lngItem
    End With
  End With
  
  If blnRefreshStatusBar Then
    Application.StatusBar = strStatusBar
  End If
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptApplyAdjustment", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshAdjustment()
  'objects
  Dim oTasks As MSProject.Tasks
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngItem As Long
  'integers
  'doubles
  Dim dblPercent As Double
  Dim dblETC As Double
  Dim dblNewETC As Double
  Dim dblNewTotal As Double
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptAdjustment_frm
    .lboAdjustmentPreview.Clear
    For Each oTask In ActiveSelection.Tasks
      If oTask Is Nothing Then GoTo next_task
      If oTask.ExternalTask Then GoTo next_task
      If Not oTask.Active Then GoTo next_task
      If IsDate(oTask.ActualFinish) Then GoTo next_task
      If oTask.Assignments.Count > 0 Then
        For Each oAssignment In oTask.Assignments
          If Not .tglScope And oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment 'limit to pjWork
          If .cboResources.Value <> "0" And oAssignment.ResourceUniqueID <> CLng(.cboResources.Value) Then GoTo next_assignment 'limit to selected resource
          .lboAdjustmentPreview.AddItem
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 0) = oTask.UniqueID
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 1) = oAssignment.ResourceUniqueID
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 2) = oAssignment.ResourceName 'todo: necessary if filtered?
          If Not .tglScope Then 'hours
            .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 3) = Format(Round((oAssignment.RemainingWork / 60), 2), "#,###,##0.00")
            dblETC = dblETC + oAssignment.RemainingWork
          Else
            .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 3) = Format(Round((oAssignment.RemainingCost), 2), "currency")
            dblETC = dblETC + CDbl(oAssignment.RemainingCost)
          End If
next_assignment:
        Next oAssignment
      End If
next_task:
    Next oTask

    If .lboAdjustmentPreview.ListCount > 0 Then
      If Len(.txtAmount) > 0 Then
        If .optDelta Then
          If Not .tglScope Then 'hours
            dblNewETC = (dblETC / 60) + CDbl(.txtAmount.Value)
          Else 'hours
            dblNewETC = dblETC + CDbl(.txtAmount.Value)
          End If
          If dblNewETC < 0 Then
            If Not .tglScope Then 'hours
              dblNewETC = -(dblETC / 60) + 0.5
              .txtAmount.Value = Format(dblNewETC, "#,##0.00")
            Else
              dblNewETC = -dblETC + 0.5
              .txtAmount.Value = Format(dblNewETC, "currency")
            End If
          End If
        ElseIf .optTarget Then
          dblNewETC = CDbl(.txtAmount.Value)
          If dblNewETC < 0 Then
            dblNewETC = 0.5
            If Not .tglScope Then 'hours
              .txtAmount.Value = Format(dblNewETC, "#,##0.00")
            Else
              .txtAmount.Value = Format(dblNewETC, "currency")
            End If
          End If
        ElseIf .optPercent Then
          dblPercent = CDbl(.txtAmount.Value)
          If dblPercent = -1 Then
            dblPercent = 0.99
            .txtAmount.Value = "0.99"
          End If
        End If
      End If
      
      For lngItem = 0 To .lboAdjustmentPreview.ListCount - 1
        If CDbl(.lboAdjustmentPreview.List(lngItem, 3)) = 0 Then
          .lboAdjustmentPreview.List(lngItem, 4) = 0
          GoTo next_item
        Else
          If dblNewETC <> 0 Then
            If Not .tglScope Then 'hours
              .lboAdjustmentPreview.List(lngItem, 4) = Format(((.lboAdjustmentPreview.List(lngItem, 3) * 60) / dblETC) * dblNewETC, "#,##0.00")
              dblNewTotal = dblNewTotal + ((.lboAdjustmentPreview.List(lngItem, 3) * 60) / dblETC) * dblNewETC
            Else
              .lboAdjustmentPreview.List(lngItem, 4) = Format((.lboAdjustmentPreview.List(lngItem, 3) / dblETC) * dblNewETC, "currency")
              dblNewTotal = dblNewTotal + (.lboAdjustmentPreview.List(lngItem, 3) / dblETC) * dblNewETC
            End If
          ElseIf dblPercent <> 0 Then
            If dblPercent < 0 Then dblPercent = 1 + dblPercent
            If Not .tglScope Then 'hours
              .lboAdjustmentPreview.List(lngItem, 4) = Format((.lboAdjustmentPreview.List(lngItem, 3)) * dblPercent, "#,##0.00")
              dblNewTotal = dblNewTotal + (.lboAdjustmentPreview.List(lngItem, 3) * dblPercent)
            Else
              .lboAdjustmentPreview.List(lngItem, 4) = Format((.lboAdjustmentPreview.List(lngItem, 3)) * dblPercent, "currency")
              dblNewTotal = dblNewTotal + (.lboAdjustmentPreview.List(lngItem, 3) * dblPercent)
            End If
          Else
            .lboAdjustmentPreview.List(lngItem, 4) = .lboAdjustmentPreview.List(lngItem, 3) 'Format(.lboAdjustmentPreview.List(lngItem, 3), "#,##0.00")
            dblNewTotal = dblNewTotal + CDbl(.lboAdjustmentPreview.List(lngItem, 3))
          End If
        End If
next_item:
      Next lngItem
      
    End If
    
    'update totals
    If Not .tglScope Then 'hours
      .lboTotal.List(0, 2) = Format((dblETC / 60), "#,##0.00")
      .lboTotal.List(0, 3) = Format(dblNewTotal, "#,##0.00")
    Else
      .lboTotal.List(0, 2) = Format((dblETC), "currency")
      .lboTotal.List(0, 3) = Format(dblNewTotal, "currency")
    End If
    .lboTotal.Top = .lboHeader.Top + .lboHeader.Height + .lboAdjustmentPreview.Height
    
  End With

exit_here:
  On Error Resume Next
  Set oTasks = Nothing
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptRefreshAdjustment", Err, Erl)
  Resume exit_here
End Sub

Sub cptTargetToCost()
  'objects
  Dim oTSV As MSProject.TimeScaleValue
  Dim oTSVS As MSProject.TimeScaleValues
  Dim oTasks As MSProject.Tasks
  Dim oTask As MSProject.Task
  Dim oAssignment As MSProject.Assignment
  'strings
  Dim strMsg As String
  Dim strProrated As String
  'longs
  'integers
  'doubles
  Dim dblTotalCostPerUse As Double
  Dim dblTotalRemainingCost As Double
  Dim dblCostPerUse As Double
  Dim dblTargetCost As Double
  Dim dblRemainingCost As Double
  Dim dblRate As Double
  Dim dblCost As Double
  'booleans
  Dim blnProrated As Boolean
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If oTasks Is Nothing Then GoTo exit_here
    
  dblTargetCost = CDbl(cptAdjustment_frm.lboTotal.List(0, 3))
  
  Application.OpenUndoTransaction "Target to Cost"
  
  blnProrated = True
  
  For Each oTask In oTasks
    If IsDate(oTask.ActualFinish) Then
      strMsg = "Task UID " & oTask.UniqueID & " is marked 100% complete." & vbCrLf & vbCrLf
      strMsg = strMsg & "This feature is for incomplete tasks only."
      strMsg = strMsg & vbCrLf & vbCrLf & "Action canceled."
      MsgBox strMsg, vbCritical + vbOKOnly, "Task is Completed"
      GoTo exit_here
    End If
    For Each oAssignment In oTask.Assignments
      'only work with resources that are prorated
      If oAssignment.Resource.AccrueAt <> pjProrated Then
        strProrated = oAssignment.ResourceUniqueID & "|" & oAssignment.ResourceName & "|" & oAssignment.Resource.AccrueAt
        blnProrated = False
        Exit For
      End If
      'do not support overtime
      If oAssignment.OvertimeWork > 0 Or Val(oAssignment.OvertimeCost) <> 0 Then
        strMsg = "Resource '" & oAssignment.ResourceName & "' assigned to Task UID " & oAssignment.Task.UniqueID & " has Overtime Work/Costs." & vbCrLf & vbCrLf
        strMsg = strMsg & "This feature does not support Overtime Work/Cost."
        strMsg = strMsg & vbCrLf & vbCrLf & "Action canceled."
        MsgBox strMsg, vbCritical + vbOKOnly, "Overtime Not Supported"
        GoTo exit_here
      End If
      'If unstarted account for CostPerUse
      If Not IsDate(oTask.ActualStart) Or oAssignment.Resource.AccrueAt = pjEnd Then
        dblCostPerUse = cptGetPayRate(oAssignment, oAssignment.Start, "CostPerUse") 'oAssignment.Start not oTask.Start, in case of splits
      Else
        dblCostPerUse = 0
      End If
      dblTotalCostPerUse = dblTotalCostPerUse + dblCostPerUse
      dblTotalRemainingCost = dblTotalRemainingCost + (oAssignment.RemainingCost - dblCostPerUse)
    Next oAssignment
  Next oTask
  
  If Not blnProrated Then
    
    strMsg = "Resource UID " & Split(strProrated, "|")(0) & " - " & Chr(34) & Split(strProrated, "|")(1) & Chr(34)
    strMsg = strMsg & " is set to accrue costs at " & Choose(Split(strProrated, "|")(2), "Start", "End") & vbCrLf & vbCrLf
    strMsg = strMsg & "This feature is only compatible with the 'Prorated' accrual method." & vbCrLf & vbCrLf
    strMsg = strMsg & "Action canceled."
    MsgBox strMsg, vbCritical + vbOKOnly, "Incompatible Accrual Method"
    GoTo exit_here
  End If
  
  dblTargetCost = dblTargetCost - dblTotalCostPerUse
  
  'todo: if work resource has no cost, then ignore it
  'todo: rounding is causing problems
  
  For Each oTask In oTasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.ExternalTask Then GoTo next_task
    If Not oTask.Active Then GoTo next_task
    For Each oAssignment In oTask.Assignments
      If oAssignment.RemainingCost = 0 Then GoTo next_assignment
      If oAssignment.ResourceType <> pjResourceTypeCost Then
        Set oTSVS = oAssignment.TimeScaleData(oAssignment.Start, oAssignment.Finish, pjAssignmentTimescaledWork, pjTimescaleDays, 1)
        For Each oTSV In oTSVS
          'limit to remaining work by skipping actual work
          If oAssignment.TimeScaleData(oTSV.StartDate, oTSV.EndDate, pjAssignmentTimescaledActualWork, pjTimescaleDays)(1) = "" Then
            'limit to working days
            If oTSV.Value <> "" Then 'todo: what about if assignment.finish <> task.finish
              'get item rate
              dblRate = cptGetPayRate(oAssignment, oTSV.StartDate, "StandardRate")
              dblCost = oTSV.Value * dblRate
              oTSV.Value = Round(((dblCost / dblTotalRemainingCost) * dblTargetCost) / dblRate, 6)
              Application.CalculateProject
            End If
          End If
        Next oTSV
      ElseIf oAssignment.ResourceType = pjResourceTypeCost Then
        oAssignment.Cost = oAssignment.ActualCost + ((oAssignment.RemainingCost / dblTotalRemainingCost) * dblTargetCost)
      End If 'oAssignment.ResourceType
next_assignment:
    Next oAssignment
next_task:
  Next oTask
  
  'optionally refresh the status bar sums
  If cptGetShowStatusBarCountFirstRun Then
    If ActiveSelection.FieldIDList.Count = 1 Then
      Call cptGetSums(oTasks, ActiveSelection.FieldIDList(1))
    End If
  End If
  
exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Application.CalculateProject
  Application.Calculation = pjAutomatic
  Set oTSV = Nothing
  Set oTSVS = Nothing
  Set oTasks = Nothing
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptTargetToCost", Err, Erl)
  Resume exit_here
End Sub

Function cptGetPayRate(ByRef oAssignment As MSProject.Assignment, dtDate As Date, strAttribute As String) As Double
  'input: assignment so we have access to which CostRateTable is being used
  'input: date to pinpoint payrate
  'output: strAttribute StandardRate; OvertimeRate; CostPerUse
  Dim oCostRateTable As MSProject.CostRateTable
  Dim oPayRate As MSProject.PayRate
  Dim lngPayRates As Long
  Dim lngPayRate As Long
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oCostRateTable = oAssignment.Resource.CostRateTables(oAssignment.CostRateTable + 1)
  lngPayRates = oCostRateTable.PayRates.Count
  For lngPayRate = lngPayRates To 1 Step -1
    Set oPayRate = oCostRateTable.PayRates(lngPayRate)
    If dtDate >= oPayRate.EffectiveDate Then 'per use cost determined by assignment start
      If strAttribute = "CostPerUse" And oAssignment.ResourceType <> pjResourceTypeCost Then
        cptGetPayRate = oPayRate.CostPerUse
      ElseIf strAttribute = "StandardRate" And oAssignment.ResourceType <> pjResourceTypeCost Then
        cptGetPayRate = Replace(oPayRate.StandardRate, "/h", "")
      ElseIf strAttribute = "OvertimeRate" And oAssignment.ResourceType = pjResourceTypeWork Then
        cptGetPayRate = Replace(oPayRate.OvertimeRate, "/h", "")
      Else
        cptGetPayRate = 0
      End If
      Exit For
    End If
  Next lngPayRate
    
exit_here:
  On Error Resume Next
  Set oPayRate = Nothing
  Set oCostRateTable = Nothing
  
  Exit Function
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptGetPayRate", Err, Erl)
  Resume exit_here

End Function

