Attribute VB_Name = "cptAdjustment_bas"
'<cpt_version>v0.0.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
Public vResources As Variant
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowAdjustment_frm()
  'objects
  Dim oResource As MSProject.Resource
  'strings
  Dim strResources As String
  
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  For Each oResource In ActiveProject.Resources
    If oResource.Type = pjResourceTypeWork Then
      strResources = strResources & oResource.Name & ","
    End If
  Next
  strResources = Left(strResources, Len(strResources) - 1)
  vResources = Split(strResources, ",")
  
  Call cptQuickSort(vResources, 0, UBound(vResources))
  
  vResources = Split("All Resources," & Join(vResources, ","), ",")
  
  cptStartEvents
  
  With cptAdjustment_frm
    .lboHeader.Clear
    .lboHeader.AddItem
    .lboHeader.List(.lboHeader.ListCount - 1, 0) = "UID"
    .lboHeader.List(.lboHeader.ListCount - 1, 1) = "RESOURCE"
    .lboHeader.List(.lboHeader.ListCount - 1, 2) = "ETC"
    .lboHeader.List(.lboHeader.ListCount - 1, 3) = "NEW ETC"
    .cboResources.Clear
    .cboResources.List = vResources
    .cboResources.Value = "All Resources"
    .lblETC.Left = .lboAdjustmentPreview.Left + 25 + 150
    .lblETC.Top = .lboAdjustmentPreview.Top + .lboAdjustmentPreview.Height
    .lblETC.Caption = "-"
    .lblNewETC.Left = .lboAdjustmentPreview.Left + 25 + 150 + 54
    .lblNewETC.Top = .lboAdjustmentPreview.Top + .lboAdjustmentPreview.Height
    .lblNewETC.Caption = "-"
    .optTarget = True
    cptRefreshAdjustment
    .Show False
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptShowAdjustment_frm", Err)
  Resume exit_here

End Sub

Sub cptApplyAdjustment(strResource As String, strType As String, dblAmount As Double)
  'objects
  Dim oAssignment As MSProject.Assignment
  Dim oTask As MSProject.Task
  'strings
  'longs
  Dim lngType As Long
  'integers
  'doubles
  Dim dblTotal As Double
  'booleans
  Dim blnEffortDriven As Boolean
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.OpenUndoTransaction "cptAdjustment"

  'get total remaining work of all resource assignments
  For Each oTask In ActiveSelection.Tasks
    For Each oAssignment In oTask.Assignments
      'todo: allow application to all assignments on task
      If cptAdjustment_frm.cboResources.Value = "All Resources" Then
        dblTotal = dblTotal + (oAssignment.RemainingWork / 60) 'dblTotal starts in HOURS not MINUTES
      ElseIf oAssignment.ResourceName = cptAdjustment_frm.cboResources.Value Then
        dblTotal = dblTotal + (oAssignment.RemainingWork / 60) 'dblTotal starts in HOURS not MINUTES
      End If
    Next
  Next
  
'  Debug.Print "FROM: " & dblTotal
  
  'get new total remaining work
  If dblAdjustment <> 0 Then
    dblNewTotal = dblTotal + dblAdjustment 'in HOURS not MINUTES
'    Debug.Print "TO: " & dblNewTotal
  Else
'    Debug.Print "TO: " & dblNewTotal 'in HOURS not MINUTES
  End If

  'loop through and apply adjustments
  For Each oTask In ActiveSelection.Tasks
    'capture task settings
    lngType = oTask.Type
    oTask.Type = pjFixedDuration
    blnEffortDriven = oTask.EffortDriven
    oTask.EffortDriven = False
    For Each oAssignment In oTask.Assignments
      If oAssignment.ResourceName = cptAdjustment_frm.cboResources.Value Then
        'todo: allow application to all assignments on task
        'conversion from HOURS to MINUTES happens on next line
        oAssignment.RemainingWork = (((oAssignment.RemainingWork / 60) / (dblTotal / 60)) * (dblNewTotal / 60)) * 60
      End If
    Next oAssignment
    'restore task settings
    oTask.Type = lngType
    If oTask.Type <> pjFixedWork Then oTask.EffortDriven = blnEffortDriven
  Next oTask

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  Set oAssignment = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdjustment_bas", "cptApplyAdjustment", Err, Erl)
  Resume exit_here
End Sub

Sub cptRefreshAdjustment(Optional strResource As String, Optional dblAdjustment As Double, Optional dblNewTotal As Double)
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
  'booleans
  'variants
  'dates
  
  On Error Resume Next
  Set oTasks = ActiveSelection.Tasks
  If oTasks Is Nothing Then Exit Sub
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptAdjustment_frm
    If Not IsNull(.cboResources.Value) Then strResource = .cboResources.Value
    .lboAdjustmentPreview.Clear
    For Each oTask In ActiveSelection.Tasks
      If oTask Is Nothing Then GoTo next_task
      If oTask.ExternalTask Then GoTo next_task
      If Not oTask.Active Then GoTo next_task
      If IsDate(oTask.ActualFinish) Then GoTo next_task
      If oTask.Assignments.Count > 0 Then
        For Each oAssignment In oTask.Assignments
          If oAssignment.ResourceType <> pjResourceTypeWork Then GoTo next_assignment
          If .cboResources.Value <> "All Resources" And oAssignment.ResourceName <> strResource Then GoTo next_assignment
          .lboAdjustmentPreview.AddItem
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 0) = oTask.UniqueID
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 1) = oAssignment.ResourceName 'todo: necessary if filtered?
          .lboAdjustmentPreview.List(.lboAdjustmentPreview.ListCount - 1, 2) = Round((oAssignment.RemainingWork / 60), 2)
          dblETC = dblETC + oAssignment.RemainingWork
next_assignment:
        Next oAssignment
      End If
next_task:
    Next oTask

    If .lboAdjustmentPreview.ListCount > 0 Then
      If Len(.txtAmount) > 0 Then
        If .optDelta Then
          dblNewETC = (dblETC / 60) + CDbl(.txtAmount.Value)
          If dblNewETC < 0 Then
            dblNewETC = -(dblETC / 60) + 0.5
            .txtAmount.Value = Format(dblNewETC, "#,##0.00")
          End If
        ElseIf .optTarget Then
          dblNewETC = CDbl(.txtAmount.Value)
          If dblNewETC < 0 Then
            dblNewETC = 0.5
            .txtAmount.Value = Format(dblNewETC, "#,##0.00")
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
        If CDbl(.lboAdjustmentPreview.List(lngItem, 2)) = 0 Then
          .lboAdjustmentPreview.List(lngItem, 3) = 0
          GoTo next_item
        Else
          If dblNewETC <> 0 Then
            .lboAdjustmentPreview.List(lngItem, 3) = Format(((.lboAdjustmentPreview.List(lngItem, 2) * 60) / dblETC) * dblNewETC, "#,##0.00")
            dblNewTotal = dblNewTotal + ((.lboAdjustmentPreview.List(lngItem, 2) * 60) / dblETC) * dblNewETC
          ElseIf dblPercent <> 0 Then
            If dblPercent < 0 Then dblPercent = 1 + dblPercent
            .lboAdjustmentPreview.List(lngItem, 3) = Format((.lboAdjustmentPreview.List(lngItem, 2)) * dblPercent, "#,##0.00")
            dblNewTotal = dblNewTotal + (.lboAdjustmentPreview.List(lngItem, 2) * dblPercent)
          Else
            .lboAdjustmentPreview.List(lngItem, 3) = Format(.lboAdjustmentPreview.List(lngItem, 2), "#,##0.00")
            dblNewTotal = dblNewTotal + .lboAdjustmentPreview.List(lngItem, 2)
          End If
        End If
next_item:
      Next lngItem
      
    End If
    
    .lblNewETC.Caption = Format(dblNewTotal, "#,##0.00")
wrap_up:
    .lblETC.Caption = Format((dblETC / 60), "#,##0.00")
    .lblETC.Left = .lboAdjustmentPreview.Left + 25 + 150
    .lblETC.Top = .lboAdjustmentPreview.Top + .lboAdjustmentPreview.Height
    .lblNewETC.Left = .lboAdjustmentPreview.Left + 25 + 150 + 54
    .lblNewETC.Top = .lboAdjustmentPreview.Top + .lboAdjustmentPreview.Height

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
