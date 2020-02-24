Attribute VB_Name = "cptQuickMetrics_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowQuickMetricsForm()
'objects
Dim aTaskFilterList As Object
Dim vFilter As Variant
'strings
'longs
Dim lngFilter As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set aTaskFilterList = CreateObject("System.Collections.SortedList")
  
  For Each vFilter In ActiveProject.TaskFilterList
    aTaskFilterList.Add vFilter, vFilter
  Next vFilter

  With cptQuickMetrics_frm
    .ComboBox1.Clear
    For lngFilter = 0 To aTaskFilterList.Count - 1
      .ComboBox1.AddItem aTaskFilterList.getKey(lngFilter)
    Next lngFilter
    .Show False
  End With

exit_here:
  On Error Resume Next
  Set aTaskFilterList = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptQuickMetrics_bas", "cptShowQuickMetricsForm", Err)
  MsgBox err.Number & ": " & err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub
