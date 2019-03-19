VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDynamicFilter_frm 
   Caption         =   "Dynamic Filter"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   OleObjectBlob   =   "cptDynamicFilter_frm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cptDynamicFilter_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'<cpt_version>v0.1</cpt_version>

Private Sub cboField_Change()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub cboOperator_Change()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub chkHideSummaries_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub chkHighlight_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub chkKeepSelected_Click()
Dim Task As Task
  If Me.chkKeepSelected = True Then
    On Error Resume Next
    Set Task = ActiveSelection.Tasks(1)
    On Error GoTo 0
    If Task Is Nothing Then Me.chkKeepSelected = False
    Set Task = Nothing
  End If
End Sub

Private Sub chkShowRelatedSummaries_Click()
  If Me.Visible Then Me.txtFilter_Change
End Sub

Private Sub cmdClear_Click()
  Me.txtFilter.Text = ""
  If ActiveWindow.ActivePane <> ActiveWindow.TopPane Then ActiveWindow.TopPane.Activate
  On Error Resume Next
  OutlineShowAllTasks
  FilterApply "All Tasks"
  Me.txtFilter.SetFocus
End Sub

Private Sub cmdDone_Click()
  Me.Hide
End Sub

Sub txtFilter_Change()
'strings
Dim strField As String, strOperator As String, strFilterText As String
'booleans
Dim blnHideSummaryTasks As Boolean, blnHighlight As Boolean, blnKeepSelected As Boolean
'longs
Dim lgOriginalUID As Long
  
  If Me.ActiveControl.Name = "cmdClear" Then Exit Sub

  'assign values to variables
  On Error Resume Next
  lgOriginalUID = ActiveSelection.Tasks(1).UniqueID
  strField = Me.cboField
  strOperator = Me.cboOperator
  blnHideSummaryTasks = Not Me.chkHideSummaries
  blnShowRelatedSummaries = Me.chkShowRelatedSummaries
  blnHighlight = Me.chkHighlight
  blnKeepSelected = Me.chkKeepSelected
  If lgOriginalUID = 0 Then blnKeepSelected = False
  If blnHighlight Then
    strFilter = "Dynamic Highlight"
  Else
    strFilter = "Dynamic Filter"
  End If
  strFilterText = Me.txtFilter.Text
  ActiveWindow.TopPane.Activate
  
  'capture formatting that resembles a field name "[x]" and add a space "[x] "
  If Left(strFilterText, 1) = "[" And Right(strFilterText, 1) = "]" Then strFilterText = strFilterText & " "
  
  'capture wildcard - not allowed
  If InStr(strFilterText, "*") > 0 Or InStr(strFilterText, "%") > 0 Then
    MsgBox "Wildcards ('*') not allowed.", vbExclamation + vbOKOnly, "Error"
    strFilterText = Replace(strFilterText, "*", "")
    strFilterText = Replace(strFilterText, "%", "")
    Me.txtFilter = strFilterText
    Me.Show False
    Me.txtFilter.SetFocus
    Exit Sub
  End If
  
  Application.ScreenUpdating = False
  
  'build custom filter on the fly and apply it
  If Len(strFilterText) > 0 And Len(strOperator) > 0 Then
    If strField = "Task Name" Then strField = "Name"
    FilterEdit Name:=strFilter, taskfilter:=True, Create:=True, OverwriteExisting:=True, FieldName:=strField, test:=strOperator, Value:=strFilterText, operation:=IIf(blnKeepSelected Or blnHideSummaries, "Or", "None"), ShowInMenu:=False, showsummarytasks:=blnShowRelatedSummaries
  End If
  If blnKeepSelected Then
    FilterEdit Name:=strFilter, taskfilter:=True, NewFieldName:="Unique ID", test:="equals", Value:=lgOriginalUID, operation:="Or"
  End If
  If blnHideSummaryTasks Then
    FilterEdit Name:=strFilter, taskfilter:=True, NewFieldName:="Summary", test:="equals", Value:="No", operation:="And", parenthesis:=blnKeepSelected
  End If
  FilterEdit Name:=strFilter, showsummarytasks:=blnShowRelatedSummaries
  ActiveWindow.TopPane.Activate
  If Len(strFilterText) > 0 Then
    FilterApply strFilter, blnHighlight
  Else
    FilterClear
  End If
  On Error Resume Next
  'EditGoTo ActiveProject.Tasks.UniqueID(lgOriginalUID).ID
  Application.ScreenUpdating = True
  If lgOriginalUID > 0 And blnKeepSelected Then Application.Find "Unique ID", "equals", lgOriginalUID
  
End Sub
