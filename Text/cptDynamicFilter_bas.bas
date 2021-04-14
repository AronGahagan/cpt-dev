Attribute VB_Name = "cptDynamicFilter_bas"
'<cpt_version>v1.5.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private pCachedRegexes As Dictionary

Sub cptShowDynamicFilter_frm()
'objects
'strings
'longs
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  With cptDynamicFilter_frm
    .txtFilter = ""
    With .cboField
      .Clear
      .AddItem "Task Name"
      'todo: add all local custom text fields
    End With
    With .cboOperator
      .Clear
      .AddItem "equals"
      .AddItem "does not equal"
      .AddItem "contains"
      .AddItem "does not contain"
    End With
    .cboField = "Task Name"
    .chkKeepSelected = cptGetSetting("DynamicFilter", "KeepSelected") = "1"
    .chkHideSummaries = cptGetSetting("DynamicFilter", "IncludeSummaries") = "1"
    .chkShowRelatedSummaries = cptGetSetting("DynamicFilter", "RelatedSummaries") = "1"
    .chkHighlight = cptGetSetting("DynamicFilter", "Highlight") = "1"
    .tglRegEx = cptGetSetting("DynamicFilter", "geekMode") = "1"
    .chkHighlight.Visible = Not .tglRegEx
    .cboOperator.Value = cptGetSetting("DynamicFilter", "Operator")
    If .cboOperator.Value = "" Then
      If .tglRegEx Then .cboOperator.Value = "matches" Else .cboOperator = "contains"
    End If
    .Show False
    .txtFilter.SetFocus
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_bas", "cptShowDynamicFilter_frm", Err, Erl)
  Resume exit_here
End Sub

Sub cptGoRegEx(strRegEx As String)
  'objects
  Dim oTask As Task
  'strings
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True
  
  If Len(strRegEx) = 0 Then
    SetAutoFilter "Marked"
    GoTo exit_here
  End If
  
  lngUID = 0
  If cptDynamicFilter_frm.chkKeepSelected Then
    On Error Resume Next
    Set oTask = ActiveSelection.Tasks(1)
    lngUID = oTask.UniqueID
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  End If
  
  For Each oTask In ActiveProject.Tasks
    If oTask Is Nothing Then GoTo next_task
    If oTask.Marked Then oTask.Marked = False
    If cptDynamicFilter_frm.chkHideSummaries And oTask.Summary Then
      If Len(cptRxMatch(oTask.Name, strRegEx)) > 0 Then oTask.Marked = True
    ElseIf Not oTask.Summary Then
      If Len(cptRxMatch(oTask.Name, strRegEx)) > 0 Then oTask.Marked = True
    End If
next_task:
  Next oTask
  
  If lngUID > 0 Then ActiveProject.Tasks.UniqueID(lngUID).Marked = True
  
  FilterClear 'in case Dynamic Filter is applied
  OptionsViewEx displaysummarytasks:=True
  OutlineShowAllTasks
  OptionsViewEx displaysummarytasks:=cptDynamicFilter_frm.chkShowRelatedSummaries
  
  SetAutoFilter "Marked", pjAutoFilterFlagYes
  'todo: allow user-selected Flag or Marked
  'todo: if Dependency Browser is visible then do not allow use of Marked

exit_here:
  On Error Resume Next
  cptSpeed False
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDynamicFilter_bas", "cptGoRegEx", Err, Erl)
  Resume exit_here
End Sub

'===============================================================================
'attribution: https://bytecomb.com/increasing-performance-of-regular-expressions-in-vba/
'Private pCachedRegexes As Dictionary 'moved to top of module - AG
 
Public Function cptGetRegex( _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As RegExp
      
    ' Ensure the dictionary has been initialized
    If pCachedRegexes Is Nothing Then Set pCachedRegexes = CreateObject("Scripting.Dictionary")
    
    ' Build the unique key for the regex: a combination
    ' of the boolean properties and the pattern itself
    Dim rxKey As String
    rxKey = IIf(IgnoreCase, "1", "0") & _
            IIf(MultiLine, "1", "0") & _
            IIf(MatchGlobal, "1", "0") & _
            Pattern
            
    ' If the RegExp object doesn't already exist, create it
    If Not pCachedRegexes.Exists(rxKey) Then
        Dim oRegExp As New RegExp
        With oRegExp
            .Pattern = Pattern
            .IgnoreCase = IgnoreCase
            .MultiLine = MultiLine
            .Global = MatchGlobal
        End With
        Set pCachedRegexes(rxKey) = oRegExp
    End If
 
    ' Fetch and return the pre-compiled RegExp object
    Set cptGetRegex = pCachedRegexes(rxKey)
End Function

Public Function cptRxTest( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Boolean
 
    ' Wow, that was easy:
    cptRxTest = cptGetRegex(Pattern, IgnoreCase, MultiLine, False).test(SourceString)
    
End Function

Public Function cptRxMatch( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True) As Variant
    
    Dim oMatches As MatchCollection
    
    With cptGetRegex(Pattern, IgnoreCase, MultiLine, False)
        Set oMatches = .Execute(SourceString)
        If oMatches.Count > 0 Then
            cptRxMatch = oMatches(0).Value
        Else
            cptRxMatch = Null
        End If
    End With
 
End Function

Public Function cptRxMatches( _
    ByVal SourceString As String, _
    ByVal Pattern As String, _
    Optional ByVal IgnoreCase As Boolean = True, _
    Optional ByVal MultiLine As Boolean = True, _
    Optional ByVal MatchGlobal As Boolean = True) As Variant
 
    Dim oMatch As Match
    Dim arrMatches
    Dim lngCount As Long
    
    arrMatches = Array()
    With cptGetRegex(Pattern, IgnoreCase, MultiLine, MatchGlobal)
        For Each oMatch In .Execute(SourceString)
            ReDim Preserve arrMatches(lngCount)
            arrMatches(lngCount) = oMatch.Value
            lngCount = lngCount + 1
        Next
    End With
 
    cptRxMatches = arrMatches
End Function
