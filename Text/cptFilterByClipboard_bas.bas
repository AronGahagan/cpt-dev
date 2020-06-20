Attribute VB_Name = "cptFilterByClipboard_bas"
'<cpt_version>1.0.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptShowFilterByClipboardFrm()
'objects
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With cptFilterByClipboard_frm
    .tglEdit = False
    .lboHeader.Height = 12.5
    .lboHeader.Clear
    .lboHeader.AddItem
    .optUID = True
    .lboHeader.List(.lboHeader.ListCount - 1, 0) = "UID"
    .lboHeader.List(.lboHeader.ListCount - 1, 1) = "Task Name"
    .lboHeader.Width = .lboFilter.Width
    .lboFilter.Top = .lboHeader.Top + .lboHeader.Height
    .txtFilter.Top = .lboFilter.Top
    .txtFilter.Width = .lboFilter.Width
    .txtFilter.Height = .lboFilter.Height
    .txtFilter.Visible = True
    .lboFilter.Visible = False
    .chkFilter = True
    .Show False
  End With
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptShowFilterByClipboardFrm", Err, Erl)
  Resume exit_here
  
End Sub

Sub cptClipboardJump()
  'objects
  'strings
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vList As Variant
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(cptFilterByClipboard_frm.txtFilter.Text) = 0 Then Exit Sub
  vList = Split(cptFilterByClipboard_frm.txtFilter.Text, ",")
  If UBound(vList) > 0 Then
    If cptFilterByClipboard_frm.txtFilter.SelStart = Len(cptFilterByClipboard_frm.txtFilter.Text) Then Exit Sub
    lngUID = vList(Len(Left(cptFilterByClipboard_frm.txtFilter.Text, IIf(cptFilterByClipboard_frm.txtFilter.SelStart = 0, 1, cptFilterByClipboard_frm.txtFilter.SelStart))) - Len(Replace(Left(cptFilterByClipboard_frm.txtFilter.Text, IIf(cptFilterByClipboard_frm.txtFilter.SelStart = 0, 1, cptFilterByClipboard_frm.txtFilter.SelStart)), ",", "")))
  Else
    lngUID = vList(0)
  End If
  If cptFilterByClipboard_frm.lboFilter.ListCount > 0 Then cptFilterByClipboard_frm.lboFilter.Value = lngUID

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptCliipboardJump", Err, Erl)
  Resume exit_here
End Sub

Sub cptUpdateClipboard()
Dim oTask As Task
'strings
Dim strFilter As String
'longs
Dim lngItem As Long
Dim lngUID As Long
'integers
'doubles
'booleans
'variants
Dim vUID As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'cptFilterByClipboard_frm.lboFilter.Clear
  strFilter = cptFilterByClipboard_frm.txtFilter.Text
  If Len(strFilter) = 0 Then
    FilterClear
    GoTo exit_here
  End If
  
  vUID = Split(strFilter, ",")
  strFilter = ""
  If IsEmpty(vUID) Then GoTo exit_here
  For lngItem = 0 To UBound(vUID)
    If vUID(lngItem) = "" Then GoTo next_item
    lngUID = vUID(lngItem)
    strFilter = strFilter & lngUID & Chr$(9)
    If cptFilterByClipboard_frm.lboFilter.ListCount = lngItem Then
      cptFilterByClipboard_frm.lboFilter.AddItem
      cptFilterByClipboard_frm.lboFilter.List(cptFilterByClipboard_frm.lboFilter.ListCount - 1, 0) = lngUID
    End If
    On Error Resume Next
    If cptFilterByClipboard_frm.optUID Then
      Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
    Else
      Set oTask = ActiveProject.Tasks(lngUID)
    End If
    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    If Not oTask Is Nothing Then
      If cptFilterByClipboard_frm.lboFilter.ListCount = lngItem + 1 Then
        cptFilterByClipboard_frm.lboFilter.List(cptFilterByClipboard_frm.lboFilter.ListCount - 1, 1) = oTask.Name
      End If
      strFilter = strFilter & lngUID & Chr$(9)
      'oTask.Number20 = lngItem
      Set oTask = Nothing
    Else
      cptFilterByClipboard_frm.lboFilter.List(cptFilterByClipboard_frm.lboFilter.ListCount - 1, 1) = "< not found >"
    End If
next_item:
  Next lngItem
  
  If Not cptFilterByClipboard_frm.tglEdit Then
    cptFilterByClipboard_frm.lboFilter.Visible = True
    cptFilterByClipboard_frm.txtFilter.Visible = False
  End If
  
  If Len(strFilter) > 0 And cptFilterByClipboard_frm.chkFilter Then
    ScreenUpdating = False
    OptionsViewEx displaysummarytasks:=True
    SelectAll
    OutlineShowAllTasks
    SelectBeginning
    strFilter = Left(strFilter, Len(strFilter) - 1)
    If cptFilterByClipboard_frm.optUID Then
      SetAutoFilter "Unique ID", FilterType:=pjAutoFilterIn, Criteria1:=strFilter
    ElseIf cptFilterByClipboard_frm.optID Then
      SetAutoFilter "ID", FilterType:=pjAutoFilterIn, Criteria1:=strFilter
    End If
    OptionsViewEx projectsummary:=False, displayoutlinenumber:=False, displaynameindent:=False, displaysummarytasks:=False
    'Sort "Number20"
  End If
  
exit_here:
  On Error Resume Next
  ScreenUpdating = True
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_bas", "cptUpdateClipboard", Err, Erl)
  Resume exit_here
End Sub
