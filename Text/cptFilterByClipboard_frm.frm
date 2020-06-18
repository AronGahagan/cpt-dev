VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFilterByClipboard_frm 
   Caption         =   "Filter By Clipboard"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   OleObjectBlob   =   "cptFilterByClipboard_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFilterByClipboard_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdClear_Click()
  Me.txtFilter = ""
End Sub

Private Sub lboFilter_Click()
'objects
Dim oTask As Task
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  On Error Resume Next
  If Me.optUID Then
    Set oTask = ActiveProject.Tasks.UniqueID(Me.lboFilter.Value)
  ElseIf Me.optID Then
    Set oTask = ActiveProject.Tasks.Item(CLng(Me.lboFilter.Value))
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTask Is Nothing Then
    EditGoTo oTask.ID
  End If
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptFilterByClipboard_frm", "lboFilter_Click", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub txtFilter_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
Dim vList As Variant

  'vList = Split(Data, vbTab)
  'todo: remove headers before pasting
  
End Sub

Private Sub txtFilter_Change()
'objects
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

  'todo: user can select tasks by navigating in the text box
  'todo: remove duplicates
  'todo: include csv,tsv,;,
  'todo: use regex to extract uid
  'todo: allow ID or UID
  
  Me.lboFilter.Clear
  strFilter = Me.txtFilter.Text
  If Len(strFilter) = 0 Then
    FilterClear
    GoTo exit_here
  End If
  
'  If Me.txtFilter.LineCount > 1 Then 'assume delimited list
'    If InStr(strFilter, ",") > 0 Then
'      vUID = Split(strFilter, ",")
'    ElseIf InStr(strFilter, ";") > 0 Then
'      vUID = Split(strFilter, ";")
'    End If
'    For lngItem = 0 To UBound(vUID) - 1
'      vUID(lngItem) = cptRegEx(vUID, "[0-9]*")
'    Next lngItem
'  Else 'assume multi-column
'    vUID = Split(strFilter, vbTab)
'  End If
  
  If InStr(strFilter, vbCrLf) > 0 Then  'vbtab
    vUID = Split(strFilter, vbCrLf)
    strFilter = ""
    For lngItem = 0 To UBound(vUID) - 1
      If cptRegEx(CStr(vUID(lngItem)), "[0-9]*") = "" Then GoTo next_line1
      If InStr(vUID(lngItem), vbTab) > 0 Then
        lngUID = Left(vUID(lngItem), InStr(vUID(lngItem), vbTab) - 1)
      Else
        lngUID = vUID(lngItem)
      End If
      strFilter = strFilter & lngUID & Chr$(9)
      Me.lboFilter.AddItem
      Me.lboFilter.List(Me.lboFilter.ListCount - 1, 0) = lngUID
      On Error Resume Next
      Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If Not oTask Is Nothing Then
        Me.lboFilter.List(Me.lboFilter.ListCount - 1, 1) = oTask.Name
        strFilter = strFilter & lngUID & Chr$(9)
        Set oTask = Nothing
      Else
        Me.lboFilter.List(Me.lboFilter.ListCount - 1, 1) = "< not found >"
      End If
next_line1:
    Next lngItem
  ElseIf InStr(strFilter, ",") > 0 Then 'csv
    vUID = Split(strFilter, ",")
    strFilter = ""
    For lngItem = 0 To UBound(vUID)
      If cptRegEx(CStr(vUID(lngItem)), "[0-9]*") = "" Then GoTo next_line2
      lngUID = vUID(lngItem)
      Me.lboFilter.AddItem
      Me.lboFilter.List(Me.lboFilter.ListCount - 1, 0) = lngUID
      On Error Resume Next
      Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If Not oTask Is Nothing Then
        Me.lboFilter.List(Me.lboFilter.ListCount - 1, 1) = oTask.Name
        strFilter = strFilter & lngUID & Chr$(9)
        Set oTask = Nothing
      Else
        Me.lboFilter.List(Me.lboFilter.ListCount - 1, 1) = "< not found >"
      End If
next_line2:
    Next lngItem
  Else
    strFilter = ""
  End If
  
  'todo: user can select highlight filter or autofilter
  If Len(strFilter) > 0 Then
    strFilter = Left(strFilter, Len(strFilter) - 1)
    If Me.optUID Then
      SetAutoFilter "Unique ID", FilterType:=pjAutoFilterIn, Criteria1:=strFilter
    ElseIf Me.optID Then
      SetAutoFilter "ID", FilterType:=pjAutoFilterIn, Criteria1:=strFilter
    End If
  End If
exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub txtFilter_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim vList As Variant
Dim lngUID As Long
  
  If Len(Me.txtFilter.Text) = 0 Then Exit Sub
  If Me.txtFilter.LineCount = 1 Then
    vList = Split(Me.txtFilter.Text, ",")
    If UBound(vList) > 0 Then
      If vList(UBound(vList)) = "" Then
        lngUID = vList(UBound(vList) - 1)
      Else
        lngUID = vList(Len(Left(Me.txtFilter.Text, IIf(Me.txtFilter.SelStart = 0, 1, Me.txtFilter.SelStart))) - Len(Replace(Left(Me.txtFilter.Text, IIf(Me.txtFilter.SelStart = 0, 1, Me.txtFilter.SelStart)), ",", "")))
      End If
    Else
      lngUID = vList(0)
    End If
    If Me.lboFilter.ListCount > 0 Then Me.lboFilter.Value = lngUID
  Else
    If Not IsNull(Me.txtFilter.CurLine) Then
      If Me.lboFilter.List(Me.txtFilter.CurLine, 1) <> "< not found >" Then
        If Me.optUID Then
          EditGoTo ActiveProject.Tasks.UniqueID(CLng(Me.lboFilter.List(Me.txtFilter.CurLine, 0))).ID
        ElseIf Me.optID Then
          EditGoTo ActiveProject.Tasks.Item(CLng(Me.lboFilter.List(Me.txtFilter.CurLine, 0))).ID
        End If
      End If
    End If
  End If

End Sub

Private Sub txtFilter_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim vList As Variant
Dim lngUID As Long

  If Len(Me.txtFilter.Text) = 0 Then Exit Sub
  If Me.txtFilter.LineCount = 1 Then
    vList = Split(Me.txtFilter.Text, ",")
    If UBound(vList) > 0 And vList(UBound(vList)) <> "" Then
      lngUID = vList(Len(Left(Me.txtFilter.Text, IIf(Me.txtFilter.SelStart = 0, 1, Me.txtFilter.SelStart))) - Len(Replace(Left(Me.txtFilter.Text, IIf(Me.txtFilter.SelStart = 0, 1, Me.txtFilter.SelStart)), ",", "")))
    Else
      lngUID = vList(0)
    End If
    If Me.lboFilter.ListCount > 0 Then Me.lboFilter.Value = lngUID
  Else
    If Not IsNull(Me.txtFilter.CurLine) Then
      If Me.lboFilter.List(Me.txtFilter.CurLine, 1) <> "< not found >" Then
        If Me.optUID Then
          EditGoTo ActiveProject.Tasks.UniqueID(CLng(Me.lboFilter.List(Me.txtFilter.CurLine, 0))).ID
        ElseIf Me.optID Then
          EditGoTo ActiveProject.Tasks.Item(CLng(Me.lboFilter.List(Me.txtFilter.CurLine, 0))).ID
        End If
      End If
    End If
  End If
End Sub
