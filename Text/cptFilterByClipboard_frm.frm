VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFilterByClipboard_frm 
   Caption         =   "Filter By Clipboard"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "cptFilterByClipboard_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFilterByClipboard_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>1.0.6</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub chkFilter_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
End Sub

Private Sub cmdClear_Click()
  Me.lboHeader.Clear
  Me.lboHeader.ColumnCount = 2
  Me.lboHeader.AddItem "UID"
  Me.lboHeader.Column(1, 0) = "Task Name"
  Me.lboFilter.ColumnCount = 2
  Me.txtFilter.Text = ""
  Me.txtFilter.Visible = True
  Me.lboFilter.Visible = False
  Call cptClearFreeField
  Call cptUpdateClipboard
  Me.txtFilter.SetFocus
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "lblURL", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboFilter_Click()
  'objects
  Dim oTasks As Tasks
  Dim oTask As Task
  'strings
  Dim strField As String
  'longs
  Dim lngUID As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  Dim dtGoTo As Date

  'round([Task's master project UID] / 4194304) = InsertedSubproject ID in Master
  'Task.UniqueId-(X*4194304)+X) where X is Subproject UID gets Task Index
  'task.uniqueid
  On Error Resume Next
  If Me.optUID Then
    lngUID = CLng(Me.lboFilter.Value)
    Set oTask = ActiveProject.Tasks.UniqueID(lngUID)
    strField = "Unique ID"
  ElseIf Me.optID Then
    lngUID = CLng(Me.lboFilter.Value)
    Set oTask = ActiveProject.Tasks.Item(lngUID)
    strField = "ID"
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not oTask Is Nothing Then
    
    If IsDate(oTask.Stop) Then
      dtGoTo = oTask.Stop
    Else
      dtGoTo = oTask.Start
    End If
    If ActiveWindow.ActivePane <> ActiveWindow.TopPane Then ActiveWindow.TopPane.Activate
  
    If ActiveProject.Subprojects.Count = 0 Then 'use EditGoto
      On Error Resume Next
      If Not EditGoTo(oTask.ID, dtGoTo) Then
        If MsgBox("Task " & strField & " " & lngUID & " is currently hidden. Would you like to remove all filters, show summary tasks, and show all tasks in order to find it?", vbQuestion + vbYesNo, "Reset View?") = vbYes Then
          ScreenUpdating = False
          FilterClear
          OptionsViewEx displaysummarytasks:=True
          SelectAll
          OutlineShowAllTasks
          ScreenUpdating = True
          If Not EditGoTo(oTask.ID, dtGoTo) Then
            MsgBox "An unknown error has occured--can't find it!", vbCritical + vbOKOnly, "Still can't find it"
          End If
        End If
      End If
    
    ElseIf ActiveProject.Subprojects.Count > 0 Then 'use Find
      On Error Resume Next
      If Not FindEx(strField, "equals", lngUID) Then
        If MsgBox("Task " & strField & " " & lngUID & " is currently hidden. Would you like to remove all filters, show summary tasks, and show all tasks in order to find it?", vbQuestion + vbYesNo, "Reset View?") = vbYes Then
          ScreenUpdating = False
          FilterClear
          OptionsViewEx displaysummarytasks:=True
          SelectAll
          OutlineShowAllTasks
          ScreenUpdating = True
          If Not FindEx(strField, "equals", lngUID) Then
            MsgBox "An unknown error has occured--can't find it!", vbCritical + vbOKOnly, "Still can't find it"
          End If
        End If
      End If
      
    End If 'ActiveProject.Subprojects.Count = 0
  Else
    MsgBox "Task " & strField & " " & lngUID & " not found in this project.", vbExclamation + vbOKOnly, strField & " not found"
  End If 'Not oTask Is Nothing
  
exit_here:
  On Error Resume Next
  ScreenUpdating = True
  Set oTasks = Nothing
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "lboFilter_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub optID_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
  Me.lboHeader.List(0, 0) = "ID"
  FilterClear
  Call cptUpdateClipboard
End Sub

Private Sub optUID_Click()
  Dim strFilter As String
  strFilter = Me.txtFilter.Text
  Me.txtFilter.Text = ""
  Me.txtFilter.Value = strFilter
  Me.lboHeader.List(0, 0) = "UID"
  FilterClear
  Call cptUpdateClipboard
End Sub

Private Sub tglEdit_Click()
  If Me.tglEdit Then
    Me.txtFilter.Visible = True
    Me.lboFilter.Visible = False
    Me.txtFilter.SetFocus
  Else
    If Len(Me.txtFilter.Value) = 0 Then
      Me.txtFilter.Visible = True
      Me.lboFilter.Visible = False
      Me.txtFilter.SetFocus
    Else
      Me.txtFilter.Visible = False
      Me.lboFilter.Visible = True
    End If
  End If
  Me.txtFilter.Height = Me.lboFilter.Height
End Sub

Private Sub txtFilter_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
  'objects
  'strings
  Dim strFilter As String
  Dim strItem As String
  'longs
  Dim lngDelimiter As Long
  Dim lngField As Long
  Dim lngRecord As Long
  Dim lngItem As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vRecord As Variant
  Dim strNewList As Variant
  Dim vData As Variant
  'dates

  'FilterClear

  'scrub the incoming data
  vData = Split(Data.GetText, vbCrLf)
  'guess the delimiter
  lngDelimiter = cptGuessDelimiter(vData, "^([^\t\,\;]*[\t\,\;])+")
  'populate lboFilter
  If UBound(vData) > 1 Then
    Me.lboFilter.Clear
    For lngRecord = 0 To UBound(vData)
      If vData(lngRecord) = "" Then GoTo next_record
        strItem = cptRegEx(CStr(vData(lngRecord)), "[0-9]*")
        If Len(strItem) > 0 Then
          'ignore UID 0
          If CLng(strItem) = 0 Then GoTo next_record
          'remove duplicates
          If cptRegEx(CStr(strNewList), "\b" & strItem & "\b") = "" Then
            strNewList = strNewList & CLng(strItem) & ","
          End If
        End If
next_record:
    Next lngRecord
    
  Else

    If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
    
    strFilter = Data.GetText
    vRecord = Split(strFilter, Chr(lngDelimiter))
        
    If IsEmpty(vRecord) Then GoTo exit_here
    
    For lngItem = 0 To UBound(vRecord)
      strItem = cptRegEx(CStr(vRecord(lngItem)), "[0-9]*")
      If Len(strItem) > 0 Then
        'ignore UID 0
        If CLng(strItem) = 0 Then GoTo next_item
        'remove duplicates
        If cptRegEx(CStr(strNewList), "\b" & strItem & "\b") = "" Then
          strNewList = strNewList & CLng(strItem) & ","
        End If
      End If
next_item:
    Next lngItem
    
  End If
  
  Cancel = True
  Me.txtFilter.Visible = True
  Me.txtFilter.Text = strNewList
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFilterByClipboard_frm", "txtFilter_BeforeDropOrPaste", Err, Erl)
  Resume exit_here

End Sub

Private Sub txtFilter_Change()
  If Me.txtFilter.Visible Then Call cptUpdateClipboard
End Sub

Private Sub txtFilter_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If Len(Me.txtFilter.Text) > 0 Then Call cptClipboardJump
End Sub

Private Sub txtFilter_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If Len(Me.txtFilter.Text) > 0 Then Call cptClipboardJump
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  If ActiveProject.Subprojects.Count > 0 Then
    Me.optID = False
    Me.optUID = True
    Me.optID.Locked = True
    Me.optID.ControlTipText = "Unavailable for Master/Subproject files"
    Me.optID.Enabled = False
  Else
    Me.optID.Enabled = True
    Me.optID.Locked = False
    Me.optID.ControlTipText = ""
  End If
End Sub

Private Sub UserForm_Terminate()
  Call cptClearFreeField
End Sub
