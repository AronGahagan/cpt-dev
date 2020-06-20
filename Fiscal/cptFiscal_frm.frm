VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFiscal_frm 
   Caption         =   "Fiscal Calendar"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9840.001
   OleObjectBlob   =   "cptFiscal_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFiscal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>0.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdApply_Click()
'objects
Dim oException As MSProject.Exception
Dim oFiscalCal As MSProject.Calendar
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Are you sure?", vbExclamation + vbYesNo, "Confirm Overwrite") = vbNo Then GoTo exit_here

  Set oFiscalCal = ActiveProject.BaseCalendars(Me.cboFiscalCal.Value)

  'delete existing
  For Each oException In oFiscalCal.Exceptions
    oException.Delete
  Next oException

  For lngItem = 0 To Me.lboExceptions.ListCount - 1
    Set oException = oFiscalCal.Exceptions.Add(pjDaily, Me.lboExceptions.List(lngItem, 0), Me.lboExceptions.List(lngItem, 0), , Me.lboExceptions.List(lngItem, 1))
    oException.Shift1.Start = "8:00 AM"
    oException.Shift1.Finish = "12:00 PM"
    oException.Shift2.Start = "1:00 PM"
    oException.Shift2.Finish = "5:00 PM"
  Next lngItem

  Call cptUpdateFiscal

exit_here:
  On Error Resume Next
  Set oException = Nothing
  Set oFiscalCal = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_frm", "cmdApply_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub cmdDelete_Click()
'objects
Dim oFiscalCal As MSProject.Calendar
'strings
'longs
Dim lngSelected As Long
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  With Me.lboExceptions
    
    For lngItem = 0 To .ListCount - 1
      If .Selected(lngItem) Then lngSelected = lngSelected + 1
    Next lngItem
  
    If lngSelected = 0 Then GoTo exit_here
  
    If MsgBox("Are you sure you want to delete these?", vbExclamation + vbOKOnly, "Please Confirm") = vbNo Then GoTo exit_here
    
    Set oFiscalCal = ActiveProject.BaseCalendars(Me.cboFiscalCal.Value)
    
    For lngItem = .ListCount - 1 To 0 Step -1
      If .Selected(lngItem) Then
        .RemoveItem (lngItem)
        oFiscalCal.Exceptions(lngItem + 1).Delete
      End If
    Next lngItem
    
    Call cptUpdateFiscal
    
  End With

exit_here:
  On Error Resume Next
  Set oFiscalCal = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptFiscal_frm", "cmdDelete", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub cmdExport_Click()
'todo: select case cbo
Call cptExportCalendarExceptions
End Sub

Private Sub cmdImport_Click()
  'todo: select case cbo
  Call cptImportCalendarExceptions
End Sub

Private Sub cmdTemplate_Click()
  'todo: select case cbo
  Call cptExportExceptionsTemplate
End Sub

Private Sub tglEdit_Click()
Dim lngItem As Long, strExceptions As String
  Me.cmdApply.Enabled = Not Me.tglEdit
  With Me.lboExceptions
    For lngItem = 0 To .ListCount - 1
      strExceptions = strExceptions & .List(lngItem, 0) & vbTab
      strExceptions = strExceptions & .List(lngItem, 1) & vbCrLf
    Next lngItem
  End With
  Me.txtExceptions.Text = strExceptions
  Me.txtExceptions.Visible = Me.tglEdit
  Me.lboExceptions.Visible = Not Me.tglEdit
  
End Sub

Private Sub txtExceptions_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'objects
'strings
Dim strExceptions As String
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
Dim vException As Variant
Dim vExceptions As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'todo: guess delimiter
  vExceptions = Split(Data.GetText, vbCrLf)
  Me.lboExceptions.Clear
  For lngItem = 0 To UBound(vExceptions)
    vException = Split(vExceptions(lngItem), vbTab)
    If UBound(vException) > 0 Then 'labels included
      If IsDate(vException(0)) Then
        strExceptions = strExceptions & vException(0) & vbTab
        Me.lboExceptions.AddItem vException(0)
        strExceptions = strExceptions & vException(1) & vbCrLf
        Me.lboExceptions.List(Me.lboExceptions.ListCount - 1, 1) = vException(1)
      End If
    Else
      If IsDate(vExceptions(lngItem)) Then
        strExceptions = strExceptions & vExceptions(lngItem)
        Me.lboExceptions.AddItem vExceptions(lngItem)
        strExceptions = strExceptions & vbTab & Format(vExceptions(lngItem), "yyyy-mm") & vbCrLf
        Me.lboExceptions.List(Me.lboExceptions.ListCount - 1, 1) = Format(vExceptions(lngItem), "yyyy-mm")
      End If
    End If
  Next lngItem

  Cancel = True
  If Len(strExceptions) > 0 Then
    Me.txtExceptions.Text = strExceptions
    Me.lboExceptions.Visible = True
    Me.txtExceptions.Visible = False
    Me.tglEdit = False
  Else
    Me.txtExceptions.Visible = True
    Me.lboExceptions.Visible = False
    Me.tglEdit = True
  End If
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptFiscal_frm", "txtExceptions_BeforeDropOrPaste", Err, Erl)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub txtExceptions_Change()
  If Me.txtExceptions.Visible Then Call cptUpdateFiscal
End Sub
