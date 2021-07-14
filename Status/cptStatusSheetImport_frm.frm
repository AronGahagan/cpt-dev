VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptStatusSheetImport_frm 
   Caption         =   "Import Status Sheets"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810.001
   OleObjectBlob   =   "cptStatusSheetImport_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptStatusSheetImport_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboAF_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboAS_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboETC_Change()
  
  'ensure user didn't select the same field for New EVP and New ETC
  '-2147483630 = Black
  Me.lblEV.ForeColor = -2147483630
  Me.lblETC.ForeColor = -2147483630
  If Me.cboEV.Value = Me.cboETC.Value Then
    Me.lblEV.ForeColor = 192
    Me.lblETC.ForeColor = 192
  Else
    Call cptRefreshStatusImportTable
  End If
  
End Sub

Private Sub cboEV_Change()

  'ensure user didn't select the same field for New EVP and New ETC
  '-2147483630 = Black
  Me.lblEV.ForeColor = -2147483630
  Me.lblETC.ForeColor = -2147483630
  If Me.cboEV.Value = Me.cboETC.Value Then
    Me.lblEV.ForeColor = 192
    Me.lblETC.ForeColor = 192
  Else
    Call cptRefreshStatusImportTable
  End If
  
End Sub

Private Sub cboFF_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub cboFS_Change()
  Call cptRefreshStatusImportTable
End Sub

Private Sub chkAppend_Click()
    Me.cboAppendTo.Enabled = Me.chkAppend
  If Me.chkAppend Then
    Me.cboAppendTo.SetFocus
    Me.cboAppendTo.DropDown
  End If
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdImport_Click()
  
  'ensure user didn't select the same field for New EVP and New ETC
  '-2147483630 = Black
  Me.lblEV.ForeColor = -2147483630
  Me.lblETC.ForeColor = -2147483630
  If Me.cboEV.Value = Me.cboETC.Value Then
    Me.lblEV.ForeColor = 192
    Me.lblETC.ForeColor = 192
    MsgBox "Cannot import EVP and ETC to the same field.", vbExclamation + vbOKOnly, "Invalid Selections"
  Else
    'capture user settings
    cptSaveSetting "StatusSheetImport", "cboAS", Me.cboAS.Value
    cptSaveSetting "StatusSheetImport", "cboAF", Me.cboAF.Value
    cptSaveSetting "StatusSheetImport", "cboFS", Me.cboFS.Value
    cptSaveSetting "StatusSheetImport", "cboFF", Me.cboFF.Value
    cptSaveSetting "StatusSheetImport", "cboEVP", Me.cboEV.Value
    cptSaveSetting "StatusSheetImport", "cboETC", Me.cboETC.Value
    cptSaveSetting "StatusSheetImport", "cboContour", Me.cboContour.Value
    cptSaveSetting "StatusSheetImport", "chkNotes", CStr(Me.chkAppend)
    If Me.chkAppend Then
      cptSaveSetting "StatusSheetImport", "cboAppendTo", Me.cboAppendTo.Value
    Else
      cptSaveSetting "StatusSheetImport", "cboAppendTo", ""
    End If
    
    Call cptStatusSheetImport
  End If
  
End Sub

Private Sub cmdSelectFiles_Click()
'objects
Dim FileDialog As FileDialog
Dim xlApp As Excel.Application
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set xlApp = CreateObject("Excel.Application")
  Set FileDialog = xlApp.FileDialog(msoFileDialogFilePicker)
  With FileDialog
    .AllowMultiSelect = True
    .ButtonName = "Import"
    .InitialView = msoFileDialogViewDetails
    .InitialFileName = ActiveProject.Path & "\"
    .Title = "Select Returned Status Sheet(s):"
    .Filters.Add "Microsoft Excel Workbook (xlsx)", "*.xlsx"
    
    If .Show = -1 Then
      If .SelectedItems.Count > 0 Then
        For lngItem = 1 To .SelectedItems.Count
          cptStatusSheetImport_frm.TreeView1.Nodes.Add Text:=.SelectedItems(lngItem)
        Next lngItem
      End If
    End If
  End With

exit_here:
  On Error Resume Next
  Set FileDialog = Nothing
  Set xlApp = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheetImport_frm", "cmdSelectFiles_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptStatusSheet_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub

'Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, _
'                                  Effect As Long, _
'                                  Button As Integer, _
'                                  Shift As Integer, _
'                                  x As Single, _
'                                  y As Single)
'  Call cptAddFiles(Data)
'End Sub

Private Sub TreeView1_DblClick()
  'objects
  Dim oExcel As Object
  'strings
  Dim strPath As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If Me.TreeView1.Nodes.Count > 0 Then
    If Me.TreeView1.SelectedItem Is Nothing Then GoTo exit_here
    strPath = Me.TreeView1.SelectedItem.Text
    If Dir(strPath) <> vbNullString Then
      On Error Resume Next
      Set oExcel = GetObject(, "Excel.Application")
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      If oExcel Is Nothing Then Set oExcel = CreateObject("Excel.Application")
      oExcel.Workbooks.Open strPath
      oExcel.Visible = True
      Application.ActivateMicrosoftApp pjMicrosoftExcel
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set oExcel = Nothing

  Exit Sub
err_here:
  'Call HandleErr("cptStatusSheetImport_frm", "TreeView1.DblClick", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Private Sub UserForm_Initialize()
  'Me.TreeView1.OLEDropMode = ccOLEDropManual
End Sub
