VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDataDictionary_frm 
   Caption         =   "IMS Data Dictionary"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465.001
   OleObjectBlob   =   "cptDataDictionary_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDataDictionary_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.1.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboOpenWorkbooks_Change()
  If Not Me.cboOpenWorkbooks.Visible Then Exit Sub
  Select Case Me.cboOpenWorkbooks
    Case "Cancel"
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
    Case "Open another workbook..."
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
      Call cptImportDataDictionary
    Case "------------"
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
    Case Else
      Me.cmdImport.SetFocus
      Me.cboOpenWorkbooks.Visible = False
      Call cptImportDataDictionary(Me.cboOpenWorkbooks.Value)
  End Select

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdCustomFields_Click()
'long
Dim lngSelected As Long
'string
Dim strDescription As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not IsNull(Me.lboCustomFields.Value) Then
    lngSelected = Me.lboCustomFields.Value
    strDescription = Me.lboCustomFields.Column(2)
  End If
  Application.CustomizeField
  cptRefreshDictionary
  If lngSelected > 0 Then
    If Len(CustomFieldGetName(lngSelected)) > 0 Then
      Me.lboCustomFields.Value = lngSelected
      Me.lboCustomFields.Column(2) = strDescription
    End If
  End If

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "cmdCustomFields_Click()", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdExport_Click()
  Call cptExportDataDictionary
End Sub

Private Sub cmdFormGrow_Click()
  Me.cmdFormGrow.Visible = False
  Me.imgLogo.Visible = True
  Me.Height = 300
  Me.Width = 485.25
End Sub

Private Sub cmdFormShrink_Click()
  Me.cmdFormGrow.Visible = True
  Me.imgLogo.Visible = False
  Me.Height = 65
  Me.Width = 50
End Sub

Private Sub cmdImport_Click()
'objects
Dim xlApp As Object 'Excel.Application
Dim Workbook As Object 'Workbook
Dim Worksheet As Object 'Worksheet
'strings
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  On Error Resume Next
  Set xlApp = GetObject(, "Excel.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If Not xlApp Is Nothing Then
    Me.cboOpenWorkbooks.Clear
    For lngItem = 1 To xlApp.Workbooks.Count
      Me.cboOpenWorkbooks.AddItem xlApp.Workbooks(lngItem).Name
    Next
    Me.cboOpenWorkbooks.AddItem "------------"
    Me.cboOpenWorkbooks.AddItem "Open another workbook..."
    Me.cboOpenWorkbooks.AddItem "Cancel"
    
    Me.cboOpenWorkbooks.Visible = True
    Me.cboOpenWorkbooks.DropDown
  Else
    Call cptImportDataDictionary
  End If

exit_here:
  On Error Resume Next
  Set Worksheet = Nothing
  Set Workbook = Nothing
  Set xlApp = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "cmdImport_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub lblURL_Click()
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "lblURL_Click()", Err, Erl)
  Resume exit_here

End Sub

Private Sub lboCustomFields_AfterUpdate()
  If Not IsNull(Me.lboCustomFields.Value) Then Me.txtDescription = Me.lboCustomFields.Column(2)
End Sub

Private Sub txtDescription_AfterUpdate()
'objects
'strings
Dim strGUID As String
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If IsNull(Me.lboCustomFields.Value) Then GoTo exit_here

  Me.lblStatus.Caption = "Saving..."

  'get project uid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'find and update the record
  With CreateObject("ADODB.Recordset")
    .Open cptDir & "\settings\cpt-data-dictionary.adtg"
    .Filter = "PROJECT_ID='" & strGUID & "' AND FIELD_ID=" & CLng(Me.lboCustomFields.Value)
    If Not .EOF Then
      .Fields("DESCRIPTION") = Me.txtDescription.Text
      .Update
      Me.lboCustomFields.Column(2) = Me.txtDescription.Text
    End If
    .Filter = ""
    .Save
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Me.lblStatus.Caption = "Ready..."
  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "txtDescription_Change()", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtDescription_Change()
  Me.lblCharacterCount = 500 - Len(Me.txtDescription.Text)
End Sub

Private Sub txtDescription_Enter()
  Me.lblCharacterCount.Visible = True
End Sub

Private Sub txtDescription_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  Me.lblCharacterCount.Visible = False
End Sub

Private Sub txtFilter_Change()
'objects
Dim rst As ADODB.Recordset
'strings
Dim strDictionary As String, strFilter As String, strText As String, strGUID As String
'longs
Dim lngItem As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strDictionary = cptDir & "\settings\cpt-data-dictionary.adtg"

  If Dir(strDictionary) <> vbNullString Then
    Me.lboCustomFields.Clear
    Me.txtDescription = ""
    strGUID = ActiveProject.GetServerProjectGuid
    With CreateObject("ADODB.Recordset")
      .Open strDictionary
      .Sort = "CUSTOM_NAME"
      If Len(Me.txtFilter.Text) > 0 Then
        strText = cptRemoveIllegalCharacters(Me.txtFilter.Text)
        strFilter = "(CUSTOM_NAME LIKE '*" & strText & "*' AND PROJECT_ID='" & strGUID & "') "
        strFilter = strFilter & "OR "
        strFilter = strFilter & "(FIELD_NAME LIKE '*" & strText & "*' AND PROJECT_ID='" & strGUID & "') "
      Else
        strFilter = "PROJECT_ID='" & strGUID & "' "
      End If
      .Filter = strFilter
      If Not .EOF Then .MoveFirst
      lngItem = 0
      Do While Not .EOF
        Me.lboCustomFields.AddItem
        Me.lboCustomFields.Column(0, lngItem) = .Fields("FIELD_ID")
        Me.lboCustomFields.Column(1, lngItem) = .Fields("CUSTOM_NAME") & " (" & .Fields("FIELD_NAME") & ")"
        Me.lboCustomFields.Column(2, lngItem) = .Fields("DESCRIPTION")
        .MoveNext
        lngItem = lngItem + 1
      Loop
      .Close
    End With
  Else
    MsgBox "IMS Data Dictionary file not found!" & vbCrLf & "Please close and re-open the form to reset.", vbExclamation + vbOKOnly, "Error"
    GoTo exit_here
  End If

  Me.lblStatus.Caption = Me.lboCustomFields.ListCount & " result" & IIf(Me.lboCustomFields.ListCount = 1, "", "s")

exit_here:
  On Error Resume Next
  Set rst = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "txtFilter_Change()", Err, Erl)
  Resume exit_here
End Sub
