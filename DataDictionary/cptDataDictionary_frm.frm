VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDataDictionary_frm 
   Caption         =   "IMS Data Dictionary"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   OleObjectBlob   =   "cptDataDictionary_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDataDictionary_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

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
  Call cptHandleErr("cptDataDictionary_frm", "cmdCustomFields_Click()", err)
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

  'get project uid
  If Application.Version < 12 Then
    strGUID = ActiveProject.DatabaseProjectUniqueID
  Else
    strGUID = ActiveProject.GetServerProjectGuid
  End If
  
  'find and update the record
  With CreateObject("ADODB.Recordset")
    .Open cptDir & "\settings\data-dictionary.adtg"
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

  Exit Sub
err_here:
  Call cptHandleErr("cptDataDictionary_frm", "txtDescription_Change()", err)
  Resume exit_here
End Sub

Private Sub txtFilter_Change()
'need to capture native field name (or 'enterprise') and custom field name in adtg file
End Sub
