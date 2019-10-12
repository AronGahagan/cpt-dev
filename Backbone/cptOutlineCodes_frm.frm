VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptOutlineCodes_frm 
   Caption         =   "Backbone"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11085
   OleObjectBlob   =   "cptOutlineCodes_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptOutlineCodes_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboExport_Change()

  Select Case Me.cboExport
    Case "To Excel Workbook"
      'include header
    Case "To CSV for MPM"
      'include header
    Case "To CSV for COBRA"
      'include header
    Case "To DI-MGMT-81334D"
      'hide include header
      'get template?
  End Select

End Sub

Private Sub cboImport_Change()

  Me.cmdExportTemplate.Visible = False
  Me.chkAlsoCreateTasks.Visible = False
  Me.lblNote.Caption = ""
  Select Case Me.cboImport
    Case "From Excel Workbook"
      Me.cmdImport.Caption = "Import..."
      Me.cmdExportTemplate.Visible = True
      Me.chkAlsoCreateTasks.Visible = True
      Me.lblNote.Caption = "Import .xlsx: Headers CODE,LEVEL,TITLE in [A1:C1]"
    Case "From MIL-STD-881D Appendix B"
      Me.cmdImport.Caption = "Load"
      Me.lblNote.Caption = "Import generic CWBS as starting point."
    Case "From Existing Tasks"
      Me.cmdImport.Caption = "Create"
      Me.lblNote.Caption = "Replicate current task structure into " & Me.cboOutlineCodes.Value & "."
  End Select
  
End Sub

Private Sub cboOutlineCodes_Change()
  Me.TreeView1.Nodes.Clear
  Me.txtReplace.Text = ""
  Me.txtReplacement.Text = ""
  If InStr(Me.cboOutlineCodes.Value, "(") > 0 Then
    Call cptRefreshOutlineCodePreview(CStr(Me.cboOutlineCodes.Value))
  End If
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdExportTemplate_Click()
  Call cptExportTemplate
End Sub

Private Sub cmdImport_Click()
'objects
'strings
Dim strOutlineCode As String
'longs
Dim lngOutlineCode As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'only fields with custom names have a left parenthesis
  If InStr(Me.cboOutlineCodes.Value, "(") > 0 Then
    strOutlineCode = Left(Me.cboOutlineCodes.Value, InStr(Me.cboOutlineCodes.Value, " (") - 1)
  Else
    strOutlineCode = Me.cboOutlineCodes.Value
  End If
  lngOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
  If Len(Me.txtNameIt.Value) = 0 Then
    MsgBox "Please provide a name.", vbExclamation + vbOKOnly, "No Name"
    GoTo exit_here
  Else
    strOutlineCode = Me.txtNameIt
  End If
  
  CustomFieldRename lngOutlineCode, strOutlineCode
  'ActiveProject.OutlineCodes.Add lngOutlineCode, strOutlineCode
  Select Case Me.cboImport
    Case "From Excel Workbook"
      Call cptImportCWBSFromExcel(lngOutlineCode)
      
    Case "From MIL-STD-881D Appendix B"
      Call cptImportAppendixB(lngOutlineCode)
      
    Case "From Existing Tasks"
      Call cptCreateCode(lngOutlineCode, strOutlineCode)
  
  End Select
  
exit_here:
  On Error Resume Next
  
  Exit Sub

err_here:
  Call cptHandleErr("cptOutlineCodes_bas", "cmdGo_Click", err, Erl)
  Resume exit_here
  
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "lblURL_Click", err, Erl)
  Resume exit_here

End Sub

Private Sub optImport_Click()
  Me.optExport.Value = Not Me.optImport.Value
End Sub

Private Sub txtReplace_Change()
Dim lngEntry As Long
  
  For lngEntry = 1 To Me.TreeView1.Nodes.Count
    Me.TreeView1.Nodes(lngEntry).Checked = False
    If Len(Me.txtReplace.Text) > 0 And InStr(Me.TreeView1.Nodes(lngEntry).Text, Me.txtReplace.Text) > 0 Then
      Me.TreeView1.Nodes(lngEntry).Checked = True
    End If
  Next lngEntry
  
End Sub

Private Sub txtReplace_Enter()
  Me.TreeView1.Checkboxes = True
End Sub

Private Sub txtReplacement_Change()
'objects
Dim OutlineCode As OutlineCode, LookupTable As LookupTable
'strings
Dim strOutlineCode As String
'long
Dim lngEntry As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strOutlineCode = Replace(Replace(Me.cboOutlineCodes.Value, cptRegEx(Me.cboOutlineCodes.Value, "Outline Code[1-10] \("), ""), ")", "")
  Set OutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set LookupTable = OutlineCode.LookupTable
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If LookupTable Is Nothing Then GoTo exit_here
  If Len(Me.txtReplace.Text) > 0 Then
    If Len(Me.txtReplacement.Text) > 0 Then
      For lngEntry = 1 To Me.TreeView1.Nodes.Count
        Me.TreeView1.Nodes(lngEntry).Text = Replace(LookupTable.Item(lngEntry).Description, Me.txtReplace.Text, Me.txtReplacement.Text)
      Next lngEntry
    Else
      For lngEntry = 1 To Me.TreeView1.Nodes.Count
        Me.TreeView1.Nodes(lngEntry).Text = LookupTable.Item(lngEntry).Description
      Next lngEntry
    End If
  End If
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptOutlineCodes_frm", "txtReplacement_Change", err, Erl)
  Resume exit_here
End Sub

Private Sub txtNameIt_Change()
'longs
Dim lngField As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'reset to default
  Me.txtNameIt.BorderColor = -2147483642
  Me.txtNameIt.ForeColor = -2147483640
  Me.lblStatus.Caption = "Ready..."
  
  'if name already exists then flag it
  lngField = 0
  On Error Resume Next
  lngField = FieldNameToFieldConstant(Me.txtNameIt.Text)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If lngField <> 0 Then 'exists
    Me.txtNameIt.BorderColor = 255
    Me.txtNameIt.ForeColor = 255
    Me.lblStatus.Caption = FieldConstantToFieldName(FieldNameToFieldConstant(Me.txtNameIt.Text)) & " is already named '" & Me.txtNameIt.Text & "!"
  End If
  
exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call cptHandleErr("cptOutlineCodes_frm", "txtNameIt_Change", err, Erl)
  Resume exit_here
  
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'objects
'strings
Dim strNewName As String
Dim strCustomName As String
Dim strOutlineCode As String
'longs
Dim lngItem As Long
Dim lngSelected As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'has anything changed?
  lngSelected = Me.cboOutlineCodes.ListIndex
  For lngItem = 0 To 9
    With cptOutlineCodes_frm.cboOutlineCodes
      strOutlineCode = .List(lngItem, 0)
      If InStr(strOutlineCode, "(") > 0 Then
        strOutlineCode = cptRegEx(strOutlineCode, "Outline Code[0-9]{1,2}")
        strCustomName = Replace(Replace(.List(lngItem, 0), strOutlineCode & " (", ""), ")", "")
      Else
        strCustomName = ""
      End If
      strNewName = CustomFieldGetName(FieldNameToFieldConstant(strOutlineCode))
      If strNewName <> strCustomName Then
        If Len(strNewName) > 0 Then
          .List(lngItem, 0) = strOutlineCode & " (" & strNewName & ")"
        Else
          .List(lngItem, 0) = strOutlineCode
        End If
        'cptRefreshOutlineCodePreview
      End If
    End With
  Next

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptUpgrades_frm", "UserForm_MouseMove", err, Erl)
  Resume exit_here

End Sub
