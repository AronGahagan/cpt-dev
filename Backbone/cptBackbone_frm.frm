VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptBackbone_frm 
   Caption         =   "Backbone"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11085
   OleObjectBlob   =   "cptBackbone_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptBackbone_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.2.2</cpt_version>
Option Explicit

Private Sub cboExport_Change()

  Me.chkIncludeThresholds.Enabled = False

  Select Case Me.cboExport
    Case "To Excel Workbook"
      'include header
      Me.chkIncludeHeaders = True
      Me.chkIncludeHeaders.Enabled = False
    Case "To CSV for MPM"
      'include header
      Me.chkIncludeHeaders = False
      Me.chkIncludeHeaders.Enabled = True
    Case "To CSV for COBRA"
      'include header
      Me.chkIncludeHeaders = True
      Me.chkIncludeHeaders.Enabled = False
      Me.chkIncludeThresholds.Enabled = True
    Case "To DI-MGMT-81334D Template"
      'hide include header
      Me.chkIncludeHeaders = True
      Me.chkIncludeHeaders.Enabled = False
      'get template
  End Select
  Me.cmdExport.SetFocus
  
End Sub

Private Sub cboImport_Change()

  Me.cmdExportTemplate.Visible = False
  Me.lblNote.Caption = ""
  Select Case Me.cboImport
    Case "From Excel Workbook"
      Me.cmdImport.Caption = "Import..."
      Me.cmdExportTemplate.Visible = True
      Me.chkAlsoCreateTasks.Visible = True
      Me.lblNote.Caption = "Import *.xlsx: Header CODE,LEVEL,DESCRIPTION in [A1:C1]"
    Case "From MSP Server Outline Code Export"
      Me.cmdImport.Caption = "Import..."
      Me.chkAlsoCreateTasks.Visible = False
      Me.lblNote.Caption = "Import *.xlsx: Header LEVEL,VALUE,DESCRIPTION in [A1:C1]"
    Case "From MIL-STD-881D Appendix B"
      Me.cmdImport.Caption = "Load"
      Me.chkAlsoCreateTasks.Visible = True
      Me.chkAlsoCreateTasks = True
      Me.chkAlsoCreateTasks.Enabled = False
      Me.lblNote.Caption = "Import generic CWBS as starting point."
    Case "From MIL-STD-881D Appendix E"
      Me.cmdImport.Caption = "Load"
      Me.chkAlsoCreateTasks.Visible = True
      Me.chkAlsoCreateTasks = True
      Me.chkAlsoCreateTasks.Enabled = False
      Me.lblNote.Caption = "Import generic CWBS as starting point."
    Case "From Existing Tasks"
      Me.cmdImport.Caption = "Create"
      Me.lblNote.Caption = "Replicate current task structure into " & Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 1) & "."
  End Select
  Me.txtNameIt.SetFocus
  
End Sub

Private Sub cboOutlineCodes_Change()
  Me.txtReplace.Text = ""
  Me.txtReplacement.Text = ""
  Me.lboOutlineCode.Clear
  If Not IsNull(Me.cboOutlineCodes.Value) Then
    If Len(CustomFieldGetName(Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0))) > 0 Then
      Call cptRefreshOutlineCodePreview(Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 1))
    End If
    Me.txtNameIt = CustomFieldGetName(Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0))
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdExport_Click()
'objects
'strings
'longs
Dim lngOutlineCode As Long
'integers
'doubles
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  lngOutlineCode = Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0)
  Select Case Me.cboExport
    Case "To Excel Workbook"
      Call cptExportOutlineCodeToExcel(lngOutlineCode)
    Case "To CSV for MPM"
      Call cptExportOutlineCodeForMPM(lngOutlineCode)
    Case "To CSV for COBRA"
      Call cptExportOutlineCodeForCOBRA(lngOutlineCode)
    Case "To DI-MGMT-81334D Template"
      Call cptExport81334D(lngOutlineCode)
  End Select

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "cmdExport_Click", Err, Erl)
  Resume exit_here
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  If Len(Me.txtNameIt.Value) = 0 Then
    MsgBox "Please provide a name.", vbExclamation + vbOKOnly, "No Name"
    Me.txtNameIt.SetFocus
    GoTo exit_here
  Else
    strOutlineCode = Me.txtNameIt
  End If
  lngOutlineCode = Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0)
  
  'ensure toppane is selected
  If Not ActiveWindow.BottomPane Is Nothing Then WindowActivate TopPane:=True
  'ensure a task view
  If ActiveWindow.TopPane.View.Type <> pjTaskItem Then
    ViewApply Application.DefaultView
  End If
  'if calendar is selected then change it
  If ActiveWindow.ActivePane.View.Screen = 13 Then
    ViewApply Application.DefaultView
  End If
  'create the new outline code
  CustomFieldRename lngOutlineCode, strOutlineCode
  Select Case Me.cboImport
    Case "From Excel Workbook"
      Call cptImportCWBSFromExcel(lngOutlineCode)
      
    Case "From MSP Server Outline Code Export"
      Call cptImportCWBSFromServer(lngOutlineCode)
    
    Case "From MIL-STD-881D Appendix B"
      Call cptImportAppendixB(lngOutlineCode)
      
    Case "From MIL-STD-881D Appendix E"
      Call cptImportAppendixE(lngOutlineCode)
      
    Case "From Existing Tasks"
      Call cptCreateCode(lngOutlineCode)
  
  End Select
  
exit_here:
  On Error Resume Next
  
  Exit Sub

err_here:
  Call cptHandleErr("cptOutlineCodes_bas", "cmdGo_Click", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub cmdReplace_Click()
Dim strOutlineCode As String

  If Len(Me.txtReplace) > 0 And Len(Me.txtReplacement) > 0 Then
    strOutlineCode = CustomFieldGetName(Me.cboOutlineCodes.Column(0))
    Call cptRenameInsideOutlineCode(strOutlineCode, Me.txtReplace, Me.txtReplacement)
  End If

End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "lblURL_Click", Err, Erl)
  Resume exit_here

End Sub

Private Sub optExport_Click()
  Call cptBackboneHideControls
  Me.cboExport.SetFocus
  Me.cboExport.DropDown
End Sub

Private Sub optImport_Click()
  Call cptBackboneHideControls
  Me.cboImport.SetFocus
  Me.cboImport.DropDown
End Sub

Private Sub optOutlineCode_Click()
  Call cptBackboneHideControls
  Me.cboOutlineCodes.SetFocus
  Me.cboOutlineCodes.DropDown
End Sub

Private Sub optReplace_Click()
  Call cptBackboneHideControls
  Me.txtReplace.SetFocus
End Sub

Private Sub txtReplace_Change()
Dim lngEntry As Long
Dim lngSelected As Long

  lngSelected = 0
  For lngEntry = 0 To Me.lboOutlineCode.ListCount - 1
    If Len(Me.txtReplace.Text) > 0 And InStr(Me.lboOutlineCode.List(lngEntry, 2), Me.txtReplace.Text) > 0 Then
      Me.lboOutlineCode.Selected(lngEntry) = True
      lngSelected = lngSelected + 1
      If lngSelected = 1 Then Me.lboOutlineCode.TopIndex = lngEntry
    Else
      Me.lboOutlineCode.Selected(lngEntry) = False
    End If
  Next lngEntry
  If lngSelected = 0 Then Me.lboOutlineCode.TopIndex = 0
  Me.lblFeedback.Caption = Format(lngSelected, "#,##0") & " found"
  
End Sub

Private Sub txtReplacement_Change()
'objects
Dim oOutlineCode As Object 'OutlineCode
Dim oLookupTable As Object 'LookupTable
'strings
Dim strOutlineCode As String
'long
Dim lngEntry As Long
'integers
'doubles
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  strOutlineCode = CustomFieldGetName(Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0))
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oLookupTable Is Nothing Then GoTo exit_here
  If Len(Me.txtReplace.Text) > 0 Then
    If Len(Me.txtReplacement.Text) > 0 Then
      For lngEntry = 0 To Me.lboOutlineCode.ListCount - 1
        If Me.lboOutlineCode.Selected(lngEntry) Then Me.lboOutlineCode.List(lngEntry, 2) = oLookupTable.Item(lngEntry + 1).FullName & " - " & Replace(oLookupTable.Item(lngEntry + 1).Description, Me.txtReplace, Me.txtReplacement.Text)
      Next lngEntry
    Else
      For lngEntry = 0 To Me.lboOutlineCode.ListCount - 1
        If Me.lboOutlineCode.Selected(lngEntry) Then Me.lboOutlineCode.List(lngEntry, 2) = oLookupTable.Item(lngEntry + 1).FullName & " - " & oLookupTable.Item(lngEntry + 1).Description
      Next lngEntry
    End If
  End If
  
exit_here:
  On Error Resume Next
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "txtReplacement_Change", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtNameIt_Change()
'longs
Dim lngField As Long

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'reset to default
  Me.txtNameIt.BorderColor = -2147483642
  Me.txtNameIt.ForeColor = -2147483640
  Me.lblStatus.Caption = "Ready..."
  
  'if name already exists then flag it
  lngField = 0
  On Error Resume Next
  lngField = FieldNameToFieldConstant(Me.txtNameIt.Text)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If lngField <> 0 Then 'exists
    If FieldNameToFieldConstant(Me.txtNameIt.Text) <> CLng(Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0)) Then
      Me.txtNameIt.BorderColor = 255
      Me.txtNameIt.ForeColor = 255
      Me.lblStatus.Caption = FieldConstantToFieldName(FieldNameToFieldConstant(Me.txtNameIt.Text)) & " is already named '" & Me.txtNameIt.Text & "'!"
    End If
  End If
  
exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "txtNameIt_Change", Err, Erl)
  Resume exit_here
  
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'objects
Dim oOutlineCode As Object 'OutlineCode
Dim oLookupTable As Object 'LookupTable
'strings
Dim strNewName As String
Dim strCustomName As String
Dim strOutlineCode As String
'longs
Dim lngItem As Long
Dim lngOutlineCode As Long
Dim lngSelected As Long
'integers
'doubles
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Exit Sub 'this slows the whole form down
  
  If Me.optReplace Then GoTo exit_here
  
  'have any outline codes been updated? update cbo options
  lngSelected = Me.cboOutlineCodes.ListIndex
  For lngItem = 0 To 9
    With Me.cboOutlineCodes
      lngOutlineCode = .List(lngItem, 0)
      strOutlineCode = .List(lngItem, 1)
      If InStr(strOutlineCode, "(") > 0 Then
        strOutlineCode = cptRegEx(strOutlineCode, "Outline Code[0-9]{1,2}")
        strCustomName = Replace(Replace(.List(lngItem, 1), strOutlineCode & " (", ""), ")", "")
      Else
        strCustomName = ""
      End If
      strNewName = CustomFieldGetName(FieldNameToFieldConstant(strOutlineCode))
      If strNewName <> strCustomName Then
        If Len(strNewName) > 0 Then
          .List(lngItem, 1) = strOutlineCode & " (" & strNewName & ")"
        Else
          .List(lngItem, 1) = strOutlineCode
        End If
        'the above triggers cboOutlineCodes_Change() so skip
        GoTo exit_here
      End If
    End With
  Next
  'has the currently selected outline code been edited?
  strOutlineCode = CustomFieldGetName(Me.cboOutlineCodes.List(Me.cboOutlineCodes.Value, 0))
  Set oOutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  On Error Resume Next
  Set oLookupTable = oOutlineCode.LookupTable
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If Not oLookupTable Is Nothing Then
    If Me.lboOutlineCode.ListCount = 0 Then
      Call cptRefreshOutlineCodePreview(strOutlineCode)
    Else
      For lngItem = 1 To oLookupTable.Count
        If Me.lboOutlineCode.List(lngItem - 1, 2) <> oLookupTable.Item(lngItem).FullName & " - " & oLookupTable.Item(lngItem).Description Then
          Me.lboOutlineCode.List(lngItem - 1, 2) = oLookupTable.Item(lngItem).FullName & " - " & oLookupTable.Item(lngItem).Description
        End If
      Next lngItem
    End If
  End If

exit_here:
  On Error Resume Next
  Set oOutlineCode = Nothing
  Set oLookupTable = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "UserForm_MouseMove", Err, Erl)
  Resume exit_here

End Sub
