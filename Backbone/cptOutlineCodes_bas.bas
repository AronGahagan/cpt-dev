Attribute VB_Name = "cptOutlineCodes_bas"
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub ShowCptOutlineCode_frms()
'longs
Dim lngCode As Long, lngOutlineCode As Long
'strings
Dim strOutlineCode As String, strOutlineCodeName As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  cptOutlineCode_frm.cboOutlineCodes.Clear

  For lngCode = 1 To 10
    strOutlineCode = "Outline Code" & lngCode
    lngOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
    strOutlineCodeName = Application.CustomFieldGetName(lngOutlineCode)
    cptOutlineCode_frm.cboOutlineCodes.AddItem
    If Len(strOutlineCodeName) > 0 Then
      strOutlineCode = strOutlineCode & " (" & strOutlineCodeName & ")"
      'cptOutlineCode_frm.cboOutlineCodes.Column(1, lngCode - 1) = "(" & strOutlineCodeName & ")"
    End If
    cptOutlineCode_frm.cboOutlineCodes.Column(0, lngCode - 1) = strOutlineCode
  Next lngCode
  
  cptOutlineCode_frm.txtNameIt = ""
  cptOutlineCode_frm.cmdCancel.Caption = "Cancel"
  cptOutlineCode_frm.cboOutlineCodes.Value = cptOutlineCode_frm.cboOutlineCodes.List(0)
  cptOutlineCode_frm.Show False
  cptOutlineCode_frm.cboOutlineCodes.SetFocus

exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call cptHandleErr("cptOutlineCodes_bas", "ShowcptOutlineCode_frms", err, Erl)
  Resume exit_here
End Sub

Sub cptCreateCode(lngOutlineCode As Long, strOutlineCodeName As String)
'objects
Dim objOutlineCode As OutlineCode, objLookupTable As LookupTable, objLookupTableEntry As LookupTableEntry
Dim Task As Task, xlApp As Excel.Application
'strings
Dim strWBS As String, strParent As String, strChild As String
'longs
Dim lngUID As Long, lngTasks As Long, lngTask As Long, lngLevel As Long
'variants
Dim aOutlineCode As Variant, tmr As Date

  tmr = Now
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'ensure name doesn't already exist - trust form formatting
  If cptOutlineCode_frm.txtNameIt.BorderColor = 255 Then GoTo exit_here

  'first name the field and create the code mask
  Application.CustomFieldRename lngOutlineCode, strOutlineCodeName
  For lngLevel = 1 To 10
    CustomOutlineCodeEditEx lngOutlineCode, Level:=lngLevel, Sequence:=pjCustomOutlineCodeCharacters, Length:="Any", Separator:="."
  Next lngLevel
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLookUpTableCodes:=False, OnlyLeaves:=False, LookupDefault:=False, SortOrder:=0
  Set objOutlineCode = ActiveProject.OutlineCodes("CWBS")
  Set objLookupTable = objOutlineCode.LookupTable
  
  lngTasks = ActiveProject.Tasks.Count
  
  For Each Task In ActiveProject.Tasks
    lngTask = lngTask + 1
    If Task.OutlineLevel = 1 Then
      Set objLookupTableEntry = objLookupTable.AddChild(Task.WBS)
      objLookupTableEntry.Description = Task.Name
    End If
    Task.SetField lngOutlineCode, Task.WBS
    objLookupTable.Item(lngTask).Description = Task.Name
    cptOutlineCode_frm.lblProgress.Width = ((lngTask - 1) / lngTasks) * cptOutlineCode_frm.lblStatus.Width
    cptOutlineCode_frm.lblStatus.Caption = Format(lngTask - 1, "#,##0") & " / " & Format(lngTasks, "#,##0") & " (" & Format((lngTask - 1) / lngTasks, "0%") & ") [" & Format(Now - tmr, "hh:nn:ss") & "]"
  Next Task
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLeaves:=True, OnlyLookUpTableCodes:=True
  cptOutlineCode_frm.lblStatus.Caption = "Complete."
  Application.StatusBar = "Complete."
  cptOutlineCode_frm.cmdCancel.Caption = "Done"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  SpeedOFF
  Set objOutlineCode = Nothing
  Set objLookupTable = Nothing
  Set objLookupTableEntry = Nothing
  Set Task = Nothing
  xlApp.Quit
  Set xlApp = Nothing
  Exit Sub
err_here:
  MsgBox err.Number & ": " & err.Description, vbExclamation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Sub RenameInsideOutlineCode(strOutlineCode As String, strFind As String, strReplace As String)
'usage: Call RenameOutlineCode("CWBS","BOSS","IBRS")
'objects
Dim OutlineCode As OutlineCode, LookupTable As LookupTable, LookupTableEntry As LookupTableEntry
'longs
Dim lngEntry As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set OutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  Set LookupTable = OutlineCode.LookupTable
  For lngEntry = 1 To LookupTable.Count
    If InStr(LookupTable(lngEntry).Description, strFind) > 0 Then
      Debug.Print LookupTable(lngEntry).Description
      LookupTable(lngEntry).Description = Replace(LookupTable(lngEntry).Description, strFind, strReplace)
      Debug.Print LookupTable(lngEntry).Description
    End If
  Next lngEntry
  
exit_here:
  On Error Resume Next
  Set OutlineCode = Nothing
  Set LookupTable = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptOutlineCodes", "RenameInsideOutlineCode", err, Erl)
  Resume exit_here
End Sub
