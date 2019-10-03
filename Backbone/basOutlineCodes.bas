Attribute VB_Name = "basOutlineCodes"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False

Sub CreateCode(lngOutlineCode As Long, strOutlineCodeName As String)
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
  If frmOutlineCode.txtNameIt.BorderColor = 255 Then GoTo exit_here

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
    frmOutlineCode.lblProgress.Width = ((lngTask - 1) / lngTasks) * frmOutlineCode.lblStatus.Width
    frmOutlineCode.lblStatus.Caption = Format(lngTask - 1, "#,##0") & " / " & Format(lngTasks, "#,##0") & " (" & Format((lngTask - 1) / lngTasks, "0%") & ") [" & Format(Now - tmr, "hh:nn:ss") & "]"
  Next Task
  CustomOutlineCodeEditEx lngOutlineCode, OnlyLeaves:=True, OnlyLookUpTableCodes:=True
  frmOutlineCode.lblStatus.Caption = "Complete."
  Application.StatusBar = "Complete."
  frmOutlineCode.cmdCancel.Caption = "Done"
  
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

Sub ShowFrmOutlineCodes()
'longs
Dim lngCode As Long, lngOutlineCode As Long
'strings
Dim strOutlineCode As String, strOutlineCodeName As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  frmOutlineCode.cboOutlineCodes.Clear

  For lngCode = 1 To 10
    strOutlineCode = "Outline Code" & lngCode
    lngOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
    strOutlineCodeName = Application.CustomFieldGetName(lngOutlineCode)
    frmOutlineCode.cboOutlineCodes.AddItem
    If Len(strOutlineCodeName) > 0 Then
      strOutlineCode = strOutlineCode & " (" & strOutlineCodeName & ")"
      'frmOutlineCode.cboOutlineCodes.Column(1, lngCode - 1) = "(" & strOutlineCodeName & ")"
    End If
    frmOutlineCode.cboOutlineCodes.Column(0, lngCode - 1) = strOutlineCode
  Next lngCode
  
  frmOutlineCode.txtNameIt = ""
  frmOutlineCode.cmdCancel.Caption = "Cancel"
  frmOutlineCode.cboOutlineCodes.Value = frmOutlineCode.cboOutlineCodes.List(0)
  frmOutlineCode.Show False
  frmOutlineCode.cboOutlineCodes.SetFocus

exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call HandleErr("basOutlineCodes", "ShowFrmOutlineCodes", err)
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
  Call HandleErr("basOutlineCodes", "RenameInsideOutlineCode", err)
  Resume exit_here
End Sub
