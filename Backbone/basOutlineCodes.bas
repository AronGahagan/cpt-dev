Attribute VB_Name = "basOutlineCodes"
'from https://technet.microsoft.com/en-us/subscriptions/index/ms453877(v=office.12).aspx

Option Explicit

Sub CreateCode(lgOutlineCode As Long, strOutlineCodeName As String)
'objects
Dim objOutlineCode As OutlineCode, objLookupTable As LookupTable, objLookupTableEntry As LookupTableEntry
Dim Task As Task, xlApp As Excel.Application
'strings
Dim strWBS As String, strParent As String, strChild As String
'doubles
Dim dblUID As Double, dblTasks As Double, dblTask As Double
'variants
Dim aOutlineCode As Variant, tmr As Date

  tmr = Now

  'first name the field and create the code
  On Error Resume Next
  ActiveProject.OutlineCodes(strOutlineCodeName).Delete
  On Error GoTo 0 'err_here
  
  Application.Calculation = pjManual
  Application.ScreenUpdating = False
  
  Set objOutlineCode = CreateOutlineCode(lgOutlineCode, strOutlineCodeName)
  Set objLookupTable = objOutlineCode.LookupTable
  Set objLookupTableEntry = objLookupTable.AddChild("1")
  objLookupTableEntry.Description = ActiveProject.Tasks.UniqueID(0).Name
  
  dblTasks = ActiveProject.Tasks.Count
  
  'dblTasks+1 so we can capture Project-Level "1.0"
  ReDim aOutlineCode(dblTasks + 1, 1)
  
  Set xlApp = CreateObject("Excel.Application")
  If MsgBox("Is this for the CWBS?", vbQuestion + vbYesNo, "Confirm Structure") = vbYes Then
    dblTask = 1
    aOutlineCode(0, 0) = "1"
    aOutlineCode(0, 1) = objLookupTableEntry.UniqueID
    For Each Task In ActiveProject.Tasks
      strParent = Left(Task.WBS, InStrRev(Task.WBS, ".") - 1)
      strChild = Mid(Task.WBS, InStrRev(Task.WBS, ".") + 1)
      dblUID = xlApp.WorksheetFunction.VLookup(strParent, aOutlineCode, 2, False)
      Set objLookupTableEntry = objLookupTable.AddChild(strChild, dblUID)
      objLookupTableEntry.Description = Task.Name
      dblTask = dblTask + 1
      aOutlineCode(dblTask - 1, 0) = Task.WBS
      aOutlineCode(dblTask - 1, 1) = objLookupTableEntry.UniqueID
      If Not Task.Summary Then Task.SetField lgOutlineCode, Task.WBS
      frmOutlineCode.lblProgress.Width = ((dblTask - 1) / dblTasks) * frmOutlineCode.lblStatus.Width
      frmOutlineCode.lblStatus.Caption = Format(dblTask - 1, "#,##0") & " / " & Format(dblTasks, "#,##0") & " (" & Format((dblTask - 1) / dblTasks, "0%") & ") [" & Format(Now - tmr, "hh:nn:ss") & "]"
      'Application.StatusBar = Format(dblTask - 1, "#,##0") & " / " & Format(dblTasks, "#,##0") & " (" & Format((dblTask - 1) / dblTasks, "0%") & ")"
    Next Task
  Else
    For Each Task In ActiveProject.Tasks
      dblTask = dblTask + 1
      If Task.OutlineLevel = 1 Then
        strChild = Task.WBS
        Set objLookupTableEntry = objLookupTable.AddChild(strChild)
      Else
        strChild = Mid(Task.WBS, InStrRev(Task.WBS, ".") + 1)
        strParent = Left(Task.WBS, InStrRev(Task.WBS, ".") - 1)
        dblUID = xlApp.WorksheetFunction.VLookup(strParent, aOutlineCode, 2, False)
        Set objLookupTableEntry = objLookupTable.AddChild(strChild, dblUID)
      End If
      objLookupTableEntry.Description = Task.Name
      aOutlineCode(dblTask - 1, 0) = Task.WBS
      aOutlineCode(dblTask - 1, 1) = objLookupTableEntry.UniqueID
      Task.SetField lgOutlineCode, Task.WBS
      frmOutlineCode.lblProgress.Width = ((dblTask - 1) / dblTasks) * frmOutlineCode.lblStatus.Width
      frmOutlineCode.lblStatus.Caption = Format(dblTask - 1, "#,##0") & " / " & Format(dblTasks, "#,##0") & " (" & Format((dblTask - 1) / dblTasks, "0%") & ") [" & Format(Now - tmr, "hh:nn:ss") & "]"
    Next Task
  End If
  
  frmOutlineCode.lblStatus.Caption = "Complete."
  Application.StatusBar = "Complete."
  frmOutlineCode.cmdCancel.Caption = "Done"
  
exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Application.Calculation = pjAutomatic
  Application.ScreenUpdating = True
  Set objOutlineCode = Nothing
  Set objLookupTable = Nothing
  Set objLookupTableEntry = Nothing
  Set Task = Nothing
  Set xlApp = Nothing
  Exit Sub
err_here:
  MsgBox err.Number & ": " & err.Description, vbExclamation + vbOKOnly, "Error"
  Resume exit_here
End Sub

Function CreateOutlineCode(lgOutlineCode As Long, strOutlineCodeName As String) As OutlineCode
Dim objOutlineCode As OutlineCode, dblLevel As Double
  
  ActiveProject.OutlineCodes(1).Delete
  
  Set objOutlineCode = ActiveProject.OutlineCodes.Add(lgOutlineCode, strOutlineCodeName)
  For dblLevel = 1 To 10
    objOutlineCode.CodeMask.Add Sequence:=pjCustomOutlineCodeNumbers, Length:="Any", Separator:="."
  Next dblLevel
  objOutlineCode.SortOrder = pjListOrderAscending
  Set CreateOutlineCode = objOutlineCode

End Function

Sub ShowFrmOutlineCodes()
Dim dblCode As Double, lgOutlineCode As Long
Dim strOutlineCode As String, strOutlineCodeName As String

  On Error GoTo err_here
  
  frmOutlineCode.cboOutlineCodes.Clear

  For dblCode = 1 To 10
    strOutlineCode = "Outline Code" & dblCode
    lgOutlineCode = Application.FieldNameToFieldConstant(strOutlineCode)
    strOutlineCodeName = Application.CustomFieldGetName(lgOutlineCode)
    frmOutlineCode.cboOutlineCodes.AddItem
    If Len(strOutlineCodeName) > 0 Then
      strOutlineCode = strOutlineCode & " (" & strOutlineCodeName & ")"
      'frmOutlineCode.cboOutlineCodes.Column(1, dblCode - 1) = "(" & strOutlineCodeName & ")"
    End If
    frmOutlineCode.cboOutlineCodes.Column(0, dblCode - 1) = strOutlineCode
  Next dblCode
  
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

Sub RenameOutlineCode(strOutlineCode As String, strFind As String, strReplace As String)
'usage: Call RenameOutlineCode("CWBS","BOSS","IBRS")
Dim OutlineCode As OutlineCode, LookupTable As LookupTable, LookupTableEntry As LookupTableEntry
Dim lgEntry As Long

  Set OutlineCode = ActiveProject.OutlineCodes(strOutlineCode)
  Set LookupTable = OutlineCode.LookupTable
  For lgEntry = 1 To LookupTable.Count
    If InStr(LookupTable(lgEntry).Description, strFind) > 0 Then
      Debug.Print LookupTable(lgEntry).Description
      LookupTable(lgEntry).Description = Replace(LookupTable(lgEntry).Description, strFind, strReplace)
      Debug.Print LookupTable(lgEntry).Description
    End If
  Next lgEntry
  Set OutlineCode = Nothing
  Set LookupTable = Nothing
End Sub
