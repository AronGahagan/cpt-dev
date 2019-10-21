Attribute VB_Name = "cptAdmin_bas"
'>no cpt version - not for release<
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptCreateCurrentVersionsXML(Optional strRepo As String)
'objects
Dim arrModules As Object
Dim arrTypes As Object
Dim oStream As Object, vbComponent As Object 'adodb.stream
'strings
Dim strMsg As String
Dim strModule As String
Dim strDirectory As String
Dim strFile As String
Dim strXML As String, strVersion As String, strFileName As String
Dim strBranch As String
'longs
Dim lngItem As Long
Dim lngFile As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'confirm repo selected
  If Len(frmGitVBA.cboRepo.Value) = 0 Or Dir(frmGitVBA.cboRepo.Value & "\.git\", vbDirectory) = vbNullString Then
    MsgBox "Please select a valid git repo.", vbExclamation + vbOKOnly, "Nope"
    frmGitVBA.cboRepo.SetFocus
    frmGitVBA.cboRepo.DropDown
    GoTo exit_here
  Else
    strRepo = frmGitVBA.cboRepo.Value
  End If

  'confirm branch selected
  If Len(frmGitVBA.cboBranch.Value) = 0 Then
    MsgBox "Please select a valid branch.", vbExclamation + vbOKOnly, "Nope"
    frmGitVBA.cboBranch.SetFocus
    frmGitVBA.cboBranch.DropDown
    GoTo exit_here
  Else
    strBranch = Replace(Replace(frmGitVBA.cboBranch.Value, Chr(32), ""), "*", "")
  End If

  'measure twice...
  If MsgBox("Writing to repo (branch): " & vbCrLf & strRepo & " (" & strBranch & ")", vbQuestion + vbYesNo, "Please Confirm") = vbNo Then GoTo exit_here

  'use arrTypes
  Set arrTypes = CreateObject("System.Collections.SortedList")
  arrTypes.Add 1, ".bas"
  arrTypes.Add 2, ".cls"
  arrTypes.Add 3, ".frm"
  arrTypes.Add 100, ".cls"
  
  '<issue18> sort the list to limit merge conflicts - added
  Set arrModules = CreateObject("System.Collections.SortedList")
  For Each vbComponent In ThisProject.VBProject.VBComponents
    If vbComponent.Name = "cptAdmin_bas" Then GoTo next_vbComponent
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      arrModules.Add vbComponent.Name, vbComponent.Name
    End If
next_vbComponent:
  Next vbComponent
  '</issue18>
  
  'write xml
  strXML = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf
  strXML = strXML & "<Modules>" & vbCrLf
  '<issue18>removed
  'For Each vbComponent In ThisProject.VBProject.VBComponents - removed
  '  If vbComponent.Name = "cptAdmin_bas" Then GoTo next_vbComponent - removed
  '  If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then - removed
  '</issue18>
  For lngItem = 0 To arrModules.Count - 1
    Set vbComponent = ThisProject.VBProject.VBComponents(arrModules.getKey(lngItem))
    Debug.Print arrModules.getKey(lngItem)
    strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
    strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
    strXML = strXML & String(1, vbTab) & "<Module>" & vbCrLf
    strModule = Replace(vbComponent.Name, cptRegEx(vbComponent.Name, "_frm|_bas|_cls"), "")
    strXML = strXML & String(2, vbTab) & "<Name>" & vbComponent.Name & "</Name>" & vbCrLf
    strXML = strXML & String(2, vbTab) & "<FileName>" & vbComponent.Name & arrTypes(CInt(vbComponent.Type)) & "</FileName>" & vbCrLf
    strXML = strXML & String(2, vbTab) & "<Version>" & strVersion & "</Version>" & vbCrLf
    strXML = strXML & String(2, vbTab) & "<Type>" & vbComponent.Type & "</Type>" & vbCrLf
    strDirectory = Replace(vbComponent.Name, cptRegEx(vbComponent.Name, "_frm|_bas|_cls"), "")
    strXML = strXML & String(2, vbTab) & "<Directory>" & Replace(cptSetDirectory(CStr(vbComponent.Name)), "\", "") & "</Directory>" & vbCrLf
    strXML = strXML & String(1, vbTab) & "</Module>" & vbCrLf
  Next lngItem
  '<issue18>  End If - removed
'next_vbComponent: - removed
  'Next vbComponent - removed
  '</issue18>
  strXML = strXML & "</Modules>" & vbCrLf
  
  'ensure correct branch is active
  frmGitVBA.txtNotes.Value = frmGitVBA.txtNotes.Value & vbCrLf & String(53, "-") & vbCrLf & Redirect("git", "-C " & strRepo & " checkout " & Replace(Replace(frmGitVBA.cboBranch.Value, Chr(32), ""), "*", ""))
  Call gitScrollDown
  
  'write to the file
  Set oStream = CreateObject("ADODB.Stream")
  oStream.Type = 2 'adTypeText
  oStream.Charset = "utf-8"
  strFileName = strRepo & "CurrentVersions.xml"
  oStream.Open
  oStream.WriteText strXML
  oStream.SaveToFile strFileName, 2 'adSaveCreateOverWrite
  oStream.Close
  Set oStream = Nothing
  
  frmGitVBA.txtNotes.Value = frmGitVBA.txtNotes.Value & vbCrLf & String(53, "-") & vbCrLf & Redirect("git", "-C " & strRepo & " add CurrentVersions.xml")
  Call gitScrollDown
  
exit_here:
  On Error Resume Next
  Set arrModules = Nothing
  Set arrTypes = Nothing
  Set vbComponent = Nothing
  If oStream.State <> adStateClosed Then oStream.Close
  Set oStream = Nothing
  Exit Sub
  
err_here:
  Call cptHandleErr("cptAdmin_bas", "CreateCurrentVersionXML", err)
  Resume exit_here

End Sub

Sub cptDocument()
'objects
Dim vbComponent As vbComponent
Dim xlApp As Object, Workbook As Object, Worksheet As Object
'strings
Dim strModule As String
Dim strProcName As String
'longs
Dim lngSLOC As Long
Dim lngLines As Long
Dim lngLine As Long
Dim lngRow As Long
Dim lngCountDecl As Long
'integers
'booleans
'variants
Dim arrHeader As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'get excel
  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  
  'set the header
  arrHeader = Array("Ribbon Group", "Module", "SLOC", "Procedure", "SLOC", "Directory", "HelpDoc", "Author")
  Worksheet.Range(Worksheet.[A1], Worksheet.[A1].Offset(0, UBound(arrHeader))) = arrHeader
  Worksheet.Columns.AutoFit
  
  lngRow = 2
  
  For Each vbComponent In ThisProject.VBProject.VBComponents
    strModule = vbComponent.Name
    Debug.Print "working on " & strModule & "..."
    If strModule = "ThisProject" Or Left(strModule, 3) = "cpt" Then
      With vbComponent.CodeModule
        lngCountDecl = .CountOfDeclarationLines
        lngLines = .CountOfLines
        Worksheet.Cells(lngRow, 2) = .Name
        Worksheet.Cells(lngRow, 3) = .CountOfLines
        strProcName = .ProcOfLine(lngCountDecl + 1, 0) '0 = vbext_pk_Proc
        Worksheet.Cells(lngRow, 4) = strProcName
        Worksheet.Cells(lngRow, 5) = .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
        lngSLOC = lngSLOC + .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
        Worksheet.Columns.AutoFit
        For lngLine = lngCountDecl + 1 To lngLines
          If .ProcOfLine(lngLine, 0) <> strProcName Then '0 = vbext_pk_Proc
            strProcName = .ProcOfLine(lngLine, 0) '0 = vbext_pk_Proc
            lngRow = lngRow + 1
            Worksheet.Cells(lngRow, 2) = strModule
            Worksheet.Cells(lngRow, 4) = strProcName
            Worksheet.Cells(lngRow, 5) = .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
            lngSLOC = lngSLOC + .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
            Worksheet.Columns.AutoFit
            If lngRow > 10 Then xlApp.ActiveWindow.ScrollRow = lngRow - 10
          End If
        Next
      End With
      lngRow = lngRow + 1
      If lngRow > 10 Then xlApp.ActiveWindow.ScrollRow = lngRow - 10
    End If
  Next vbComponent
  
  xlApp.ActiveWindow.ScrollRow = 2
  
  MsgBox "Documented." & vbCrLf & vbCrLf & "(" & Format(lngSLOC, "#,##0") & " SLOC)", vbInformation + vbOKOnly, "Documenter"
  
exit_here:
  On Error Resume Next
  Set vbComponent = Nothing
  Set xlApp = Nothing
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set xlApp = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "Document", err)
  Resume exit_here
End Sub

Sub cptCheckAllVersions()
Dim vbComponent As vbComponent

  For Each vbComponent In ThisProject.VBProject.VBComponents
    If Left(vbComponent.Name, 3) = "cpt" Then
      Debug.Print vbComponent.Name & ": " & vbComponent.CodeModule.Lines(1, 1)
    End If
  Next vbComponent
  Set vbComponent = Nothing
  
End Sub

Function cptSetDirectory(strComponentName As String) As String
'strings
Dim strDirectory As String
      
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'remove the prefix
  strDirectory = Replace(strComponentName, "cpt", "")
  'remove the suffix
  If InStr(strDirectory, "_") > 0 Then strDirectory = Left(strDirectory, InStr(strDirectory, "_") - 1)
  Select Case strDirectory
    'Setup
    Case "Setup"
      strDirectory = ""
    Case "Patch"
      strDirectory = ""
    'Core
    Case "BrowseFolder"
      strDirectory = "Core"
    Case "About"
      strDirectory = "Core"
    Case "Upgrades"
      strDirectory = "Core"
    Case "ThisProject"
      strDirectory = "Core"
    Case "Events"
      strDirectory = "Core"
    'count
    Case "CountTasks"
      strDirectory = "Count"
    'Integration
    Case "IMSCobraExport"
      strDirectory = "Integration"
    'TextTools
    Case "DynamicFilter"
      strDirectory = "Text"
    'Status
    Case "SmartDuration"
      strDirectory = "Status"
    Case "StatusSheet"
      strDirectory = "Status"
    'Trace
    Case "CriticalPath"
      strDirectory = "Trace"
    Case "CriticalPathTools"
      strDirectory = "Trace"
    Case Else
      
        
  End Select
  
  cptSetDirectory = strDirectory & "\"
  
exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptAdmin_bas", "cptSetDirectory()", err)
  Resume exit_here

End Function

Sub cptSQL(strFile As String, Optional strFilter As String)
'objects
Dim cn As ADODB.Connection, rst As ADODB.Recordset
'strings
Dim strRecord As String
Dim strFields As String
Dim strCon As String, strDir As String, strSQL As String
'longs
Dim lngField As Long
'integers
'doubles
'booleans
'variants
'dates

  'cpt-export-resource-userfields.adtg
  'cpt-status-sheet.adtg
  'cpt-status-sheet-userfields.adtg
  'cpt-data-dictionary.adtg
  'git-vba-repo.adtg
  'vba-backup-modules.adtg

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strFile = Environ("USERPROFILE") & "\cpt-backup\settings\" & strFile

  If Dir(strFile) = vbNullString Then
    Debug.Print "Invalid file: " & strFile
    GoTo exit_here
  End If
    
  With CreateObject("ADODB.Recordset")
    .Open strFile
    'get field names
    For lngField = 0 To .Fields.Count - 1
      strFields = strFields & .Fields(lngField).Name & " | "
    Next lngField
    Debug.Print strFields
    'get records
    If Not .EOF Then .MoveFirst
    Do While Not .EOF
      strRecord = ""
      For lngField = 0 To .Fields.Count - 1
        strRecord = strRecord & .Fields(lngField) & " | "
      Next lngField
      Debug.Print strRecord
      .MoveNext
    Loop
  End With
  
exit_here:
  On Error Resume Next
  If rst.State Then rst.Close
  Set rst = Nothing
  If cn.State Then cn.Close
  Set cn = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "cptSQL", err, Erl)
  Resume exit_here
End Sub

Sub cptCreateIssue()
'ABANDONED - requires OAuth / username password
'objects
Dim xmlHttpDoc As Object
'strings
Dim strURL As String
Dim strBody As String
Dim strTitle As String
Dim strContent As String
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'create the post title and body
  'include windows os, office version, msproj version
  'include date, username, feature and feature version
  strTitle = "Title"
  strBody = "Body"
  strContent = "{"
  strContent = strContent & Chr(34) & "title" & Chr(34) & ": " & Chr(34) & strTitle & Chr(34) & ","
  strContent = strContent & Chr(34) & "body" & Chr(34) & ": " & Chr(34) & strBody & Chr(34) & ","
  strContent = strContent & Chr(34) & "assignees" & Chr(34) & ": " & "[]" & ","
  strContent = strContent & Chr(34) & "milestone" & Chr(34) & ": " & "None" & ","
  strContent = strContent & Chr(34) & "labels" & Chr(34) & ": " & "[]"
  strContent = strContent & "}"
  Debug.Print strContent
  
  'create and send the json object
  Set xmlHttpDoc = CreateObject("WinHttp.WinHttpRequest.5.1")
  strURL = "https://api.github.com/repos/AronGahagan/cpt-dev/issues"
  xmlHttpDoc.Open "POST", strURL, False
  xmlHttpDoc.setRequestHeader "Content-Type", "application/json"
  xmlHttpDoc.setRequestHeader "Accept", "application/vnd.github.machine-man-preview+json"
  xmlHttpDoc.Send strContent
  If xmlHttpDoc.Status = 201 Then
    'include issue number
    MsgBox xmlHttpDoc.Status & ": " & xmlHttpDoc.StatusText
  Else
    Debug.Print xmlHttpDoc.responseText
    MsgBox xmlHttpDoc.Status & ": " & xmlHttpDoc.StatusText, vbExclamation + vbOKOnly, "Issue Not Posted"
  End If
exit_here:
  On Error Resume Next
  Set xmlHttpDoc = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptAdmin_bas", "cptCreateIssue", err, Erl)
  Resume exit_here
End Sub
