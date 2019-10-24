Attribute VB_Name = "cptCore_bas"
'<cpt_version>v1.6.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private oMSPEvents As cptEvents_cls

Sub cptStartEvents()
  Set oMSPEvents = New cptEvents_cls
End Sub

Sub cptStopEvents()
  Set oMSPEvents = Nothing
End Sub

Sub cptSpeed(blnOn As Boolean)

  Application.ScreenUpdating = Not blnOn
  Application.Calculation = pjAutomatic = Not blnOn

End Sub

Function cptGetUserForm(strModuleName As String) As UserForm
'NOTE: this only works if the form is loaded
'objects
Dim UserForm As Object
'strings
'longs
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For Each UserForm In VBA.UserForms
    If UserForm.Name = strModuleName Then
      Set cptGetUserForm = UserForm
      Exit For
    End If
  Next

exit_here:
  On Error Resume Next
  Set UserForm = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "GetModule()", err, Erl)
  Resume exit_here
End Function

Function cptGetControl(ByRef cptForm_frm As UserForm, strControlName As String) As control
'NOTE: this only works for loaded forms

  Set cptGetControl = cptForm_frm.Controls(strControlName)

End Function

Function cptGetUserFullName()
'used to add user's name to PowerPoint title slide
Dim objAllNames As Object, objIndName As Object

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  On Error Resume Next
  Set objAllNames = GetObject("Winmgmts:").instancesof("win32_networkloginprofile")
  For Each objIndName In objAllNames
    If Len(objIndName.FullName) > 0 Then
      cptGetUserFullName = objIndName.FullName
      Exit For
    End If
  Next

exit_here:
  On Error Resume Next
  Set objAllNames = Nothing
  Set objIndName = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetUserFullName", err, Erl)
  Resume exit_here

End Function

Function cptGetVersions() As String
'requires reference: Microsoft Scripting Runtime
Dim vbComponent As Object, strVersion As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      cptGetVersions = cptGetVersions & vbComponent.Name & ": " & strVersion & vbCrLf
    End If
  Next vbComponent

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetVersions", err, Erl)
  Resume exit_here

End Function

'<issue31>
Sub cptUpgrade(Optional strFileName As String)
'objects
Dim oStream As Object
Dim xmlHttpDoc As Object
'strings
Dim strNewFileName As String
Dim strModule As String
Dim strError As String
Dim strURL As String
'longs
Dim lngLine As Long
'integers
'doubles
'booleans
Dim blnExists As Boolean
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Len(strFileName) = 0 Then strFileName = "Core/cptUpgrades_frm.frm"

  'go get it
  strURL = strGitHub
  strURL = strURL & strFileName
  strFileName = Replace(cptRegEx(strFileName, "\/.*\.[A-z]{3}"), "/", "")
frx:
  Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
  xmlHttpDoc.Open "GET", strURL, False
  xmlHttpDoc.Send
  If xmlHttpDoc.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1 'adTypeBinary
    oStream.Write xmlHttpDoc.responseBody
    If Dir(cptDir & "\" & strFileName) <> vbNullString Then Kill cptDir & "\" & strFileName
    oStream.SaveToFile cptDir & "\" & strFileName
    oStream.Close
    'need to fetch the .frx first
    If Right(strURL, 4) = ".frm" Then
      strURL = Replace(strURL, ".frm", ".frx")
      strFileName = Replace(strFileName, ".frm", ".frx")
      GoTo frx
    ElseIf Right(strURL, 4) = ".frx" Then
      strURL = Replace(strURL, ".frx", ".frm")
      strFileName = Replace(strFileName, ".frx", ".frm")
    End If
  Else
    strError = strError & "- " & strFileName & vbCrLf
  End If

  'remove if exists
  strModule = Left(strFileName, InStr(strFileName, ".") - 1)
  blnExists = Not ThisProject.VBProject.VBComponents(strModule) Is Nothing
  If blnExists Then
    'Set vbComponent = ThisProject.VBProject.VBComponents("cptUpgrades_frm")
    Application.StatusBar = "Removing obsolete version of " & strModule
    strNewFileName = strModule & "_" & Format(Now, "hhnnss")
    ThisProject.VBProject.VBComponents(strModule).Name = strNewFileName
    DoEvents
    ThisProject.VBProject.VBComponents.remove ThisProject.VBProject.VBComponents(strNewFileName)
    cptCore_bas.cptStartEvents
    DoEvents
  End If

  'import the module
  Application.StatusBar = "Importing " & strFileName & "..."
  ThisProject.VBProject.VBComponents.import cptDir & "\" & strFileName
  DoEvents
  
  '<issue24> remove the whitespace added by VBE import/export
  With ThisProject.VBProject.VBComponents(strModule).CodeModule
    For lngLine = .CountOfDeclarationLines To 1 Step -1
      If Len(.Lines(lngLine, 1)) = 0 Then .DeleteLines lngLine, 1
    Next lngLine
  End With '</issue24>

  'MsgBox "The Upgrade Form was itself just upgraded. Please repeat your click.", vbInformation + vbOKOnly, "Upgraded the Upgrader"

exit_here:
  On Error Resume Next
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Application.StatusBar = ""
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptUpgrade", err, Erl)
  Resume exit_here

End Sub '<issue31>

Sub ShowCptAbout_frm()
'objects
'Dim frmAbout As UserForm
'Dim ctl As control
'strings
Dim strAbout As String
'longs
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If Not cptModuleExists("cptAbout_frm") Then '<issue19>
    MsgBox "Please re-run cptSetup() to restore the missing About form.", vbCritical + vbOKOnly, "Missing Form"
    GoTo exit_here
  End If '</issue19>

  'contact and license
  strAbout = vbCrLf & "The ClearPlan Toolbar" & vbCrLf
  strAbout = strAbout & "by ClearPlan Consulting, LLC" & vbCrLf & vbCrLf
  'strAbout = strAbout & "http://ClearPlanConsulting.com" & vbCrLf & vbCrLf
  strAbout = strAbout & "This software is provided free of charge," & vbCrLf
  strAbout = strAbout & "AS IS and without warranty." & vbCrLf
  strAbout = strAbout & "It is free to use, free to distribute with prior written consent from the developers/copyright holders and without modification." & vbCrLf & vbCrLf
  strAbout = strAbout & "All rights reserved." & vbCrLf & "Copyright 2019, ClearPlanConsulting, LLC"
'  Set frmAbout = cptGetUserForm("cptAbout_frm") '<issue19>
'  Set ctl = cptGetControl(frmAbout, "txtAbout") '<issue19>
'  ctl.Value = strAbout
  cptAbout_frm.txtAbout.Value = strAbout  '<issue19>

  'follow the project
  strAbout = vbCrLf & vbCrLf & "Follow the Project:" & vbCrLf & vbCrLf
  strAbout = strAbout & "http://GitHub.com/ClearPlan/cpt" & vbCrLf & vbCrLf
'  Set ctl = cptGetControl(frmAbout, "txtGitHub") '<issue19>
'  ctl.Value = strAbout
  cptAbout_frm.txtGitHub.Value = strAbout '<issue19>

  'show/hide
'  Set ctl = cptGetControl(frmAbout, "lblScoreboard") '<issue19>
'  ctl.Visible = IIf(Now < #10/24/2019#, False, True) '<issue19>
  cptAbout_frm.lblScoreBoard.Visible = IIf(Now < #10/24/2019#, False, True) '<issue19>
  cptAbout_frm.Show '<issue19>
'  frmAbout.Show '<issue19>

  '<issue19> added error handling
exit_here:
  On Error Resume Next
'  Set ctl = Nothing
'  Set frmAbout = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "ShowCptAbout_frm", err, Erl)
  Resume exit_here '</issue19>

End Sub

Function cptReferenceExists(strReference As String) As Boolean
'used to ensure a reference exists, returns boolean
Dim Ref As Object, blnExists As Boolean

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnExists = False

  For Each Ref In ThisProject.VBProject.References
    If Ref.Name = strReference Then
      blnExists = True
      Exit For
    End If
  Next Ref

  cptReferenceExists = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptReferenceExists", err, Erl)
  Resume exit_here
End Function

Sub cptGetReferences()
'prints the current uesr's selected references
'this would be used to troubleshoot with users real-time
'although simply runing setreferences would fix it
Dim Ref As Object

  For Each Ref In ThisProject.VBProject.References
	  Debug.Print Ref.Name & " (" & Ref.Description & ") " & Ref.FullPath
  Next Ref

End Sub

Function cptGetDirectory(strModule As String) As String
'this function retrieves the directory of the module from CurrentVersions.xml on gitHub
'objects
Dim xmlDoc As Object
Dim xmlNode As Object
'strings
Dim strDirectory As String
Dim strURL As String
'longs
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'the calling subroutine should catch the Not cptInternetIsConnected function before calling this

  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.ErrorCode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
  Else
    Set xmlNode = xmlDoc.SelectSingleNode("//Name[text()='" + strModule + "']").ParentNode.SelectSingleNode("Directory")
    strDirectory = xmlNode.Text
  End If

  cptGetDirectory = strDirectory

exit_here:
  On Error Resume Next
  Set xmlDoc = Nothing
  Set xmlNode = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetDirectory()", err, Erl)
  Resume exit_here
End Function

Sub cptGetEnviron()
'list the environment variables and their associated values
Dim lgIndex As Long

  For lgIndex = 1 To 200
    Debug.Print lgIndex & ": " & Environ(lgIndex)
  Next

End Sub

Function cptCheckReference(strReference As String) As Boolean
'this routine will be called ahead of any subroutine requiring a reference
'returns boolean and subroutine only proceeds if true

  On Error GoTo err_here

  cptCheckReference = True

  Select Case strReference
    'CommonProgramFiles
    Case "Office"
      If Not cptReferenceExists("Office") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\OFFICE16\MSO.DLL"
      End If
    Case "VBIDE"
      If Not cptReferenceExists("VBIDE") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
      End If
    Case "VBA"
      If Not cptReferenceExists("VBA") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
      End If
    Case "ADODB"
      If Not cptReferenceExists("ADODB") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\System\ado\msado15.dll"
      End If

    'Office Applications
    Case "Excel"
      If Not cptReferenceExists("Excel") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\EXCEL.EXE"
      End If
    Case "Outlook"
      If Not cptReferenceExists("Outlook") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSOUTL.OLB"
      End If
    Case "PowerPoint"
      If Not cptReferenceExists("PowerPoint") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSPPT.OLB"
      End If
    Case "MSProject"
      If Not cptReferenceExists("MSProject") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSPRJ.OLB"
      End If
    Case "Word"
      If Not cptReferenceExists("Word") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSWORD.OLB (Word)"
      End If

    'Windows Common
    Case "MSForms"
      If Not cptReferenceExists("MSForms") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\FM20.DLL"
      End If
    Case "Scripting"
      If Not cptReferenceExists("Scripting") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\scrrun.dll"
      End If
    Case "stdole"
      If Not cptReferenceExists("stdole") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\stdole2.tlb"
      End If
    Case "mscorlib"
      If Not cptReferenceExists("mscorlib") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
      End If
    Case "MSXML2"
      If Not cptReferenceExists("MSXML2") Then '</issue33>
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\msxml3.dll"
      End If
    Case Else
      cptCheckReference = False

  End Select

  If Not cptCheckReference Then
    MsgBox "Missing Reference: " & strReference, vbExclamation + vbOKOnly, "CP Tool Bar"
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  cptCheckReference = False
  Resume exit_here

End Function

Sub cptResetAll()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  Application.OpenUndoTransaction "Reset All"

  FilterClear
  GroupClear
  OptionsViewEx displaynameindent:=True, displaysummarytasks:=True, displayoutlinesymbols:=True
  SelectAll 'needed for master/sub projects
  Sort "ID"
  OutlineShowAllTasks
  SelectBeginning

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptResetAll", err, Erl)
  Resume exit_here

End Sub

Sub ShowCptUpgrades_frm()
'objects
Dim REMatch As Object
Dim REMatches As Object
Dim RE As Object
Dim oStream As Object
Dim xmlHttpDoc As Object
Dim arrDirectories As Object
Dim vbComponent As Object
Dim arrCurrent As Object, arrInstalled As Object
Dim xmlDoc As Object
Dim xmlNode As Object
Dim FindRecord As Object
'long
Dim lngItem As Long
'strings
Dim strBranch As String
Dim strFileName As String
Dim strInstVer As String
Dim strCurVer As String
Dim strURL As String
Dim strVersion As String
'booleans
Dim blnUpdatesAreAvailable As Boolean
'variants
Dim vCol As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'todo:user should still be able to check currently installed versions
  If Not cptInternetIsConnected Then
    MsgBox "You must be connected to the internet to perform updates.", vbInformation + vbOKOnly, "No Connection"
    GoTo exit_here
  End If

  If Not cptCheckReference("VBA") Or Not cptCheckReference("VBIDE") Then
    GoTo exit_here
  End If

  'get current versions
  Set arrCurrent = CreateObject("System.Collections.SortedList")
  Set arrDirectories = CreateObject("System.Collections.SortedList")
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.ErrorCode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
    GoTo exit_here
  Else
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      arrCurrent.Add xmlNode.SelectSingleNode("Name").Text, xmlNode.SelectSingleNode("Version").Text
      'Debug.Print xmlNode.SelectSingleNode("Name").Text & " - " & xmlNode.SelectSingleNode("Directory").Text
      arrDirectories.Add xmlNode.SelectSingleNode("Name").Text, xmlNode.SelectSingleNode("Directory").Text
    Next
  End If

  'get installed versions
  Set arrInstalled = CreateObject("System.Collections.SortedList")
  blnUpdatesAreAvailable = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      arrInstalled.Add vbComponent.Name, strVersion
    End If
  Next vbComponent
  Set vbComponent = Nothing

  '<issue31> if cptUpgrade_frm is updated, install it automatically
  If arrInstalled.Contains("cptUpgrades_frm") And arrCurrent.Contains("cptUpgrades_frm") Then
    If cptVersionStatus(arrInstalled("cptUpgrades_frm"), arrCurrent("cptUpgrades_frm")) <> "ok" Then
      Call cptUpgrade(arrDirectories("cptUpgrades_frm") & "/cptUpgrades_frm.frm") 'uri slash
      'update the version number in the array
      arrInstalled.Item("cptUpgrades_frm") = arrCurrent("cptUpgrades_frm")
    End If
  End If '</issue31>

  'populate the listbox header
  lngItem = 0
  cptUpgrades_frm.lboHeader.AddItem
  For Each vCol In Array("Module", "Directory", "Current", "Installed", "Status", "Type")
    cptUpgrades_frm.lboHeader.List(0, lngItem) = vCol
    lngItem = lngItem + 1
  Next vCol
  cptUpgrades_frm.lboHeader.Height = 16

  'populate the listbox
  cptUpgrades_frm.lboModules.Clear
  For lngItem = 0 To arrCurrent.Count - 1
    'If arrCurrent.getKey(lngItem) = "ThisProject" Then GoTo next_lngItem '</issue25'
    strCurVer = arrCurrent.getValueList()(lngItem)
    If arrInstalled.Contains(arrCurrent.getKey(lngItem)) Then
      strInstVer = arrInstalled.getValueList()(arrInstalled.indexOfKey(arrCurrent.getKey(lngItem)))
    Else
      strInstVer = "<not installed>"
    End If
    cptUpgrades_frm.lboModules.AddItem
    cptUpgrades_frm.lboModules.List(lngItem, 0) = arrCurrent.getKey(lngItem) 'module name
    cptUpgrades_frm.lboModules.List(lngItem, 1) = arrDirectories.getValueList()(lngItem) 'directory
    cptUpgrades_frm.lboModules.List(lngItem, 2) = strCurVer 'arrCurrent.getValueList()(lngItem) 'current version
    If arrInstalled.Contains(arrCurrent.getKey(lngItem)) Then 'installed version
      cptUpgrades_frm.lboModules.List(lngItem, 3) = strInstVer 'arrInstalled.getValueList()(arrInstalled.indexOfKey(arrCurrent.getKey(lngItem)))
    Else
      cptUpgrades_frm.lboModules.List(lngItem, 3) = "<not installed>"
    End If

    Select Case strInstVer 'cptUpgrades_frm.lboModules.List(lngItem, 3)
      Case Is = strCurVer 'cptUpgrades_frm.lboModules.List(lngItem, 2)
        cptUpgrades_frm.lboModules.List(lngItem, 4) = "< ok >"
      Case Is = "<not installed>"
        cptUpgrades_frm.lboModules.List(lngItem, 4) = "< install >"
      Case Is <> strCurVer 'cptUpgrades_frm.lboModules.List(lngItem, 2)
        cptUpgrades_frm.lboModules.List(lngItem, 4) = "< " & cptVersionStatus(strInstVer, strCurVer) & " >"
    End Select
    'capture the type while we're at it - could have just pulled the FileName
    Set FindRecord = xmlDoc.SelectSingleNode("//Name[text()='" + cptUpgrades_frm.lboModules.List(lngItem, 0) + "']").ParentNode.SelectSingleNode("Type")
    cptUpgrades_frm.lboModules.List(lngItem, 5) = FindRecord.Text
next_lngItem:
  Next lngItem
  
  'populate branches
  Set xmlHttpDoc = CreateObject("MSXML2.XMLHTTP.6.0")
  strURL = "https://api.github.com/repos/AronGahagan/cpt-dev/branches"
  xmlHttpDoc.Open "GET", strURL, False
  xmlHttpDoc.setRequestHeader "Content-Type", "application/json"
  xmlHttpDoc.setRequestHeader "Accept", "application/json"
  xmlHttpDoc.Send
  If xmlHttpDoc.Status = 200 And xmlHttpDoc.readyState = 4 Then
    Set RE = CreateObject("vbscript.regexp")
    With RE
      .MultiLine = False
      .Global = True
      .ignorecase = True
      '.Pattern = Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "[A-z0-9\-]*"
      .Pattern = Chr(34) & "name" & Chr(34) & ":" & Chr(34) & "[A-z0-9-.]*"
    End With
    Set REMatches = RE.Execute(xmlHttpDoc.responseText)
    cptUpgrades_frm.cboBranches.Clear
    For Each REMatch In REMatches
      cptUpgrades_frm.cboBranches.AddItem Replace(REMatch, Chr(34) & "name" & Chr(34) & ":" & Chr(34), "")
    Next
    cptUpgrades_frm.cboBranches.Value = "master"
  Else
    cptUpgrades_frm.cboBranches.Clear
    cptUpgrades_frm.cboBranches.AddItem "<unavailable>"
  End If
  
  cptUpgrades_frm.Show

exit_here:
  On Error Resume Next
  Set REMatch = Nothing
  Set REMatches = Nothing
  Set RE = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Application.StatusBar = ""
  Set arrDirectories = Nothing
  Set vbComponent = Nothing
  Set arrCurrent = Nothing
  Set arrInstalled = Nothing
  Set xmlDoc = Nothing
  Set xmlNode = Nothing
  Set FindRecord = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "ShowCptUpgrades_frm", err, Erl)
  Resume exit_here

End Sub

Sub cptSetReferences()
'this is a one-time shot to set all references currently required by the cp toolbar
Dim strDir As String

  On Error Resume Next

  'CommonProgramFiles
  strDir = Environ("CommonProgramFiles")
  If Not cptReferenceExists("Office") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\OFFICE16\MSO.DLL"
  End If
  If Not cptReferenceExists("VBIDE") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
	'todo: need win64 file path '<issue53>
    'C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB?
  End If
  If Not cptReferenceExists("VBA") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
  End If
  If Not cptReferenceExists("ADODB") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\System\ado\msado15.dll"
  End If

  'office applications
  strDir = Application.Path 'OR cptRegEx(environ("PATH"),"C\:.*Microsoft Office[A-z0-9\\]*;")
  If Not cptReferenceExists("Excel") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\EXCEL.EXE"
  End If
  If Not cptReferenceExists("Outlook") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSOUTL.OLB"
  End If
  If Not cptReferenceExists("PowerPoint") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSPPT.OLB"
  End If
  If Not cptReferenceExists("MSProject") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSPRJ.OLB"
  End If
  If Not cptReferenceExists("Word") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSWORD.OLB"
  End If

  'Windows Common
  If Not cptReferenceExists("MSForms") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\FM20.DLL"
  End If
  If Not cptReferenceExists("Scripting") Then
    ThisProject.VBProject.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll"
  End If
  If Not cptReferenceExists("stdole") Then
    ThisProject.VBProject.References.AddFromFile "C:\Windows\SysWOW64\stdole2.tlb"
  End If
  If Not cptReferenceExists("mscorlib") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
  End If
  '<issue33> added
  If Not cptReferenceExists("MSXML2") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\msxml3.dll"
  End If '</issue33>

End Sub

Sub cptSubmitIssue()
  If Not Application.FollowHyperlink("https://forms.office.com/Pages/ResponsePage.aspx?id=Ro5H7jf1GEu_K_zo12S-I41LrliPQfRIoKdHTo6ZR7RUQ0VSV1JBRU4xQ1E5VUkyQjE5RDcwQllWRSQlQCN0PWcu", , , True) Then
    Call cptSendMail("Issue")
  End If
End Sub

Sub cptSubmitRequest()
  If Not Application.FollowHyperlink("https://forms.office.com/Pages/ResponsePage.aspx?id=Ro5H7jf1GEu_K_zo12S-I41LrliPQfRIoKdHTo6ZR7RUNVBET1RGUzRWMzZHN0pYNFZBUjZCUzgzNSQlQCN0PWcu", , , True) Then
    Call cptSendMail("Request")
  End If
End Sub

Sub cptSubmitFeedback()
  If Not Application.FollowHyperlink("https://forms.office.com/Pages/ResponsePage.aspx?id=Ro5H7jf1GEu_K_zo12S-I41LrliPQfRIoKdHTo6ZR7RUNERTVDRISUhVVVFSWjBBMlVLQThCRFlHQiQlQCN0PWcu", , , True) Then
    Call cptSendMail("Feedback")
  End If
End Sub

Sub cptSendMail(strCategory As String)
'objects
Dim objOutlook As Object 'Outlook.Application
Dim MailItem As Object 'MailItem
'strings
Dim strHTML As String
Dim strURL As String
'longs
'integers
'doubles
'booleans
'variants
'dates

  'get outlook
  On Error Resume Next
  Set objOutlook = GetObject(, "Outlook.Application")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If objOutlook Is Nothing Then
    Set objOutlook = CreateObject("Outlook.Application")
  End If

  'create the email and set generic settings
  Set MailItem = objOutlook.CreateItem(0) 'olMailItem
  MailItem.To = "cpt@ClearPlanConsulting.com"
  MailItem.Importance = 2 'olImportanceHigh
  MailItem.Display

  'get strURL and message body
  Select Case strCategory
    Case "Issue"
      MailItem.Subject = "Issue: <enter brief summary of the issue>"
      strHTML = "<h3>Please Describe Your Environment:</h3><p>"
      strHTML = strHTML & "<i>Operating System</i>: [operating system]<p>"
      strHTML = strHTML & "<i>Microsoft Project Version</i>: [Standard / Professional] [Year]<p>"
      strHTML = strHTML & "<i>Do you have unfettered internet access (try opening <a href=""https://github.com/AronGahagan/cpt-dev/blob/master/README.md"">this page</a>)?</i> [Yes/No]<p>"
      strHTML = strHTML & "<h3>Please Describe the Issue:</h3><p><p>"
      strHTML = strHTML & "<i>Please be as detailed as possible: what were you trying to do, what selections did you make, describe the file you are working on, etc.</i><p>"
      strHTML = strHTML & "<h3>Please Include Screenshot(s):</h3><p>Please include any screenshot(s) of any error messages or anything else that might help us troubleshoot this issue for you.<p><p>"
      strHTML = strHTML & "<i>Thank you for helping us improve the ClearPlan Toolbar!</i>"
      MailItem.HTMLBody = strHTML & MailItem.HTMLBody
      
    Case "Request"
      MailItem.Subject = "Feature Request: <enter brief description of the feature>"
      strHTML = "<h3>Please Describe the Feature you are Requesting:</h3><p>&nbsp;<p>&nbsp;"
      strHTML = strHTML & "<i>Thank you for contributing to the ClearPlan Toolbar project!</i>"
      MailItem.HTMLBody = strHTML & MailItem.HTMLBody
      
    Case "Feedback"
      MailItem.Subject = "Feedback: <enter summary of feedback>"
      strHTML = "<h3>Feedback:</h3><p>&nbsp;<p>&nbsp;<i>We sincerely appreciate any and all constructive feedback. Thank you for contributing!</i>"
      MailItem.HTMLBody = strHTML & MailItem.HTMLBody
      
  End Select
  
exit_here:
  On Error Resume Next
  Set objOutlook = Nothing
  Set MailItem = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptSendMail", err, Erl)
  Resume exit_here
End Sub

Function cptRemoveIllegalCharacters(strText As String) As String
'written by Ryan Beard (RyanBeard@ClearPlanConsulting.com)
    Const cstrIllegals As String = "\,/,:,*,?,"",<,>,|"

    Dim lngCounter As Long
    Dim astrChars() As String

    astrChars() = Split(cstrIllegals, ",")

    For lngCounter = LBound(astrChars()) To UBound(astrChars())
        strText = Replace(strText, astrChars(lngCounter), vbNullString)
    Next lngCounter

    cptRemoveIllegalCharacters = strText

End Function

Sub cptWrapItUp()
'objects
'strings
'longs
Dim lgLevel As Long
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Dynamic Filter"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===

  cptSpeed True 'speed up
  Application.OpenUndoTransaction "WrapItUp"
  'FilterClear 'do not reset, keep autofilters
  'GroupClear 'do not reset, applies to groups to
  OptionsViewEx displaysummarytasks:=True
  SelectAll
  OutlineShowAllTasks
  OutlineShowTasks pjTaskOutlineShowLevelMax
  'pjTaskOutlineShowLevelMax = 65,535 = do not use
  For lgLevel = 20 To pjTaskOutlineShowLevel2 Step -1
    OutlineShowTasks lgLevel
  Next lgLevel
  SelectBeginning

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction
  cptSpeed False
  Exit Sub

'no_tasks:
'  MsgBox "This project has no tasks to collapse.", vbExclamation + vbOKOnly, "WrapItUp"
'  GoTo exit_here

err_here:
  Call cptHandleErr("cptCore_bas", "cptWrapItUp", err, Erl)
  Resume exit_here
End Sub

Function cptVersionStatus(strInstalled As String, strCurrent As String) As String
'objects
'strings
'longs
Dim lngVersion As Long
'integers
'booleans
'variants
Dim aCurrent As Variant
Dim aInstalled As Variant
Dim vVersion As Variant
Dim vLevel As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'clean the versions - include all three levels
  For Each vVersion In Array(strInstalled, strCurrent)
    'following line doesn't remove non-numeric characters
    vVersion = cptRegEx(CStr(vVersion), "([0-9].*.?){1,3}")
    If Len(vVersion) - Len(Replace(vVersion, ".", "")) < 1 Then
      vVersion = vVersion & ".0"
    End If
    If Len(vVersion) - Len(Replace(vVersion, ".", "")) < 2 Then
      vVersion = vVersion & ".0"
    End If
    If lngVersion = 0 Then
      aInstalled = Split(vVersion, ".")
    Else
      aCurrent = Split(vVersion, ".")
    End If
    lngVersion = lngVersion + 1
  Next

  'figure out the things
  For Each vLevel In Array(0, 1, 2)
    If aCurrent(vLevel) <> aInstalled(vLevel) Then
      cptVersionStatus = Choose(vLevel + 1, "major", "minor", "patch")
      If CLng(aCurrent(vLevel)) > CLng(aInstalled(vLevel)) Then '<issue62>
        cptVersionStatus = cptVersionStatus & " upgrade"
      Else
        cptVersionStatus = cptVersionStatus & " downgrade"
      End If
      Exit For
    End If
  Next vLevel

  If cptVersionStatus = "" Then
    cptVersionStatus = "ok"
  Else
    cptVersionStatus = "install " & cptVersionStatus
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptVersionStatus", err, Erl)
  Resume exit_here

End Function
