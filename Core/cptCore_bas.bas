Attribute VB_Name = "cptCore_bas"
'<cpt_version>v1.4.3</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private oMSPEvents As cptEvents_cls

Sub cptStartEvents()
  Set oMSPEvents = New cptEvents_cls
End Sub

'<issue19> attempting to avoid trailing 1
Sub cptStopEvents()
  Set oMSPEvents = Nothing
End Sub

Sub cptSpeed(blnOn As Boolean)

  Application.ScreenUpdating = Not blnOn
  Application.Calculation = pjAutomatic = Not blnOn

End Sub

Function cptGetUserForm(strModuleName As String) As UserForm
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
  Call cptHandleErr("cptCore_bas", "GetModule()", err)
  Resume exit_here
End Function

Function cptGetControl(ByRef cptForm_frm As UserForm, strControlName As String) As control
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
  Call cptHandleErr("cptCore_bas", "cptGetUserFullName", err)
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
  Call cptHandleErr("cptCore_bas", "cptGetVersions", err)
  Resume exit_here

End Function

Sub ShowCptAbout_frm()
'objects
Dim UserForm As UserForm
Dim ctl As control
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
  Set UserForm = ThisProject.VBProject.VBComponents("cptAbout_frm") '<issue19>
  Set ctl = UserForm.Controls("txtAbout") '<issue19>
  ctl.Value = strAbout
  'cptAbout_frm.txtAbout.Value = strAbout  '<issue19>

  'follow the project
  strAbout = vbCrLf & vbCrLf & "Follow the Project:" & vbCrLf & vbCrLf
  strAbout = strAbout & "http://GitHub.com/ClearPlan/cpt" & vbCrLf & vbCrLf
  Set ctl = cptGetControl(UserForm, "txtGitHub") '<issue19>
  ctl.Value = strAbout
  'cptAbout_frm.txtGitHub.Value = strAbout '<issue19>

  'show/hide
  Set ctl = cptGetControl(UserForm, "lblScoreboard") '<issue19>
  ctl.Visible = IIf(Now < #10/24/2019#, False, True) '<issue19>
  'cptAbout_frm.lblScoreBoard.Visible = IIf(Now < #10/24/2019#, False, True) '<issue19>
  'cptAbout_frm.Show '<issue19>
  UserForm.Show '<issue19>
  
  '<issue19> added error handling
exit_here:
  On Error Resume Next
  Set ctl = Nothing
  Set UserForm = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "ShowCptAbout_frm", err)
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
  Call cptHandleErr("cptCore_bas", "cptReferenceExists", err)
  Resume exit_here
End Function

Sub cptGetReferences()
'prints the current uesr's selected references
'this would be used to troubleshoot with users real-time
'although simply runing setreferences would fix it
Dim Ref As Object

  For Each Ref In ThisProject.VBProject.References
    Debug.Print Ref.Name & " (" & Ref.Description & ")" & Ref.FullPath
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
  Call cptHandleErr("cptCore_bas", "cptGetDirectory()", err)
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
      If Not cptReferenceExists("") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
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
  OutlineShowAllTasks
  Sort "ID"
  SelectBeginning

exit_here:
  On Error Resume Next
  Application.CloseUndoTransaction

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptResetAll", err)
  Resume exit_here

End Sub

Sub ShowCptUpgrades_frm()
'objects
Dim arrDirectories As Object
Dim vbComponent As Object
Dim arrCurrent As Object, arrInstalled As Object
Dim xmlDoc As Object, xmlNode As Object, xmlHttpDoc As Object, FindRecord As Object
Dim oStream As Object
'long
Dim lngItem As Long
'strings
Dim strInstVer As String
Dim strCurVer As String
Dim strURL As String, strVersion As String
'booleans
Dim blnUpdatesAreAvailable As Boolean
'variants
Dim vCol As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'user should still be able to check currently installed versions
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
  For lngItem = 0 To arrCurrent.count - 1
    'If arrCurrent.getKey(lngItem) = "ThisProject" Then GoTo next_lngItem '</issue25'
    strCurVer = arrCurrent.getValueList()(lngItem)
    If arrInstalled.contains(arrCurrent.getKey(lngItem)) Then
      strInstVer = arrInstalled.getValueList()(arrInstalled.indexofkey(arrCurrent.getKey(lngItem)))
    Else
      strInstVer = "<not installed>"
    End If
    cptUpgrades_frm.lboModules.AddItem
    cptUpgrades_frm.lboModules.List(lngItem, 0) = arrCurrent.getKey(lngItem) 'module name
    cptUpgrades_frm.lboModules.List(lngItem, 1) = arrDirectories.getValueList()(lngItem) 'directory
    cptUpgrades_frm.lboModules.List(lngItem, 2) = strCurVer 'arrCurrent.getValueList()(lngItem) 'current version
    If arrInstalled.contains(arrCurrent.getKey(lngItem)) Then 'installed version
      cptUpgrades_frm.lboModules.List(lngItem, 3) = strInstVer 'arrInstalled.getValueList()(arrInstalled.indexofkey(arrCurrent.getKey(lngItem)))
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

  cptUpgrades_frm.Show False

exit_here:
  On Error Resume Next
  Set arrDirectories = Nothing
  Set vbComponent = Nothing
  Set arrCurrent = Nothing
  Set arrInstalled = Nothing
  Set xmlDoc = Nothing
  Set xmlNode = Nothing
  Set xmlHttpDoc = Nothing
  Set FindRecord = Nothing
  Set oStream = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "UpdatesAreAvailable", err)
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

End Sub

Sub cptHandleErr(strModule As String, strProcedure As String, objErr As ErrObject)
'common error handling prompt
Dim strMsg As String

    strMsg = "Uh oh! Please contact cpt@ClearPlanConsulting.com for assistance if needed." & vbCrLf & vbCrLf
    strMsg = strMsg & "Error Source:" & vbCrLf
    strMsg = strMsg & "Module: " & strModule & vbCrLf
    strMsg = strMsg & "Procedure: " & strProcedure & vbCrLf & vbCrLf
    strMsg = strMsg & "Error Code:" & vbCrLf
    strMsg = strMsg & err.Number & ": " & err.Description
    MsgBox strMsg, vbExclamation + vbOKOnly, "Unknown Error"

End Sub

Sub cptSubmitIssue()
  Application.OpenBrowser "https://forms.office.com/Pages/ResponsePage.aspx?id=Ro5H7jf1GEu_K_zo12S-I41LrliPQfRIoKdHTo6ZR7RUQ0VSV1JBRU4xQ1E5VUkyQjE5RDcwQllWRSQlQCN0PWcu"
End Sub

Sub cptSubmitRequest()
  Application.OpenBrowser "https://forms.office.com/Pages/ResponsePage.aspx?id=Ro5H7jf1GEu_K_zo12S-I41LrliPQfRIoKdHTo6ZR7RUNVBET1RGUzRWMzZHN0pYNFZBUjZCUzgzNSQlQCN0PWcu"
End Sub

Sub cptSubmitFeedback()
  Application.OpenBrowser "https://forms.office.com/Pages/ResponsePage.aspx?id=Ro5H7jf1GEu_K_zo12S-I41LrliPQfRIoKdHTo6ZR7RUNERTVDRISUhVVVFSWjBBMlVLQThCRFlHQiQlQCN0PWcu"
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
  Call cptHandleErr("cptCore_bas", "cptWrapItUp", err)
  Resume exit_here
End Sub

Public Function cptBuildRibbonTab()
Dim ribbonXML As String
Dim lngCleanUp As Long

  'buuld ClearPlan Ribbon Tab XML
  ribbonXML = ribbonXML + vbCrLf & "<mso:tab id=""tCommon"" label=""ClearPlan"" >" 'insertBeforeQ=""mso:TabTask"">"

  'common tools
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""custom_view"" label=""View"" visible=""true"">"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:SummaryTasks"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:NameIndent"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:OutlineSymbolsShow"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:AutoFilterProject"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:FilterClear"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:SplitViewCreate"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetAll"" label=""Reset All"" imageMso=""FilterClear"" onAction=""cptResetAll"" visible=""true"" size=""large"" />"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bWrapItUp"" label=""WrapItUp"" imageMso=""CollapseAll"" onAction=""cptWrapItUp"" visible=""true"" size=""large"" />"   'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'task counters
  If cptModuleExists("cptCountTasks_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCount"" label=""Count"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountAll"" label=""All Tasks"" imageMso=""NumberInsert"" onAction=""cptCountTasksAll"" visible=""true""/>" 'SelectWholeLayout
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountVisible"" label=""Visible Tasks"" imageMso=""NumberInsert"" onAction=""cptCountTasksVisible"" visible=""true""/>" 'SelectRows
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountSelected"" label=""Selected Tasks"" imageMso=""NumberInsert"" onAction=""cptCountTasksSelected"" visible=""true""/>" 'SelectTaskCell
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'text tools
  If cptModuleExists("cptText_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gTextTools"" label=""Text"" visible=""true"" >"
    If cptModuleExists("cptDynamicFilter_bas") And cptModuleExists("cptDynamicFilter_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDynamicFilter"" label=""Dynamic Filter"" imageMso=""FilterBySelection"" onAction=""ShowcptDynamicFilter_frm"" visible=""true"" size=""large"" />"
    End If
    If cptModuleExists("cptText_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbText"" size=""large"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdvancedTextTools"" label=""Advanced"" imageMso=""AdvancedFilterDialog"" onAction=""ShowcptText_frm"" />" 'visible=""true""
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mText"">"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Utilities"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""cptBulkPrepend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""cptBulkAppend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""cptMyReplace"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""cptEnumerate"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""cptTrimTaskNames"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process (WIP)"" imageMso=""DuplicateSelectedSlides"" onAction=""cptReplicateProcess"" visible=""true"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""cptFindDuplicateTaskNames"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetRowHeight"" label=""Reset Row Height"" imageMso=""RowHeight"" onAction=""cptResetRowHeight"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
    Else
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTextTools"" label=""Tools"" imageMso=""TextBoxInsert"" visible=""true"" size=""large"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""cptBulkPrepend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""cptBulkAppend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""cptMyReplace"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""cptEnumerate"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""cptTrimTaskNames"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process"" imageMso=""DuplicateSelectedSlides"" onAction=""cptReplicateProcess"" visible=""true"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""cptFindDuplicateTaskNames"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetRowHeight"" label=""Reset Row Height"" imageMso=""RowHeight"" onAction=""cptResetRowHeight"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'trace tools
  If cptModuleExists("cptCriticalPathTools_bas") Or cptModuleExists("cptCriticalPath_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCPA"" label=""Trace"" visible=""true"">"
    If cptModuleExists("cptCriticalPathTools_bas") And cptModuleExists("cptCriticalPath_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbTrace"" size=""large"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" imageMso=""TaskDrivers"" label=""Driving Path"" onAction=""DrivingPaths"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTrace"">"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Export"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPowerPoint"" label="">> PowerPoint"" imageMso=""SlideNew"" onAction=""cptExportCriticalPathSelected"" />"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
    Else
      If cptModuleExists("cptCriticalPath_bas") Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" label=""Driving Path"" imageMso=""TaskDrivers"" onAction=""DrivingPaths"" visible=""true"" size=""large"" />"
      End If
      If cptModuleExists("cptCriticalPathTools_bas") Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bExport"" label="">> PowerPoint"" imageMso=""SlideNew"" onAction=""cptExportCriticalPathSelected"" visible=""true"" size=""large"" />"
      End If
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'status
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gStatus"" label=""Status"" visible=""true"" >"
  If cptModuleExists("cptStatusSheet_bas") And cptModuleExists("cptStatusSheet_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bStatusSheet"" label=""Create Status Sheet"" imageMso=""DateAndTimeInsertOneNote"" onAction=""ShowcptStatusSheet_frm"" visible=""true""/>"
  End If
  If cptModuleExists("cptSmartDuration_frm") And cptModuleExists("cptSmartDuration_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSmartDuration"" label=""Smart Duration"" imageMso=""CalendarToolSelectDate"" onAction=""SmartDuration"" visible=""true""/>"
  End If
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'snapshots

  'resource allocation
  If cptModuleExists("cptResourceDemand_bas") And cptModuleExists("cptResourceDemand_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gResourceDemand"" label=""Resource Demand"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResourceDemandExcel"" label=""Export to Excel"" imageMso=""Chart3DColumnChart"" onAction=""ShowFrmExportResourceDemand"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'scenarios

  'compare

  'metrics

  'integration
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gIntegration"" label=""Integration"" visible=""true"" >"
  If cptModuleExists("cptIMSCobraExport_bas") And cptModuleExists("cptIMSCobraExport_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCOBRA"" label=""COBRA Export Tool"" imageMso=""Export"" onAction=""Export_IMS"" visible=""true""/>"
  End If
  'mpm
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'bcr

  'about
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gHelp"" label=""Help"" visible=""true"" >"
  If cptInternetIsConnected Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mHelp"" label=""Help"" imageMso=""Help"" visible=""true"" size=""large"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Upgrades"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUpdate"" label=""Check for Upgrades"" imageMso=""PreviousUnread"" onAction=""ShowCptUpgrades_frm"" />" 'supertip=" & Chr(34) & strSuperTip & Chr(34) & "
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Contribute"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bIssue"" label=""Submit an Issue"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitIssue"" visible=""true"" />" 'supertip=" & Chr(34) & strSuperTip & Chr(34) & "
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bRequest"" label=""Submit a Feature Request"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitRequest"" visible=""true"" />" 'supertip=" & Chr(34) & strSuperTip & Chr(34) & "
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFeedback"" label=""Submit Other Feedback"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitFeedback"" visible=""true"" />" 'supertip=" & Chr(34) & strSuperTip & Chr(34) & "
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAbout"" onAction=""ShowcptAbout_frm""  size=""large"" visible=""true""  label=""About"" imageMso=""Info"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  ribbonXML = ribbonXML + vbCrLf & "</mso:tab>"

  'Debug.Print "<mso:customUI ""xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >" & ribbonXML
  cptBuildRibbonTab = ribbonXML

End Function

Function cptIncrement(ByRef lgCleanUp As Long) As Long
  lgCleanUp = lgCleanUp + 1
  cptIncrement = lgCleanUp
End Function

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

  'for now, assumes semantic version control - see https://semver.org/
  'useing yy.dd.mm would be easier, would show age of the release
  'and does it really matter if we 'have the lates' patch if we know we have *the latest*?

  'todo: indicate install, (major upgrade, minor upgrade, patch, downgrade) available

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
      If aCurrent(vLevel) > aInstalled(vLevel) Then
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
  Call cptHandleErr("cptCore_bas", "cptVersionStatus", err)
  Resume exit_here

End Function

Function cptRegEx(strText As String, strRegEx As String) As String
Dim RE As Object, REMatch As Variant, REMatches As Object
Dim strMatch As String

    On Error GoTo err_here

    Set RE = CreateObject("vbscript.regexp")
    With RE
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
        .Pattern = strRegEx
    End With

    Set REMatches = RE.Execute(strText)
    For Each REMatch In REMatches
      strMatch = REMatch
      Exit For
    Next
    cptRegEx = strMatch

exit_here:
    On Error Resume Next
    Set RE = Nothing
    Set REMatches = Nothing
    Exit Function
err_here:
  If err.Number = 5 Then
    cptRegEx = ""
    err.Clear
  End If
  Resume exit_here
End Function
