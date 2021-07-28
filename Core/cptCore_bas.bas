Attribute VB_Name = "cptCore_bas"
'<cpt_version>v1.9.6</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private oMSPEvents As cptEvents_cls
#If Win64 And VBA7 Then
  Private Declare PtrSafe Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare PtrSafe Function SetPrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
  Private Declare Function GetPrivateProfileString lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare Function SetPrivateProfileString lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

Sub cptStartEvents()
  Set oMSPEvents = New cptEvents_cls
End Sub

Sub cptStopEvents()
  Set oMSPEvents = Nothing
End Sub

Sub cptSpeed(blnOn As Boolean)

  Application.Calculation = pjAutomatic = Not blnOn
  Application.ScreenUpdating = Not blnOn

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
  Call cptHandleErr("cptCore_bas", "GetModule()", Err, Erl)
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
  Call cptHandleErr("cptCore_bas", "cptGetUserFullName", Err, Erl)
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
  Call cptHandleErr("cptCore_bas", "cptGetVersions", Err, Erl)
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
  blnExists = cptModuleExists(strModule)
  If blnExists Then
    'Set vbComponent = ThisProject.VBProject.VBComponents("cptUpgrades_frm")
    Application.StatusBar = "Removing obsolete version of " & strModule
    strNewFileName = strModule & "_" & Format(Now, "hhnnss")
    ThisProject.VBProject.VBComponents(strModule).Name = strNewFileName
    DoEvents
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(strNewFileName)
    cptCore_bas.cptStartEvents
    DoEvents
  End If

  'import the module
  Application.StatusBar = "Importing " & strFileName & "..."
  ThisProject.VBProject.VBComponents.Import cptDir & "\" & strFileName
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
  Call cptHandleErr("cptCore_bas", "cptUpgrade", Err, Erl)
  Resume exit_here

End Sub '<issue31>

Sub cptShowAbout_frm()
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
  strAbout = strAbout & "This software is provided free of charge," & vbCrLf
  strAbout = strAbout & "AS IS and without warranty." & vbCrLf
  strAbout = strAbout & "It is free to use, free to distribute with prior written consent from the developers/copyright holders and without modification." & vbCrLf & vbCrLf
  strAbout = strAbout & "All rights reserved." & vbCrLf & "Copyright 2019, ClearPlanConsulting, LLC"
  cptAbout_frm.txtAbout.Value = strAbout  '<issue19>

  'follow the project
  strAbout = vbCrLf & vbCrLf & "Follow the Project:" & vbCrLf & vbCrLf
  strAbout = strAbout & "http://GitHub.com/ClearPlan/cpt" & vbCrLf & vbCrLf
  cptAbout_frm.txtGitHub.Value = strAbout '<issue19>

  'show/hide
  cptAbout_frm.lblScoreBoard.Visible = IIf(Now <= #10/25/2019#, False, True) '<issue19>
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b1" EWR > MSY
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b2" MSY > EWR
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b3" 'EWR > SAN
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b4" 'SAN > EWR
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b5" 'EWR > NAS
  cptAbout_frm.lblScoreBoard.Caption = "t0 : b6" 'NAS > EWR
  cptAbout_frm.Show '<issue19>

exit_here:
  On Error Resume Next
  
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowAbout_frm", Err, Erl)
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
  Call cptHandleErr("cptCore_bas", "cptReferenceExists", Err, Erl)
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
  Call cptHandleErr("cptCore_bas", "cptGetDirectory()", Err, Erl)
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
Dim strDir As String
Dim strRegEx As String

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
  End Select
    
  'Office Applications
  strRegEx = "C\:.*Microsoft Office[A-z0-9\\]*Office[0-9]{2}"
  strDir = Replace(cptRegEx(Environ("PATH"), strRegEx), ";", "")
  Select Case strReference
    Case "Excel"
      If Not cptReferenceExists("Excel") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\EXCEL.EXE"
      End If
    Case "Outlook"
      If Not cptReferenceExists("Outlook") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSOUTL.OLB"
      End If
    Case "PowerPoint"
      If Not cptReferenceExists("PowerPoint") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSPPT.OLB"
      End If
    Case "MSProject"
      If Not cptReferenceExists("MSProject") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSPRJ.OLB"
      End If
    Case "Word"
      If Not cptReferenceExists("Word") Then
        ThisProject.VBProject.References.AddFromFile strDir & "\MSWORD.OLB"
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
    Case "MSComctlLib"
      If Not cptReferenceExists("MSComctlLib") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\MSCOMCTL.OCX"
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
  Dim rstSettings As Object 'ADODB.Recordset
  'strings
  Dim strOutlineLevel As String
  Dim strSettings As String
  Dim strFile As String
  'longs
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  Dim lngLevel As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Reset All"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===

  cptSpeed True

  strFile = cptDir & "\settings\cpt-reset-all.adtg"
  If Dir(strFile) <> vbNullString Then
    Set rstSettings = CreateObject("ADODB.Recordset")
    rstSettings.Open strFile
    rstSettings.MoveFirst
    lngSettings = rstSettings(0)
    cptSaveSetting "ResetAll", "Settings", CStr(lngSettings)
    lngOutlineLevel = rstSettings(1)
    cptSaveSetting "ResetAll", "OutlineLevel", CStr(lngOutlineLevel)
    rstSettings.Close
    Kill strFile
  Else
    strSettings = cptGetSetting("ResetAll", "Settings")
    If Len(strSettings) > 0 Then lngSettings = CLng(strSettings)
    strOutlineLevel = cptGetSetting("ResetAll", "OutlineLevel")
    If Len(strOutlineLevel) > 0 Then lngOutlineLevel = CLng(strOutlineLevel)
  End If
  
  If lngSettings > 0 Then
    'parse and apply
    If lngSettings >= 128 Then 'outline symbols
      OptionsViewEx displayoutlinesymbols:=True
      lngSettings = lngSettings - 128
    End If
    If lngSettings >= 64 Then 'display name indent
      OptionsViewEx displaynameindent:=True
      lngSettings = lngSettings - 64
    End If
    If lngSettings >= 32 Then 'clear filter
      FilterClear
      lngSettings = lngSettings - 32
    End If
    If lngSettings >= 16 Then 'sort by ID
      Sort "ID", , , , , , False, True
      lngSettings = lngSettings - 16
    End If
    If lngSettings >= 8 Then 'expand all tasks
      OptionsViewEx displaysummarytasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
          Sort "ID", , , , , , False, True
          OutlineShowAllTasks
        Else
          SelectBeginning
          GoTo exit_here
        End If
      End If
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      lngSettings = lngSettings - 8
    Else 'expand to specific level
      OptionsViewEx displaysummarytasks:=True
      On Error Resume Next
      If Not OutlineShowAllTasks Then
        If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
          Sort "ID", , , , , , False, True
          OutlineShowAllTasks
        Else
          SelectBeginning
          GoTo exit_here
        End If
      End If
      If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
      OutlineShowTasks pjTaskOutlineShowLevelMax
      For lngLevel = 20 To lngOutlineLevel Step -1
        OutlineShowTasks lngLevel
      Next lngLevel
    End If
    If lngSettings >= 4 Then 'show summaries
      OptionsViewEx displaysummarytasks:=True
      lngSettings = lngSettings - 4
    Else
      OptionsViewEx displaysummarytasks:=False
    End If
    If lngSettings >= 2 Then 'clear group
      GroupClear
      lngSettings = lngSettings - 2
    End If
    If lngSettings >= 1 Then 'hide inactive
      SetAutoFilter "Active", pjAutoFilterFlagYes
    End If
  Else 'prompt for defaults
    Call cptShowResetAll_frm
  End If
  

exit_here:
  On Error Resume Next
  If rstSettings.State Then rstSettings.Close
  Set rstSettings = Nothing
  cptSpeed False

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptResetAll", Err, Erl)
  Resume exit_here

End Sub

Sub cptShowResetAll_frm()
  'objects
  Dim rstSettings As Object 'ADODB.Recordset
  'strings
  Dim strOutlineLevel As String
  Dim strSettings As String
  Dim strFile As String
  'longs
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "Reset All"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  'populate the outline level picklist
  For lngOutlineLevel = 1 To 20
    cptResetAll_frm.cboOutlineLevel.AddItem lngOutlineLevel
  Next lngOutlineLevel
  'default to 2
  cptResetAll_frm.cboOutlineLevel.Value = 2
  
  strFile = cptDir & "\settings\cpt-reset-all.adtg"
  If Dir(strFile) <> vbNullString Then
    'get saved settings
    Set rstSettings = CreateObject("ADODB.Recordset")
    rstSettings.Open strFile
    rstSettings.MoveFirst
    lngSettings = rstSettings(0)
    cptSaveSetting "ResetAll", "Settings", CStr(lngSettings)
    lngOutlineLevel = rstSettings(1)
    cptSaveSetting "ResetAll", "OutlineLevel", CStr(lngOutlineLevel)
    rstSettings.Close
    Kill strFile
  Else
    strSettings = cptGetSetting("ResetAll", "Settings")
    If Len(strSettings) > 0 Then lngSettings = CLng(strSettings)
    strOutlineLevel = cptGetSetting("ResetAll", "OutlineLevel")
    If Len(strOutlineLevel) > 0 Then lngOutlineLevel = CLng(strOutlineLevel)
  End If
    
  If lngSettings > 0 Then
    'parse and update the form
    With cptResetAll_frm
      If lngSettings >= 128 Then
        .chkOutlineSymbols = True
        lngSettings = lngSettings - 128
      End If
      If lngSettings >= 64 Then
        .chkIndent = True
        lngSettings = lngSettings - 64
      End If
      If lngSettings >= 32 Then
        .chkFilter = True
        lngSettings = lngSettings - 32
      End If
      If lngSettings >= 16 Then
        .chkSort = True
        lngSettings = lngSettings - 16
      End If
      If lngSettings >= 8 Then
        .optShowAllTasks = True
        lngSettings = lngSettings - 8
      Else
        .optOutlineLevel = True
        .cboOutlineLevel.Value = IIf(lngOutlineLevel = 0, 2, lngOutlineLevel)
      End If
      If lngSettings >= 5 Then
        .chkSummaries = True
        lngSettings = lngSettings - 4
      End If
      If lngSettings >= 2 Then
        .chkGroup = True
        lngSettings = lngSettings - 2
      End If
      If lngSettings >= 1 Then
        .chkActiveOnly = True
      End If
    End With
  End If
  
  cptResetAll_frm.Show False
  
exit_here:
  On Error Resume Next
  Set rstSettings = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowResetAll_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptShowUpgrades_frm()
'objects
Dim REMatch As Object
Dim REMatches As Object
Dim RE As Object
Dim oStream As Object
Dim xmlHttpDoc As Object
Dim rstStatus As Object 'ADODB.Recordset
Dim vbComponent As Object
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

  'update references needed before downloading updates
  Call cptSetReferences

  'todo: user should still be able to check currently installed versions
  If Not cptInternetIsConnected Then
    MsgBox "You must be connected to the internet to perform updates.", vbInformation + vbOKOnly, "No Connection"
    GoTo exit_here
  End If

  'set up the recordset
  Set rstStatus = CreateObject("ADODB.Recordset")
  rstStatus.Fields.Append "Module", 200, 200
  rstStatus.Fields.Append "Directory", 200, 200
  rstStatus.Fields.Append "Current", 200, 200
  rstStatus.Fields.Append "Installed", 200, 200
  rstStatus.Fields.Append "Status", 200, 200
  rstStatus.Open
  
  'get current versions
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
      rstStatus.AddNew
      rstStatus(0) = xmlNode.SelectSingleNode("Name").Text
      rstStatus(1) = xmlNode.SelectSingleNode("Directory").Text
      rstStatus(2) = xmlNode.SelectSingleNode("Version").Text
      rstStatus.Update
    Next xmlNode
  End If

  'get installed versions
  blnUpdatesAreAvailable = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("'<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      rstStatus.MoveFirst
      rstStatus.Find "Module='" & vbComponent.Name & "'", , 1
      If Not rstStatus.EOF Then
        rstStatus(3) = strVersion
        rstStatus(4) = cptVersionStatus(rstStatus(2), strVersion)
        rstStatus.Update
      End If
    End If
  Next vbComponent
  Set vbComponent = Nothing

  'if cptUpgrade_frm is updated, install it automatically



  rstStatus.MoveFirst
  rstStatus.Find "Module='cptUpgrades_frm'", , 1
  If cptVersionStatus(rstStatus(2), rstStatus(3)) <> "ok" Then
    Call cptUpgrade(rstStatus(1) & "/cptUpgrades_frm.frm")
    rstStatus(3) = rstStatus(2)
    rstStatus.Update
  End If
  
  'cannot auto upgrade cptCore_bas because this is the cptCore_bas module so use cptPatch_bas
  'if cptPatch_bas is updated, install it automatically and run it
  rstStatus.MoveFirst
  rstStatus.Find "Module='cptPatch_bas'", , 1
  If cptVersionStatus(rstStatus(2), rstStatus(3)) <> "ok" Then
    Call cptUpgrade(rstStatus(1) & "/cptPatch_bas.bas")
    rstStatus(3) = rstStatus(2)
    rstStatus.Update
    
    '/=== temp fix to cptPatch_bas private/public issue ===\
    'patch code goes here
    Application.StatusBar = "Applying patch 21.04.10..."
    If Not cptReferenceExists("VBScript_RegExp_55") Then
      ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\System32\vbscript.dll\3"
    End If
    '\=== temp fix to cptPatch_bas private/public issue ===/
    
  End If



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
  rstStatus.Sort = "Module"
  rstStatus.MoveFirst
  lngItem = 0
  Do While Not rstStatus.EOF
    strCurVer = rstStatus(2)
    If Not IsNull(rstStatus(3)) Then
      strInstVer = rstStatus(3)
    Else
      strInstVer = "<not installed>"
    End If
    cptUpgrades_frm.lboModules.AddItem
    cptUpgrades_frm.lboModules.List(lngItem, 0) = rstStatus(0) 'module name
    cptUpgrades_frm.lboModules.List(lngItem, 1) = rstStatus(1) 'directory
    cptUpgrades_frm.lboModules.List(lngItem, 2) = strCurVer 'current version
    cptUpgrades_frm.lboModules.List(lngItem, 3) = strInstVer 'installed version
    
    Select Case strInstVer
      Case Is = strCurVer
        cptUpgrades_frm.lboModules.List(lngItem, 4) = "< ok >"
      Case Is = "<not installed>"
        cptUpgrades_frm.lboModules.List(lngItem, 4) = "< install >"
      Case Is <> strCurVer
        cptUpgrades_frm.lboModules.List(lngItem, 4) = "< " & cptVersionStatus(strInstVer, strCurVer) & " >"
    End Select
    'capture the type while we're at it - could have just pulled the FileName
    Set FindRecord = xmlDoc.SelectSingleNode("//Name[text()='" + cptUpgrades_frm.lboModules.List(lngItem, 0) + "']").ParentNode.SelectSingleNode("Type")
    cptUpgrades_frm.lboModules.List(lngItem, 5) = FindRecord.Text
next_lngItem:
    lngItem = lngItem + 1
    rstStatus.MoveNext
  Loop
  
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
      .IgnoreCase = True
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
  If rstStatus.State Then rstStatus.Close
  Set rstStatus = Nothing
  Set REMatch = Nothing
  Set REMatches = Nothing
  Set RE = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Application.StatusBar = ""
  Set vbComponent = Nothing
  Set xmlDoc = Nothing
  Set xmlNode = Nothing
  Set FindRecord = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowUpgrades_frm", Err, Erl)
  Resume exit_here

End Sub

Sub cptSetReferences()
'this is a one-time shot to set all references currently required by the cp toolbar
Dim oExcel As Object
Dim strDir As String
Dim strRegEx As String

  On Error Resume Next

  'CommonProgramFiles
  strDir = Environ("CommonProgramFiles")
  If Not cptReferenceExists("Office") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\OFFICE16\MSO.DLL"
  End If
  If Not cptReferenceExists("VBIDE") Then
    #If Not Win64 Then
      ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
    #Else
      ThisProject.VBProject.References.AddFromFile "C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
    #End If
  End If
  If Not cptReferenceExists("VBA") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
  End If
  If Not cptReferenceExists("ADODB") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\System\ado\msado15.dll"
  End If

  'office applications
  Set oExcel = CreateObject("Excel.Application")
  If oExcel Is Nothing Then 'weird installation or Excel not installed
    MsgBox "Microsoft Office installation is not detetcted. Some features may not operate as expected." & vbCrLf & vbCrLf & "Please contact cpt@ClearPlanConsulting.com for specialized assistance.", vbCritical + vbOKOnly, "Microsoft Office Compatibility"
    GoTo windows_common
  Else
    strDir = oExcel.Path
  End If
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
windows_common:
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
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\mscorlib.dll"
  End If
  If Not cptReferenceExists("MSComctlLib") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\MSCOMCTL.OCX"
  End If
  If Not cptReferenceExists("MSXML2") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\msxml3.dll"
  End If
  If Not cptReferenceExists("VBScript_RegExp_55") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\System32\vbscript.dll\3"
  End If
  
exit_here:
  On Error Resume Next
  If Not oExcel Is Nothing Then oExcel.Quit
  Set oExcel = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptSetReferences", Err, Erl)
  Resume exit_here

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
  Call cptHandleErr("cptCore_bas", "cptSendMail", Err, Erl)
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

Sub cptWrapItUp(Optional lngOutlineLevel As Long)
'objects
'strings
'longs
Dim lngLevel As Long
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "WrapItUp"
    GoTo exit_here
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  
  If lngOutlineLevel = 0 Then
    'check for a saved setting
    If Dir(cptDir & "\settings\cpt-reset-all.adtg") <> vbNullString Then
      With CreateObject("ADODB.Recordset")
        .Open cptDir & "\settings\cpt-reset-all.adtg"
        .MoveFirst
        lngOutlineLevel = .Fields(1)
        .Close
      End With
    Else
      lngOutlineLevel = 2
    End If
  End If
  
  cptSpeed True 'speed up
  Application.OpenUndoTransaction "WrapItUp"
  'FilterClear 'do not reset, keep autofilters
  'GroupClear 'do not reset, applies to groups to
  OptionsViewEx displaysummarytasks:=True
  SelectAll
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    Else
      SelectBeginning
      GoTo exit_here
    End If
  End If
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  OutlineShowTasks pjTaskOutlineShowLevelMax
  'pjTaskOutlineShowLevelMax = 65,535 = do not use
  For lngLevel = 20 To lngOutlineLevel Step -1
    OutlineShowTasks lngLevel
  Next lngLevel
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
  Call cptHandleErr("cptCore_bas", "cptWrapItUp", Err, Erl)
  Resume exit_here
End Sub

Sub cptWrapItUpAll()
  '===
  'Validate users selected view type
  If ActiveProject.Application.ActiveWindow.ActivePane.View.Type <> pjTaskItem Then
    MsgBox "Please select a View with a Task Table.", vbInformation + vbOKOnly, "WrapItUp"
    Exit Sub
  End If
  'Validate users selected window pane - select the task table if not active
  If ActiveProject.Application.ActiveWindow.ActivePane.Index <> 1 Then
    ActiveProject.Application.ActiveWindow.TopPane.Activate
  End If
  '===
  OptionsViewEx displaysummarytasks:=True
  On Error Resume Next
  If Not OutlineShowAllTasks Then
    If MsgBox("In order to Expand All Tasks, the Outline Structure must be retained in the Sort order. OK to Sort by ID?", vbExclamation + vbYesNo, "Conflict: Sort") = vbYes Then
      Sort "ID", , , , , , False, True
      OutlineShowAllTasks
    Else
      SelectBeginning
    End If
  End If

End Sub
Sub cptWrapItUp1()
  Call cptWrapItUp(1)
End Sub
Sub cptWrapItUp2()
  Call cptWrapItUp(2)
End Sub
Sub cptWrapItUp3()
  Call cptWrapItUp(3)
End Sub
Sub cptWrapItUp4()
  Call cptWrapItUp(4)
End Sub
Sub cptWrapItUp5()
  Call cptWrapItUp(5)
End Sub
Sub cptWrapItUp6()
  Call cptWrapItUp(6)
End Sub
Sub cptWrapItUp7()
  Call cptWrapItUp(7)
End Sub
Sub cptWrapItUp8()
  Call cptWrapItUp(8)
End Sub
Sub cptWrapItUp9()
  Call cptWrapItUp(9)
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
      If Len(aInstalled(vLevel)) = 0 Then
        cptVersionStatus = cptVersionStatus & " upgrade"
      ElseIf CLng(aCurrent(vLevel)) > CLng(aInstalled(vLevel)) Then '<issue62>
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
  Call cptHandleErr("cptCore_bas", "cptVersionStatus", Err, Erl)
  Resume exit_here

End Function

Sub cptFilterReapply()
  Dim strCurrentFilter As String
  strCurrentFilter = ActiveProject.CurrentFilter
  ScreenUpdating = False
  FilterApply "All Tasks"
  'todo: how to reapply a custom AutoFilter?
  On Error Resume Next
  If Not FilterApply(strCurrentFilter) Then
    MsgBox "Cannot reapply a Custom AutoFilter", vbInformation + vbOKCancel, "Reapply Filter"
  End If
  ScreenUpdating = True
End Sub

Sub cptGroupReapply()
  Dim strCurrentGroup As String
  Dim lngUID As Long
  lngUID = 0
  On Error Resume Next
  lngUID = ActiveSelection.Tasks(1).UniqueID
  strCurrentGroup = ActiveProject.CurrentGroup
  ScreenUpdating = False
  ActiveWindow.TopPane.Activate
  GroupApply "No Group"
  'todo: how to reapply an AutoFilter group?
  On Error Resume Next
  If Not GroupApply(strCurrentGroup) Then
    MsgBox "Cannot reapply a Custom AutoFilter Group", vbInformation + vbOKCancel, "Reapply Group"
  End If
  ScreenUpdating = True
  If lngUID > 0 Then EditGoTo ActiveProject.Tasks.UniqueID(lngUID).ID
End Sub

Function cptSaveSetting(strFeature As String, strSetting As String, strValue As String) As Boolean
  Dim strSettingsFile As String, lngWorked As Long
  strSettingsFile = cptDir & "\settings\cpt-settings.ini"
  lngWorked = SetPrivateProfileString(strFeature, strSetting, strValue, strSettingsFile)
  If lngWorked Then
    cptSaveSetting = True
  Else
    cptSaveSetting = False
  End If
End Function

Function cptGetSetting(strFeature As String, strSetting As String) As String
  Dim strSettingsFile As String, strReturned As String, lngSize As Long, lngWorked As Long
  strSettingsFile = cptDir & "\settings\cpt-settings.ini"
  strReturned = Space(128)
  lngSize = Len(strReturned)
  lngWorked = GetPrivateProfileString(strFeature, strSetting, "", strReturned, lngSize, strSettingsFile)
  If lngWorked Then
    cptGetSetting = Left$(strReturned, lngWorked)
  Else
    cptGetSetting = ""
  End If
End Function

Function cptFilterExists(strFilter As String) As Boolean
  'objects
  Dim oFilter As MSProject.Filter

  On Error Resume Next
  Set oFilter = ActiveProject.TaskFilters(strFilter)
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  cptFilterExists = Not oFilter Is Nothing
  
exit_here:
  On Error Resume Next
  Set oFilter = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptFilterExists", Err, Erl)
  Resume exit_here
End Function

Sub cptCreateFilter(strFilter As String)
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Select Case strFilter
    Case "Marked"
      FilterEdit Name:="Marked", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Marked", test:="equals", Value:="Yes", ShowInMenu:=True, showsummarytasks:=False
      
  End Select
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptCreateFilter", Err, Erl)
  Resume exit_here
End Sub
