Attribute VB_Name = "cptCore_bas"
'<cpt_version>v1.13.6</cpt_version>
Option Explicit
Private oMSPEvents As cptEvents_cls
#If Win64 And VBA7 Then
  Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
  Private Declare PtrSafe Function SetPrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
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

Function cptGetUserForm(strModuleName As String) As MSForms.UserForm
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  Call cptHandleErr("cptCore_bas", "cptGetUserForm", Err, Erl)
  Resume exit_here
End Function

Function cptGetControl(ByRef cptForm_frm As MSForms.UserForm, strControlName As String) As MSForms.Control
'NOTE: this only works for loaded forms

  Set cptGetControl = cptForm_frm.Controls(strControlName)

End Function

Function cptGetUserFullName()
'used to add user's name to PowerPoint title slide
Dim objAllNames As Object, objIndName As Object

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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

Function cptGetVersion(strModule As String) As String
  Dim vbComponent As Object, strVersion As String
  If Not cptModuleExists(strModule) Then
    cptGetVersion = "<uninstalled>"
  Else
    Set vbComponent = ThisProject.VBProject.VBComponents(strModule)
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = cptRegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
    End If
    cptGetVersion = strVersion
  End If
  
End Function

Function cptGetVersions() As String
'requires reference: Microsoft Scripting Runtime
Dim vbComponent As Object, strVersion As String

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
'strings
Dim strAbout As String
'longs
'integers
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  strAbout = strAbout & "All rights reserved." & vbCrLf & "Copyright " & Chr(169) & " " & Year(Now()) & ", ClearPlan Consulting, LLC"
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
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b6" 'NAS > EWR
  'cptAbout_frm.lblScoreBoard.Caption = "t0 : b7" 'EWR > SAV
  cptAbout_frm.lblScoreBoard.Caption = "t0 : b8" 'EWR > SAV
  
  cptAbout_frm.Caption = "The ClearPlan Toolbar - " & cptGetVersion("cptAbout_frm")
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  'although simply running cptSetReferences should fix it
  Dim oRef As Object 'Reference
  Dim lngFile As Long
  Dim strFile As String
  Dim strRef As String
  Dim lngRefs As Long
  Dim lngRef As Long
  
  lngFile = FreeFile
  strFile = Environ("tmp") & "\cpt-references.csv"
  Open strFile For Output As #lngFile
  
  Print #lngFile, "NAME,DESCRIPTION,FULL_PATH,GUID,MAJOR,MINOR,BUILT_IN,IS_BROKEN,TYPE,"
  lngRefs = ThisProject.VBProject.References.Count
  lngRef = 0
  For Each oRef In ThisProject.VBProject.References
    lngRef = lngRef + 1
    Debug.Print lngRef & "/" & lngRefs & " " & String(25, "=")
    Debug.Print oRef.Name
    Debug.Print "-- " & oRef.Description
    Debug.Print "-- " & oRef.FullPath
    Debug.Print "-- " & oRef.Guid & " | " & oRef.Major & " | " & oRef.Minor
    Debug.Print "-- BuiltIn: " & oRef.BuiltIn
    Debug.Print "-- IsBroken: " & oRef.IsBroken
    Debug.Print "-- Type: " & oRef.Type
    strRef = Join(Array(oRef.Name, oRef.Description, oRef.FullPath, oRef.Guid, oRef.Major, oRef.Minor, oRef.BuiltIn, oRef.IsBroken, oRef.Type), ",")
    Print #lngFile, strRef & ","
  Next oRef
  Close #lngFile
  
  Shell "notepad.exe """ & strFile & """", vbNormalFocus
  
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'the calling subroutine should catch the Not cptInternetIsConnected function before calling this

  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
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
  strDir = cptRegEx(Environ("PATH"), "C:\\[^;]*Office[0-9]{1,}\\")
  If Len(strDir) = 0 Then GoTo windows_common
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
windows_common:
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
  Dim strDefaultView As String
  Dim strFilter As String
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
  Dim blnFilter As Boolean
  'variants
  'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
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
  
  strDefaultView = cptGetSetting("ResetAll", "DefaultView")
  If Len(strDefaultView) > 0 And cptViewExists(strDefaultView) Then
    ActiveWindow.TopPane.Activate
    ViewApply strDefaultView
  End If
  
  If lngSettings > 0 Then
    'parse and apply
    If lngSettings >= 128 Then 'outline symbols
      OptionsViewEx displayoutlinesymbols:=True
      lngSettings = lngSettings - 128
    End If
    If lngSettings >= 64 Then 'display name indent
      OptionsViewEx DisplayNameIndent:=True
      lngSettings = lngSettings - 64
    End If
    If lngSettings >= 32 Then 'clear filter
      blnFilter = True
      FilterClear
      lngSettings = lngSettings - 32
    Else
      blnFilter = False
    End If
    If lngSettings >= 16 Then 'sort by ID
      Sort "ID", , , , , , False, True
      lngSettings = lngSettings - 16
    End If
    If lngSettings >= 8 Then 'expand all tasks
      OptionsViewEx DisplaySummaryTasks:=True
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
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      If ActiveProject.Subprojects.Count > 0 Then
        OptionsViewEx DisplaySummaryTasks:=True
        If Not blnFilter Then
          strFilter = ActiveProject.CurrentFilter
        End If
        FilterClear
        SelectAll
        OutlineShowAllTasks
        If Len(strFilter) > 0 Then FilterApply strFilter
      End If
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      lngSettings = lngSettings - 8
    Else 'expand to specific level
      OptionsViewEx DisplaySummaryTasks:=True
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
      If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
      OutlineShowTasks pjTaskOutlineShowLevelMax
      For lngLevel = 20 To lngOutlineLevel Step -1
        OutlineShowTasks lngLevel
      Next lngLevel
    End If
    If lngSettings >= 4 Then 'show summaries
      OptionsViewEx DisplaySummaryTasks:=True
      lngSettings = lngSettings - 4
    Else
      OptionsViewEx DisplaySummaryTasks:=False
    End If
    If lngSettings >= 2 Then 'clear group
      GroupClear
      lngSettings = lngSettings - 2
    End If
    If lngSettings >= 1 Then 'hide inactive
      If Edition = pjEditionProfessional Then SetAutoFilter "Active", pjAutoFilterFlagYes
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
  Dim oView As MSProject.View
  Dim rstSettings As Object 'ADODB.Recordset
  'strings
  Dim strViewList As String
  Dim strDefaultView As String
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
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
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
  
  'populate cboViews
  cptResetAll_frm.cboViews.Clear
  For Each oView In ActiveProject.Views
    If oView.Type = pjTaskItem Then
      If oView.Screen = 1 Or oView.Screen = 14 Then
        strViewList = strViewList & oView.Name & ","
'        Debug.Print oView.Type & vbTab & Choose(oView.Type + 1, "pjTaskItem", "pjResourceItem", "pjOtherItem") & vbTab & oView.Name
'        Debug.Print oView.Screen & vbTab & Choose(oView.Screen, "pjGantt", "pjNetworkDiagram", "pjRelationshipDiagram", "pjTaskForm", "pjTaskSheet", "pjResourceForm", "pjResourceSheet", "pjResourceGraph", "pjRSVDoNotUse", "pjTaskDetailsForm", "pjTaskNameForm", "pjResourceNameForm", "pjCalendar", "pjTaskUsage", "pjResourceUsage", "pjTimeline", "pjResourceScheduling")
      End If
    End If
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  Next oView
  Set oView = Nothing
  strViewList = Left(strViewList, Len(strViewList) - 1)
  Dim vViewList As Variant
  vViewList = Split(strViewList, ",")
  cptQuickSort vViewList, 0, UBound(vViewList)
  cptResetAll_frm.cboViews.List = Split("<None>," & Join(vViewList, ","), ",")
  
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
    strDefaultView = cptGetSetting("ResetAll", "DefaultView")
    If Len(strDefaultView) > 0 Then
      If strDefaultView <> "<None>" Then
        If Not cptViewExists(strDefaultView) Then
          MsgBox "Your default view '" & strDefaultView & "' does not exist.", vbExclamation + vbOKOnly, "Saved View Not Found"
        Else
          cptResetAll_frm.cboViews.Value = strDefaultView
        End If
      Else
        cptResetAll_frm.cboViews.Value = "<None>"
      End If
    Else
      cptSaveSetting "ResetAll", "DefaultView", "<None>"
      cptResetAll_frm.cboViews.Value = "<None>"
    End If
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
      If Edition = pjEditionProfessional Then
        .chkActiveOnly.Enabled = True
        If lngSettings >= 1 Then
          .chkActiveOnly = True
        End If
      ElseIf Edition = pjEditionStandard Then
        .chkActiveOnly = False
        .chkActiveOnly.Enabled = False
      End If
    End With
  End If
  
  cptResetAll_frm.Caption = "How would you like to Reset All? (" & cptGetVersion("cptResetAll_frm") & ")"
  cptResetAll_frm.Show False
  
exit_here:
  On Error Resume Next
  Set oView = Nothing
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'update references needed before downloading updates
  Application.StatusBar = "Updating VBA references..."
  DoEvents
  Call cptSetReferences

  'todo: user should still be able to check currently installed versions
  Application.StatusBar = "Confirming access to GitHub.com..."
  DoEvents
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
  Application.StatusBar = "Fetching latest versions..."
  DoEvents
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.errorcode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
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
  Application.StatusBar = "Comparing installed versions..."
  DoEvents
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
  If cptVersionStatus(rstStatus(3), rstStatus(2)) <> "ok" Then
    Application.StatusBar = "Automatically updating cptUpgrades_frm..."
    DoEvents
    Call cptUpgrade(rstStatus(1) & "/cptUpgrades_frm.frm")
    rstStatus(3) = rstStatus(2)
    rstStatus.Update
  End If
  
  'cannot auto upgrade cptCore_bas because this is the cptCore_bas module so use cptPatch_bas
  'if cptPatch_bas is updated, install it automatically and run it
  rstStatus.MoveFirst
  rstStatus.Find "Module='cptPatch_bas'", , 1
  If cptVersionStatus(rstStatus(3), rstStatus(2)) <> "ok" Then
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
  Application.StatusBar = "Preparing form..."
  DoEvents
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
    Application.StatusBar = Application.StatusBar = "Preparing form...(" & Format(lngItem / rstStatus.RecordCount, "0%") & ")"
    DoEvents
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
  cptUpgrades_frm.Caption = "Installation Status (" & cptGetVersion("cptUpgrades_frm") & ")"
  Application.StatusBar = "Ready for user input..."
  DoEvents
  cptUpgrades_frm.Show

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
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
Dim vPath As Variant

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
  'strDir = cptRegEx(Environ("PATH"), "C:\\[^;]*Office[0-9]{1,}\\")
  
  If Not cptReferenceExists("Excel") Then
    strDir = cptGetOfficeDir
    If strDir = "" Then GoTo windows_common
    ThisProject.VBProject.References.AddFromFile strDir & "\EXCEL.EXE"
  End If
  If Not cptReferenceExists("Outlook") Then
    strDir = cptGetOfficeDir
    If strDir = "" Then GoTo windows_common
    ThisProject.VBProject.References.AddFromFile strDir & "\MSOUTL.OLB"
  End If
  If Not cptReferenceExists("PowerPoint") Then
    strDir = cptGetOfficeDir
    If strDir = "" Then GoTo windows_common
    ThisProject.VBProject.References.AddFromFile strDir & "\MSPPT.OLB"
  End If
  If Not cptReferenceExists("MSProject") Then
    strDir = cptGetOfficeDir
    If strDir = "" Then GoTo windows_common
    ThisProject.VBProject.References.AddFromFile strDir & "\MSPRJ.OLB"
  End If
  If Not cptReferenceExists("Word") Then
    strDir = cptGetOfficeDir
    If strDir = "" Then GoTo windows_common
    ThisProject.VBProject.References.AddFromFile strDir & "\MSWORD.OLB"
  End If

  'Windows Common
  'odd other path: C:\Program Files\Microsoft Office\root\vfs\System
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

Function cptGetOfficeDir() As String
  Dim strDir As String
  Dim vPath As Variant
  
  strDir = ""
  For Each vPath In Split(Environ("PATH"), ";")
    If InStr(vPath, "Office") > 0 Then
      If Dir(CStr(vPath) & "EXCEL.EXE") <> vbNullString Then
        strDir = vPath
        Exit For
      End If
    End If
  Next vPath

  If Len(strDir) > 0 Then
    cptGetOfficeDir = strDir
  ElseIf Len(strDir) = 0 Then 'weird installation or Excel not installed
    cptGetOfficeDir = strDir
    MsgBox "Microsoft Office installation is not detetcted. Some features may not operate as expected." & vbCrLf & vbCrLf & "Please contact cpt@ClearPlanConsulting.com for specialized assistance.", vbCritical + vbOKOnly, "Microsoft Office Compatibility"
  End If
  
End Function

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
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  OptionsViewEx DisplaySummaryTasks:=True
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
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
  OptionsViewEx DisplaySummaryTasks:=True
  On Error Resume Next
  If ActiveProject.Subprojects.Count > 0 Then
    FilterClear
    SelectAll
  End If
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

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

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
        cptVersionStatus = "install " & cptVersionStatus
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

  ActiveWindow.TopPane.Activate
  GroupApply ActiveProject.CurrentGroup
  Exit Sub

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
  If lngUID > 0 Then Find "Unique ID", "equals", lngUID
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
  strReturned = Space(255) 'this determines the length of the returned value, not the length of the stored value
  lngSize = Len(strReturned)
  lngWorked = GetPrivateProfileString(strFeature, strSetting, "", strReturned, lngSize, strSettingsFile)
  If lngWorked Then
    cptGetSetting = Left$(strReturned, lngWorked)
  Else
    cptGetSetting = ""
  End If
End Function

Function cptViewExists(strView As String) As Boolean
  'objects
  Dim oView As MSProject.View

  On Error Resume Next
  Set oView = ActiveProject.Views(strView)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptViewExists = Not oView Is Nothing
  
exit_here:
  On Error Resume Next
  Set oView = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptViewExists", Err, Erl)
  Resume exit_here
End Function

Function cptTableExists(strTable As String) As Boolean
  'objects
  Dim oTable As MSProject.Table

  On Error Resume Next
  Set oTable = ActiveProject.TaskTables(strTable)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptTableExists = Not oTable Is Nothing
  
exit_here:
  On Error Resume Next
  Set oTable = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptTableExists", Err, Erl)
  Resume exit_here
End Function

Function cptFilterExists(strFilter As String) As Boolean
  'objects
  Dim oFilter As MSProject.Filter

  On Error Resume Next
  Set oFilter = ActiveProject.TaskFilters(strFilter)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptFilterExists = Not oFilter Is Nothing
  
exit_here:
  On Error Resume Next
  Set oFilter = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptFilterExists", Err, Erl)
  Resume exit_here
End Function

Function cptGroupExists(strGroup As String) As Boolean
  'objects
  Dim oGroup As MSProject.Group

  On Error Resume Next
  Set oGroup = ActiveProject.TaskGroups(strGroup)
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  cptGroupExists = Not oGroup Is Nothing
  
exit_here:
  On Error Resume Next
  Set oGroup = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGroupExists", Err, Erl)
  Resume exit_here
End Function

Sub cptCreateFilter(strFilter As String)
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Select Case strFilter
    Case "Marked"
      FilterEdit Name:="Marked", TaskFilter:=True, Create:=True, OverwriteExisting:=True, FieldName:="Marked", test:="equals", Value:="Yes", ShowInMenu:=True, ShowSummaryTasks:=False
      
  End Select
  
exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptCreateFilter", Err, Erl)
  Resume exit_here
End Sub

Sub cptShowSettings_frm()
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oStream As Scripting.TextStream
  Dim oFSO As Scripting.FileSystemObject
  'strings
  Dim strErrorTrapping As String
  Dim strSettingsFileNew As String
  Dim strSettingsFile As String
  Dim strProgramAcronym As String
  Dim strFeature As String
  Dim strLine As String
  'longs
  Dim lngFile As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSettings_frm.Caption = "ClearPlan Toolbar Settings (" & cptGetVersion("cptUpgrades_frm") & ")"
  
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Fields.Append "Feature", adVarChar, 100
    .Fields.Append "Setting", adVarChar, 255
    .Open
  End With
  
  strSettingsFile = cptDir & "\settings\cpt-settings.ini"
  strSettingsFileNew = cptDir & "\settings\cpt-settings-temp.ini"
  lngFile = FreeFile
  Open strSettingsFileNew For Output As #lngFile
  
  With cptSettings_frm
    .lboFeatures.Clear
    .lboSettings.Clear
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oStream = oFSO.OpenTextFile(cptDir & "\settings\cpt-settings.ini")
    Do While Not oStream.AtEndOfStream
      strLine = oStream.ReadLine
      If Left(strLine, 1) = "[" Then
        strFeature = Replace(Replace(strLine, "[", ""), "]", "")
        'todo: If [Driving Path Group] OR [Driving Path] Then skip
        'todo: If Count Then StatusBar: StatusBar.blnSummarizeSelection; StatusBar.blnShowStatusBarTaskCount
        oRecordset.AddNew Array(0), Array(strFeature)
      Else
        'todo: If [Metrics] Then remove cboLOEField > Integration.EVT
        'todo: If [Metrics] Then remove txtLOE > Integration.LOE
        'todo: If [Driving Path Group] OR [Driving Path] Then skip
        'todo: If [Integration] Then rename CWBS as WBS
        'todo: If [Integration] Then rename WPCN as WP
        oRecordset.AddNew Array(0, 1), Array(strFeature, strLine)
      End If
    Loop
    oStream.Close
    oRecordset.Sort = "Feature,Setting"
    oRecordset.MoveFirst
    Do While Not oRecordset.EOF
      If oRecordset(1) = "" Then
        .lboFeatures.AddItem oRecordset(0)
        Print #lngFile, "[" & oRecordset(0) & "]"
      Else
        Print #lngFile, oRecordset(1)
      End If
      oRecordset.MoveNext
    Loop
    Close #lngFile
    If Dir(strSettingsFile) <> vbNullString Then Kill strSettingsFile
    Name strSettingsFileNew As strSettingsFile
    If Dir(strSettingsFileNew) <> vbNullString Then Kill strSettingsFileNew
    If Dir(cptDir & "\settings\cpt-settings.adtg") <> vbNullString Then Kill cptDir & "\settings\cpt-settings.adtg"
    oRecordset.Save cptDir & "\settings\cpt-settings.adtg", adPersistADTG
    oRecordset.Close
    '.lblDir = strSettingsFile
    strProgramAcronym = cptGetProgramAcronym
    .txtProgramAcronym = strProgramAcronym
    If .lboFeatures.ListCount > 0 Then
      .lboFeatures.Value = .lboFeatures.List(0, 0)
    End If
    'get error-trapping on/off and set toggle button
    strErrorTrapping = cptGetSetting("General", "ErrorTrapping")
    If Len(strErrorTrapping) > 0 Then
      .tglErrorTrapping = strErrorTrapping = 0
    Else
      .tglErrorTrapping = False
    End If
    .Show
  End With
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing
  Set oStream = Nothing
  Set oFSO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptShowSettings_frm", Err, Erl)
  Resume exit_here
End Sub

Function cptGetProgramAcronym()
  'objects
  Dim oCustomDocumentProperty As DocumentProperty
  'strings
  Dim strMsg As String
  Dim strProgramAcronym As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  'dates
  
  On Error Resume Next
  Set oCustomDocumentProperty = ActiveProject.CustomDocumentProperties("cptProgramAcronym")
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  If oCustomDocumentProperty Is Nothing Then
    strMsg = "For some features, a unique program acronym is required to capture data (locally)." & vbCrLf & vbCrLf
    strMsg = strMsg & "This program acronym is saved in a custom document property named 'cptProgramAcronym'." & vbCrLf & vbCrLf
    strMsg = strMsg & "Please enter a program acronym for this file:"
    vResponse = InputBox(strMsg, "Program Acronym")
    If StrPtr(vResponse) = 0 Then
      MsgBox "No Program Acronym saved.", vbCritical + vbOKOnly, "Invalid Response"
      cptGetProgramAcronym = ""
    ElseIf vResponse = vbNullString Then
      MsgBox "No Program Acronym saved.", vbCritical + vbOKOnly, "Invalid Response"
      cptGetProgramAcronym = ""
    Else
      Set oCustomDocumentProperty = ActiveProject.CustomDocumentProperties.Add("cptProgramAcronym", False, msoPropertyTypeString, CStr(vResponse))
      cptGetProgramAcronym = CStr(vResponse)
      MsgBox "Program Acronym '" & CStr(vResponse) & "' saved!", vbInformation + vbOKOnly, "Success"
    End If
  Else
    cptGetProgramAcronym = ActiveProject.CustomDocumentProperties("cptProgramAcronym").Value
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetProgramAcronym", Err, Erl)
  Resume exit_here
End Function

Sub cptOpenSettingsFile()
  Shell "notepad.exe """ & cptDir & "\settings\cpt-settings.ini""", vbNormalFocus
End Sub

Function cptGetMyHeaders(strTitle As String, Optional blnRequired As Boolean = False) As String
  'objects
  'strings
  Dim strMyHeaders As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  Dim vResponse As Variant
  Dim vMyHeader As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

try_again:
  'get other fields
  strMyHeaders = cptGetSetting("Metrics", "txtMyHeaders")
  If Len(strMyHeaders) = 0 Then strMyHeaders = "CAM,WPCN,WPM,"
  If blnRequired Then
    vResponse = InputBox("At least one custom field is required." & vbCrLf & vbCrLf & "Enter a comma-separated list:", strTitle, strMyHeaders)
  Else
    vResponse = InputBox("Include other Custom Fields?" & vbCrLf & vbCrLf & "Enter a comma-separated list of Custom Field Names:", strTitle, strMyHeaders)
  End If
  
  If StrPtr(vResponse) = 0 Then 'user hit cancel
    strMyHeaders = ""
    GoTo exit_here
  ElseIf vResponse = "" Or Len(Replace(vResponse, ",", "")) = 0 Then 'user entered zero-value
    If blnRequired Then
      'nothing selected
      If MsgBox("You must select at least one custom field. Try again?", vbQuestion + vbYesNo, "Field Required") = vbYes Then
        GoTo try_again
      Else
        strMyHeaders = ""
        GoTo exit_here
      End If
    Else
      strMyHeaders = ""
      GoTo exit_here
    End If
  ElseIf Len(vResponse) > 0 Then
    strMyHeaders = CStr(vResponse)
  End If

  Application.StatusBar = "Validating custom fields..."
  DoEvents
  If Right(Trim(strMyHeaders), 1) <> "," Then strMyHeaders = Trim(strMyHeaders) & ","
  'validate strMyHeaders
  On Error Resume Next
  For Each vMyHeader In Split(strMyHeaders, ",")
    If vMyHeader = "" Then Exit For
    Debug.Print FieldNameToFieldConstant(vMyHeader)
    If Err.Number > 0 Then
      vResponse = MsgBox("Custom Field '" & vMyHeader & "' not found!" & vbCrLf & vbCrLf & "OK = skip; Cancel = try again", vbExclamation + vbOKCancel, "Invalid Field")
      If vResponse = vbCancel Then
        Err.Clear
        GoTo try_again
      Else
        Err.Clear
        strMyHeaders = Replace(strMyHeaders, vMyHeader & ",", "")
      End If
    End If
  Next vMyHeader
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  cptSaveSetting "Metrics", "txtMyHeaders", strMyHeaders

  cptGetMyHeaders = strMyHeaders

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetMyHeaders", Err, Erl)
  Resume exit_here

End Function

Function cptMasterUIDToSubUID(lngMasterUID As Long) As Long
  If ActiveProject.Subprojects.Count = 0 Then Exit Function
  cptMasterUIDToSubUID = lngMasterUID Mod 4194304
End Function

Function cptGetSubprojectUID(lngMasterUID As Long) As Long
  If ActiveProject.Subprojects.Count = 0 Then Exit Function
  cptGetSubprojectUID = Round(lngMasterUID / 4194304, 0) - 1
End Function

Function cptSubUIDToMasterUID(lngSubProjectUID As Long, lngSubUID As Long) As Long
  If ActiveProject.Subprojects.Count = 0 Then Exit Function
  cptSubUIDToMasterUID = ((lngSubProjectUID + 1) * 4194304) + lngSubUID
End Function

Function cptConvertToMasterUIDs(oTask As MSProject.Task, strReturn As String) As String
  'strReturn variable expects either "p" for predecessors or "s" for successors
  Dim oSubprojects As MSProject.Subprojects
  Dim strProject As String, strList As String, strLinkProject As String, strConvertedList As String
  Dim lngUID As Long, lngLinkUID As Long
  Dim vLink As Variant
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oSubprojects = ActiveProject.Subprojects
  If oSubprojects.Count = 0 Then GoTo exit_here
  
  strProject = oTask.Project
  If strReturn = "p" Then
    strList = oTask.UniqueIDPredecessors
  ElseIf strReturn = "s" Then
    strList = oTask.UniqueIDSuccessors
  End If
  For Each vLink In Split(strList, ",")
    If InStr(vLink, "\") > 0 Then 'handle offline and server paths
      lngUID = CLng(Mid(vLink, InStrRev(vLink, "\") + 1))   'extract source task UID
      strLinkProject = Replace(vLink, "\" & lngUID, "")     'strip source task UID
      strLinkProject = Replace(strLinkProject, ".mpp", "")  'strip file extension
      strLinkProject = Mid(strLinkProject, InStrRev(strLinkProject, "\") + 1) 'strip path
      lngLinkUID = oSubprojects(strLinkProject).InsertedProjectSummary.UniqueID + 1 'get master task UID seed
    Else
      lngUID = vLink
      lngLinkUID = oSubprojects(strProject).InsertedProjectSummary.UniqueID + 1 'get master task UID seed
    End If
    lngLinkUID = (lngLinkUID * 4194304) + lngUID            'derive master task UID
    strConvertedList = strConvertedList & lngLinkUID & ","  'build return string
  Next vLink
  strConvertedList = Left(strConvertedList, Len(strConvertedList) - 1) 'strip last comma avoiding null value
  cptConvertToMasterUIDs = strConvertedList

exit_here:
  On Error Resume Next
  Set oSubprojects = Nothing
  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptConvertToMasterUIDs", Err, Erl)
  Resume exit_here
End Function

Function cptGetShowStatusBarCountFirstRun() As Boolean
  'objects
  'strings
  Dim strShow As String
  'longs
  'integers
  'doubles
  'booleans
  Dim blnShow As Boolean
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strShow = cptGetSetting("Count", "blnShowStatusBarTaskCount")
  If Len(strShow) > 0 Then
    blnShow = CBool(strShow)
  Else
    Call cptSaveSetting("Count", "blnShowStatusBarTaskCount", "1") 'default is true
    blnShow = True
  End If

  cptGetShowStatusBarCountFirstRun = blnShow

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetShowStatusBarCountFirstRun", Err, Erl)
  Resume exit_here
End Function

Sub cptAppendColumn(strFile As String, strColumn As String, lngType As Long, Optional lngLength As Long, Optional vDefault As Variant)
  'objects
  Dim oRecordsetNew As ADODB.Recordset
  Dim oRecordset As ADODB.Recordset
  'strings
  'longs
  Dim lngField As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  Set oRecordset = CreateObject("ADODB.Recordset")
  Set oRecordsetNew = CreateObject("ADODB.Recordset")
  If InStr(strFile, cptDir) = 0 Then strFile = cptDir & strFile
  oRecordset.Open strFile, , adOpenKeyset, adLockReadOnly
  On Error Resume Next
  Debug.Print oRecordset.Fields(strColumn)
  If Err.Number = 0 Then 'field already exists
    Err.Clear
    If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
    GoTo exit_here
  End If
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  'first rebuld the existing fields
  For lngField = 0 To oRecordset.Fields.Count - 1
    With oRecordsetNew
      If oRecordset.Fields(lngField).DefinedSize > 0 Then
        .Fields.Append oRecordset.Fields(lngField).Name, oRecordset.Fields(lngField).Type, oRecordset.Fields(lngField).DefinedSize
      Else
        .Fields.Append oRecordset.Fields(lngField).Name, oRecordset.Fields(lngField).Type
      End If
    End With
  Next lngField
  'next add the new field
  If lngLength > 0 Then
    oRecordsetNew.Fields.Append strColumn, lngType, lngLength
  Else
    oRecordsetNew.Fields.Append strColumn, lngType
  End If
  oRecordsetNew.Open
  'next move the existing data over
  If Not oRecordset.EOF Then oRecordset.MoveFirst
  Do While Not oRecordset.EOF
    oRecordsetNew.AddNew
    For lngField = 0 To oRecordset.Fields.Count - 1
      oRecordsetNew.Fields(lngField) = oRecordset.Fields(lngField)
    Next lngField
    oRecordsetNew.Fields(strColumn) = vDefault
    oRecordset.MoveNext
  Loop
  oRecordset.Close
  Name strFile As Replace(strFile, ".adtg", "-backup_" & Format(Now, "yyyy-mm-dd-HH-nn-ss") & ".adtg")
  oRecordsetNew.Save strFile, adPersistADTG
  oRecordsetNew.Close
  
exit_here:
  On Error Resume Next
  Set oRecordsetNew = Nothing
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptAppendColumn", Err, Erl)
  Resume exit_here
End Sub

Sub cptGetSums(ByRef oTasks As MSProject.Tasks, lngFieldID As Long)
  'objects
  Dim oTask As MSProject.Task
  'strings
  Dim strCustomFieldName As String
  Dim strFieldName As String
  'longs
  Dim lngTask As Long
  Dim lngTasks As Long
  Dim lngDuration As Long
  'integers
  'doubles
  Dim dblWork As Double
  Dim dblCost As Double
  Dim dblNumber As Double
  'booleans
  'variants
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  dblCost = 0#
  lngDuration = 0
  dblNumber = 0#
  dblWork = 0#
  lngTasks = oTasks.Count
  
  strFieldName = FieldConstantToFieldName(lngFieldID)
  
  If Len(cptRegEx(strFieldName, "Cost|Duration|Number|Work")) > 0 Then
    lngTask = 0
    For Each oTask In oTasks
      If oTask Is Nothing Then GoTo next_task
      If Not oTask.Active Then GoTo next_task
      If oTask.GetField(lngFieldID) = "#ERROR" Then GoTo next_task
      'do not ignore external tasks
      'do not ignore summary tasks
      Select Case strFieldName
        Case "Actual Cost"
          dblCost = dblCost + oTask.ActualCost
        Case "Cost"
          dblCost = dblCost + oTask.Cost
        Case "Remaining Cost"
          dblCost = dblCost + oTask.RemainingCost
        Case "Cost1"
          dblCost = dblCost + oTask.Cost1
        Case "Cost2"
          dblCost = dblCost + oTask.Cost2
        Case "Cost3"
          dblCost = dblCost + oTask.Cost3
        Case "Cost4"
          dblCost = dblCost + oTask.Cost4
        Case "Cost5"
          dblCost = dblCost + oTask.Cost5
        Case "Cost6"
          dblCost = dblCost + oTask.Cost6
        Case "Cost7"
          dblCost = dblCost + oTask.Cost7
        Case "Cost8"
          dblCost = dblCost + oTask.Cost8
        Case "Cost9"
          dblCost = dblCost + oTask.Cost9
        Case "Cost10"
          dblCost = dblCost + oTask.Cost10
          
        Case "Baseline Cost"
          dblCost = dblCost + Val(oTask.BaselineCost)
        Case "Baseline1 Cost"
          dblCost = dblCost + Val(oTask.Baseline1Cost)
        Case "Baseline2 Cost"
          dblCost = dblCost + Val(oTask.Baseline2Cost)
        Case "Baseline3 Cost"
          dblCost = dblCost + Val(oTask.Baseline3Cost)
        Case "Baseline4 Cost"
          dblCost = dblCost + Val(oTask.Baseline4Cost)
        Case "Baseline5 Cost"
          dblCost = dblCost + Val(oTask.Baseline5Cost)
        Case "Baseline6 Cost"
          dblCost = dblCost + Val(oTask.Baseline6Cost)
        Case "Baseline7 Cost"
          dblCost = dblCost + Val(oTask.Baseline7Cost)
        Case "Baseline8 Cost"
          dblCost = dblCost + Val(oTask.Baseline8Cost)
        Case "Baseline9 Cost"
          dblCost = dblCost + Val(oTask.Baseline9Cost)
        Case "Baseline10 Cost"
          dblCost = dblCost + Val(oTask.Baseline10Cost)
          
        Case "Actual Duration"
          lngDuration = lngDuration + oTask.ActualDuration
        Case "Duration"
          lngDuration = lngDuration + oTask.Duration
        Case "Remaining Duration"
          lngDuration = lngDuration + oTask.RemainingDuration
        Case "Duration1"
          lngDuration = lngDuration + oTask.Duration1
        Case "Duration2"
          lngDuration = lngDuration + oTask.Duration2
        Case "Duration3"
          lngDuration = lngDuration + oTask.Duration3
        Case "Duration4"
          lngDuration = lngDuration + oTask.Duration4
        Case "Duration5"
          lngDuration = lngDuration + oTask.Duration5
        Case "Duration6"
          lngDuration = lngDuration + oTask.Duration6
        Case "Duration7"
          lngDuration = lngDuration + oTask.Duration7
        Case "Duration8"
          lngDuration = lngDuration + oTask.Duration8
        Case "Duration9"
          lngDuration = lngDuration + oTask.Duration9
        Case "Duration10"
          lngDuration = lngDuration + oTask.Duration10
          
        Case "Baseline Duration"
          lngDuration = lngDuration + oTask.BaselineDuration
        Case "Baseline1 Duration"
          lngDuration = lngDuration + oTask.Baseline1Duration
        Case "Baseline2 Duration"
          lngDuration = lngDuration + oTask.Baseline2Duration
        Case "Baseline3 Duration"
          lngDuration = lngDuration + oTask.Baseline3Duration
        Case "Baseline4 Duration"
          lngDuration = lngDuration + oTask.Baseline4Duration
        Case "Baseline5 Duration"
          lngDuration = lngDuration + oTask.Baseline5Duration
        Case "Baseline6 Duration"
          lngDuration = lngDuration + oTask.Baseline6Duration
        Case "Baseline7 Duration"
          lngDuration = lngDuration + oTask.Baseline7Duration
        Case "Baseline8 Duration"
          lngDuration = lngDuration + oTask.Baseline8Duration
        Case "Baseline9 Duration"
          lngDuration = lngDuration + oTask.Baseline9Duration
        Case "Baseline10 Duration"
          lngDuration = lngDuration + oTask.Baseline10Duration
                    
        Case "Number"
          dblNumber = dblNumber + oTask.Number
        Case "Number1"
          dblNumber = dblNumber + oTask.Number1
        Case "Number2"
          dblNumber = dblNumber + oTask.Number2
        Case "Number3"
          dblNumber = dblNumber + oTask.Number3
        Case "Number4"
          dblNumber = dblNumber + oTask.Number4
        Case "Number5"
          dblNumber = dblNumber + oTask.Number5
        Case "Number6"
          dblNumber = dblNumber + oTask.Number6
        Case "Number7"
          dblNumber = dblNumber + oTask.Number7
        Case "Number8"
          dblNumber = dblNumber + oTask.Number8
        Case "Number9"
          dblNumber = dblNumber + oTask.Number9
        Case "Number10"
          dblNumber = dblNumber + oTask.Number10
        Case "Number11"
          dblNumber = dblNumber + oTask.Number11
        Case "Number12"
          dblNumber = dblNumber + oTask.Number12
        Case "Number13"
          dblNumber = dblNumber + oTask.Number13
        Case "Number14"
          dblNumber = dblNumber + oTask.Number14
        Case "Number15"
          dblNumber = dblNumber + oTask.Number15
        Case "Number16"
          dblNumber = dblNumber + oTask.Number16
        Case "Number17"
          dblNumber = dblNumber + oTask.Number17
        Case "Number18"
          dblNumber = dblNumber + oTask.Number18
        Case "Number19"
          dblNumber = dblNumber + oTask.Number19
        Case "Number20"
          dblNumber = dblNumber + oTask.Number20
          
        Case "Actual Work"
          dblWork = dblWork + oTask.ActualWork
        Case "Work"
          dblWork = dblWork + oTask.Work
        Case "Remaining Work"
          dblWork = dblWork + oTask.RemainingWork
        Case "Baseline Work"
          dblWork = dblWork + oTask.BaselineWork
        Case "Baseline1 Work"
          dblWork = dblWork + oTask.Baseline1Work
        Case "Baseline2 Work"
          dblWork = dblWork + oTask.Baseline2Work
        Case "Baseline3 Work"
          dblWork = dblWork + oTask.Baseline3Work
        Case "Baseline4 Work"
          dblWork = dblWork + oTask.Baseline4Work
        Case "Baseline5 Work"
          dblWork = dblWork + oTask.Baseline5Work
        Case "Baseline6 Work"
          dblWork = dblWork + oTask.Baseline6Work
        Case "Baseline7 Work"
          dblWork = dblWork + oTask.Baseline7Work
        Case "Baseline8 Work"
          dblWork = dblWork + oTask.Baseline8Work
        Case "Baseline9 Work"
          dblWork = dblWork + oTask.Baseline9Work
        Case "Baseline10 Work"
          dblWork = dblWork + oTask.Baseline10Work
      End Select

next_task:
      lngTask = lngTask + 1
      Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | Calculating...(" & Format(lngTask / lngTasks, "0%") & ")"
    Next oTask
  End If
  
  strCustomFieldName = CustomFieldGetName(lngFieldID)
  If Len(strCustomFieldName) > 0 Then strFieldName = strCustomFieldName & " (" & strFieldName & ")"

  If dblCost > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(dblCost, "$#,###,##0.00")
  ElseIf lngDuration > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(lngDuration / 480, "#,###,##0d")
  ElseIf dblNumber > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(dblNumber, "#,###,##0.00")
  ElseIf dblWork > 0 Then
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected" & " | " & strFieldName & ": " & Format(dblWork / 60, "#,###,##0.00h")
  Else
    Application.StatusBar = Format(lngTasks, "#,##0") & " task" & IIf(lngTasks = 1, "", "s") & " selected"
  End If
  
exit_here:
  On Error Resume Next
  Set oTask = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetSums", Err, Erl)
  Resume exit_here
End Sub

Function cptGetCustomFields(strFieldTypes As String, strDataTypes As String, strInclude As String, Optional blnIncludeEnterprise As Boolean = False) As Variant
  'strFieldTypes  := comma-separated list of any of "p,t,r" [project,task,resource]
  'strDataTypes   := comma-separated list of any of "Cost,Date,Duration,Flag,Finish,Number,Start,Text,Outline Code"
  'strInclude     := comma-separated list of any of "c,fn,cfn" [constant,fieldname,customfieldname]
  'blnIncludeEnterprise := self-explanatory
  'objects
  Dim oFieldTypes As Scripting.Dictionary
  'strings
  Dim strFieldName As String
  Dim strCustomFieldName As String
  Dim strResult As String
  Dim strRow As String
  'longs
  Dim lngInclude As Long
  Dim lngFieldCount As Long
  Dim lngFieldType As Long
  Dim lngField As Long
  Dim lngConstant As Long
  Dim lngResultCount As Long
  'integers
  'doubles
  'booleans
  'variants
  Dim vInclude As Variant
  Dim vFieldType As Variant
  Dim vField As Variant
  Dim vRow() As Variant
  Dim vResult() As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Set oFieldTypes = CreateObject("Scripting.Dictionary")
  oFieldTypes.Add "p", pjProject
  oFieldTypes.Add "t", pjTask
  oFieldTypes.Add "r", pjResource
  
  vInclude = Split(strInclude, ",")
  ReDim vRow(0 To UBound(vInclude))
  For Each vFieldType In Split(strFieldTypes, ",")
    lngFieldType = oFieldTypes(vFieldType)
    For Each vField In Split(strDataTypes, ",")
      lngFieldCount = 10
      If vField = "Flag" Then lngFieldCount = 20
      If vField = "Number" Then lngFieldCount = 20
      If vField = "Text" Then lngFieldCount = 30
      For lngField = 1 To lngFieldCount
        lngConstant = FieldNameToFieldConstant(vField & lngField, lngFieldType)
        'Debug.Print lngConstant; FieldConstantToFieldName(lngConstant)
        For lngInclude = 0 To UBound(vInclude)
          If vInclude(lngInclude) = "c" Then
            vRow(lngInclude) = lngConstant
          ElseIf vInclude(lngInclude) = "fn" Then
            vRow(lngInclude) = FieldConstantToFieldName(lngConstant)
          ElseIf vInclude(lngInclude) = "cfn" Then
            If Len(CustomFieldGetName(lngConstant)) > 0 Then
              vRow(lngInclude) = CustomFieldGetName(lngConstant)
            Else
              vRow(lngInclude) = FieldConstantToFieldName(lngConstant)
            End If
          End If
        Next lngInclude
        strResult = strResult & Join(vRow, ",") & vbCrLf
        lngResultCount = lngResultCount + 1
      Next lngField
    Next vField
  Next vFieldType
  
  ReDim vResult(0 To UBound(Split(strResult, vbCrLf)), 0 To UBound(vInclude))
  For lngField = 0 To UBound(Split(strResult, vbCrLf)) - 1
    For lngInclude = 0 To UBound(vInclude)
      vResult(lngField, lngInclude) = Split(Split(strResult, vbCrLf)(lngField), ",")(lngInclude)
    Next lngInclude
  Next lngField
  
  cptGetCustomFields = vResult
  
exit_here:
  On Error Resume Next
  Set oFieldTypes = Nothing

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptGetCustomFields", Err, Erl)
  Resume exit_here
  
End Function

Sub cptGetValidMap()
  If ValidMap Then
    'do nothing
  End If
End Sub

Function cptErrorTrapping() As Boolean
  'objects
  'strings
  Dim strErrorTrapping As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  On Error GoTo err_here 'some users experiencing error on recursive call

  strErrorTrapping = cptGetSetting("General", "ErrorTrapping")
  If Len(strErrorTrapping) > 0 Then
    cptErrorTrapping = CBool(strErrorTrapping)
  Else
    cptSaveSetting "General", "ErrorTrapping", "1"
    cptErrorTrapping = True
  End If

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptCore_bas", "cptErrorTrapping", Err, Erl)
  Resume exit_here
End Function

