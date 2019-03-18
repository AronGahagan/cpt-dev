Attribute VB_Name = "cptCore_bas"
'<cpt_version>v1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private oMSPEvents As cptEvents_cls
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As Long) As Long

Sub StartEvents()
  Set oMSPEvents = New cptEvents_cls
End Sub

Public Function InternetIsConnected() As Boolean
 
    InternetIsConnected = InternetGetConnectedStateEx(0, "", 254, 0)
 
End Function

Function GetUserFullName()
'used to add user's name to PowerPoint title slide
Dim objAllNames As Object, objIndName As Object

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  On Error Resume Next
  Set objAllNames = GetObject("Winmgmts:").instancesof("win32_networkloginprofile")
  For Each objIndName In objAllNames
    If Len(objIndName.FullName) > 0 Then
      GetUserFullName = objIndName.FullName
      Exit For
    End If
  Next

exit_here:
  On Error Resume Next
  Set objAllNames = Nothing
  Set objIndName = Nothing
  Exit Function
err_here:
  Call HandleErr("basCommon", "GetUserFullName", err)
  Resume exit_here

End Function

Function GetVersions() As String
'requires reference: Microsoft Scripting Runtime
Dim vbComponent As Object, strVersion As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = RegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      'check for updates - future capability
      If InternetIsConnected Then
        'go get it
        'extract the version
        'clear obsolete version
        'insert current version
        GetVersions = GetVersions & vbComponent.Name & ": " & strVersion & vbCrLf
        'if updates available then
        'If MsgBox(GetVersions, vbYesNo + vbInformation, "Updates Available") = vbYes Then Call ApplyUpdates
      Else
        'if updates available, prompt to apply
      End If
    End If
next_component:
  Next vbComponent

exit_here:
  On Error Resume Next
  
  Exit Function
err_here:
  Call HandleErr("basCommon", "GetVersions", err)
  Resume exit_here

End Function

Sub CheckVersions()
Dim strMsg As String

  cptLogo_frm.lblVersions.Caption = "Currently Installed:" & vbCrLf & vbCrLf & GetVersions
  cptLogo_frm.Show

End Sub

Sub ApplyUpdates()
  If MsgBox("Updates are available. Apply now?", vbQuestion + vbYesNo, "Please Confirm") = vbYes Then
    'do the things
  End If
End Sub

Function ModuleExists(strModule)
Dim vbComponent As Object
Dim blnExists As Boolean

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnExists = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    If UCase(vbComponent.Name) = UCase(strModule) Then
      blnExists = True
      Exit For
    End If
  Next vbComponent
  
  ModuleExists = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call HandleErr("basCommon", "ModuleExists", err)
  Resume exit_here
  
End Function

Function ReferenceExists(strReference) As Boolean
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
  
  ReferenceExists = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call HandleErr("basCommon", "ReferenceExists", err)
  Resume exit_here
End Function

Sub GetReferences()
'prints the current uesr's selected references
'this would be used to troubleshoot with users real-time
Dim Ref As Object

  For Each Ref In ThisProject.VBProject.References
    Debug.Print Ref.Name & " (" & Ref.Description & ")" & Ref.FullPath
  Next Ref

End Sub

Sub GetEnviron()
'list the environment variables and their associated values
Dim lgIndex As Long
  
  For lgIndex = 1 To 200
    Debug.Print lgIndex & ": " & Environ(lgIndex)
  Next
  
End Sub

Sub CheckLogo()
Dim strFileName As String

  If ModuleExists("cptLogo_frm") Then
    strFileName = Environ("tmp") & "\ClearPlanLogo.jpg"
    If Dir(strFileName) = vbNullString Then
      SavePicture cptLogo_frm.Image1.Picture, strFileName
    End If
  End If

End Sub

Function CheckReference(strReference As String) As Boolean
'this routine will be called ahead of any subroutine requiring a reference
'returns boolean and subroutine only proceeds if true
Dim blnExists As Boolean

  On Error GoTo err_here

  CheckReference = True

  Select Case strReference
    'CommonProgramFiles
    Case "Office"
      If Not ReferenceExists("Office") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\OFFICE16\MSO.DLL"
      End If
    Case "VBIDE"
      If Not ReferenceExists("VBIDE") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
      End If
    Case "VBA"
      If Not ReferenceExists("VBA") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
      End If
    Case "ADODB"
      If Not ReferenceExists("ADODB") Then
        ThisProject.VBProject.References.AddFromFile Environ("CommonProgramFiles") & "\System\ado\msado15.dll"
      End If
    
    'Office Applications
    Case "Excel"
      If Not ReferenceExists("Excel") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\EXCEL.EXE"
      End If
    Case "Outlook"
      If Not ReferenceExists("Outlook") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSOUTL.OLB"
      End If
    Case "PowerPoint"
      If Not ReferenceExists("PowerPoint") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSPPT.OLB"
      End If
    Case "MSProject"
      If Not ReferenceExists("MSProject") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSPRJ.OLB"
      End If
    Case "Word"
      If Not ReferenceExists("Word") Then
        ThisProject.VBProject.References.AddFromFile Application.Path & "\MSWORD.OLB (Word)"
      End If
    
    'Windows Common
    Case "MSForms"
      If Not ReferenceExists("MSForms") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\FM20.DLL"
      End If
    Case "Scripting"
      If Not ReferenceExists("Scripting") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\scrrun.dll"
      End If
    Case "stdole"
      If Not ReferenceExists("stdole") Then
        ThisProject.VBProject.References.AddFromFile Environ("windir") & "\SysWOW64\stdole2.tlb"
      End If
    Case "mscorlib"
      If Not ReferenceExists("") Then
        ThisProject.VBProject.References.AddFromFile Environ("winddir") & "\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
      End If
    Case Else
      CheckReference = False
    
  End Select
  
  If Not CheckReference Then
    MsgBox "Missing Reference: " & strReference, vbExclamation + vbOKOnly, "CP Tool Bar"
  End If
  
exit_here:
  On Error Resume Next

  Exit Function
err_here:
  CheckReference = False
  Resume exit_here
  
End Function

Sub ResetAll()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  FilterClear
  GroupClear
  OptionsViewEx displaynameindent:=True, displaysummarytasks:=True, displayoutlinesymbols:=True
  SelectAll 'needed for master/sub projects
  OutlineShowAllTasks
  Sort "ID"
  SelectBeginning

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("basCommon", "ResetAll", err)
  Resume exit_here
  
End Sub

Sub ShowCptUpgrades_frm()
'objects
Dim vbComponent As Object
Dim arrCurrent As Object, arrInstalled As Object
Dim xmlDoc As Object, xmlNode As Object, xmlHttpDoc As Object, FindRecord As Object
Dim oStream As Object
'long
Dim lgItem As Long, lgCol As Long
'strings
Dim strURL As String, strMsg As String, strVersion As String, strFileName As String
'booleans
Dim blnUpdatesAreAvailable As Boolean, blnLoaded As Boolean
'variants
Dim vCol As Variant

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'user should still be able to check currently installed versions
  If Not InternetIsConnected Then
    MsgBox "You must be connected to the internet to perform updates.", vbInformation + vbOKOnly, "No Connection"
    GoTo exit_here
  End If

  If Not CheckReference("VBA") Or Not CheckReference("VBIDE") Then
    GoTo exit_here
  End If
  
  'get current versions
  Set arrCurrent = CreateObject("System.Collections.SortedList")
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  '/=======CHANGE THIS TO PUBLIC SITE BEFORE RELEASE=======\
  strURL = "https://raw.githubusercontent.com/AronGahagan/test/master/CurrentVersions.xml"
  '\=======CHANGE THIS TO PUBLIC SITE BEFORE RELEASE=======/
  If Not xmlDoc.Load(strURL) Then
    MsgBox xmlDoc.parseError.ErrorCode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
    GoTo exit_here
  Else
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      arrCurrent.Add xmlNode.SelectSingleNode("Name").Text, xmlNode.SelectSingleNode("Version").Text
    Next
  End If

  'get installed versions
  Set arrInstalled = CreateObject("System.Collections.SortedList")
  blnUpdatesAreAvailable = False
  For Each vbComponent In ThisProject.VBProject.VBComponents
    'is the vbComponent one of ours?
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = RegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      arrInstalled.Add vbComponent.Name, strVersion
    End If
next_component:
  Next vbComponent
  
  'populate the listbox header
  lgItem = 0
  cptUpgrades_frm.lboHeader.AddItem
  For Each vCol In Array("Module", "Current", "Installed", "Status", "Type")
    cptUpgrades_frm.lboHeader.List(0, lgItem) = vCol
    lgItem = lgItem + 1
  Next vCol
  cptUpgrades_frm.lboHeader.Height = 16
  
  'populate the listbox
  cptUpgrades_frm.lboModules.Clear
  For lgItem = 0 To arrCurrent.count - 1
    cptUpgrades_frm.lboModules.AddItem
    cptUpgrades_frm.lboModules.List(lgItem, 0) = arrCurrent.getKey(lgItem) 'module name
    cptUpgrades_frm.lboModules.List(lgItem, 1) = arrCurrent.getValueList()(lgItem) 'current version
    If arrInstalled.contains(arrCurrent.getKey(lgItem)) Then 'installed version
      cptUpgrades_frm.lboModules.List(lgItem, 2) = arrInstalled.getValueList()(arrInstalled.indexofkey(arrCurrent.getKey(lgItem)))
    Else
      cptUpgrades_frm.lboModules.List(lgItem, 2) = "<not installed>"
    End If
    
    Select Case cptUpgrades_frm.lboModules.List(lgItem, 2)
      Case Is = cptUpgrades_frm.lboModules.List(lgItem, 1)
        cptUpgrades_frm.lboModules.List(lgItem, 3) = "<installed>"
      Case Is = "<not installed>"
        cptUpgrades_frm.lboModules.List(lgItem, 3) = "<install>"
      Case Is <> cptUpgrades_frm.lboModules.List(lgItem, 1)
        cptUpgrades_frm.lboModules.List(lgItem, 3) = "<install update>"
    End Select
    Set FindRecord = xmlDoc.SelectSingleNode("//Name[text()='" + cptUpgrades_frm.lboModules.List(lgItem, 0) + "']").ParentNode.SelectSingleNode("Type")
    cptUpgrades_frm.lboModules.List(lgItem, 4) = FindRecord.Text
      
  Next lgItem
    
  cptUpgrades_frm.Show False
  
exit_here:
  On Error Resume Next
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
  Call HandleErr("basCommon", "UpdatesAreAvailable", err)
  Resume exit_here

End Sub

Sub SetReferences()
'this is a one-time shot to set all references currently required by the cp toolbar
Dim strDir As String, Ref As Object
  
  On Error Resume Next
  
  'CommonProgramFiles
  strDir = Environ("CommonProgramFiles")
  If Not ReferenceExists("Office") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\OFFICE16\MSO.DLL"
  End If
  If Not ReferenceExists("VBIDE") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
  End If
  If Not ReferenceExists("VBA") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\Microsoft Shared\VBA\VBA7.1\VBE7.DLL"
  End If
  If Not ReferenceExists("ADODB") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\System\ado\msado15.dll"
  End If
  
  'office applications
  strDir = Application.Path 'OR RegEx(environ("PATH"),"C\:.*Microsoft Office[A-z0-9\\]*;")
  If Not ReferenceExists("Excel") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\EXCEL.EXE"
  End If
  If Not ReferenceExists("Outlook") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSOUTL.OLB"
  End If
  If Not ReferenceExists("PowerPoint") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSPPT.OLB"
  End If
  If Not ReferenceExists("MSProject") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSPRJ.OLB"
  End If
  If Not ReferenceExists("Word") Then
    ThisProject.VBProject.References.AddFromFile strDir & "\MSWORD.OLB"
  End If
  
  'Windows Common
  If Not ReferenceExists("MSForms") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\SysWOW64\FM20.DLL"
  End If
  If Not ReferenceExists("Scripting") Then
    ThisProject.VBProject.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll"
  End If
  If Not ReferenceExists("stdole") Then
    ThisProject.VBProject.References.AddFromFile "C:\Windows\SysWOW64\stdole2.tlb"
  End If
  If Not ReferenceExists("mscorlib") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\mscorlib.tlb"
  End If
  
End Sub

Sub HandleErr(strModule As String, strProcedure As String, err As ErrObject)
'common error handling prompt
Dim strMsg As String

    strMsg = "Module: " & strModule & vbCrLf
    strMsg = strMsg & "Procedure: " & strProcedure & vbCrLf
    strMsg = strMsg & err.Number & ": " & err.Description
    MsgBox strMsg, vbExclamation + vbOKOnly, "Error"
    
End Sub

Function UpdatesAreAvailable() As Boolean
End Function

Function RemoveIllegalCharacters(ByVal strText As String) As String
'written by Ryan Beard (RyanBeard@ClearPlanConsulting.com)
    Const cstrIllegals As String = "\,/,:,*,?,"",<,>,|"
    
    Dim lngCounter As Long
    Dim astrChars() As String
    
    astrChars() = Split(cstrIllegals, ",")
    
    For lngCounter = LBound(astrChars()) To UBound(astrChars())
        strText = Replace(strText, astrChars(lngCounter), vbNullString)
    Next lngCounter
    
    RemoveIllegalCharacters = strText

End Function

Sub WrapItUp()
'objects
Dim Tasks As Object
'strings
'longs
Dim lgLevel As Long
'booleans
'variants
'dates

  On Error Resume Next
  Set Tasks = ActiveProject.Tasks
  If Tasks Is Nothing Then GoTo no_tasks
  If Tasks.count = 0 Then GoTo no_tasks

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Application.ScreenUpdating = False
  FilterClear
  'GroupClear
  OptionsViewEx displaysummarytasks:=True
  SelectAll
  OutlineShowAllTasks
  OutlineShowTasks pjTaskOutlineShowLevelMax
  For lgLevel = pjTaskOutlineShowLevelMax To pjTaskOutlineShowLevel1 Step -1
    OutlineShowTasks lgLevel
  Next lgLevel
  Application.ScreenUpdating = True

exit_here:
  On Error Resume Next
  Set Tasks = Nothing
  Exit Sub
  
no_tasks:
  MsgBox "This project has no tasks to collapse.", vbExclamation + vbOKOnly, "WrapItUp"
  GoTo exit_here

err_here:
  Call HandleErr("basCommon", "WrapItUp", err)
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
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & Increment(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:AutoFilterProject"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:FilterClear"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:SplitViewCreate"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & Increment(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetAll"" label=""Reset All"" imageMso=""FilterClear"" onAction=""ResetAll"" visible=""true"" size=""large"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bWrapItUp"" label=""WrapItUp"" imageMso=""CollapseAll"" onAction=""WrapItUp"" visible=""true"" size=""large"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  
  'task counters
  If ModuleExists("cptCountTasks_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCount"" label=""Count"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountAll"" label=""All Tasks"" imageMso=""NumberInsert"" onAction=""CountTasksAll"" visible=""true""/>" 'SelectWholeLayout
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountVisible"" label=""Visible Tasks"" imageMso=""NumberInsert"" onAction=""CountTasksVisible"" visible=""true""/>" 'SelectRows
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountSelected"" label=""Selected Tasks"" imageMso=""NumberInsert"" onAction=""CountTasksSelected"" visible=""true""/>" 'SelectTaskCell
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'text tools
  If ModuleExists("cptTextTools_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gTextTools"" label=""Text"" visible=""true"" >"
    If ModuleExists("cptDynamicFilter_bas") And ModuleExists("cptDynamicFilter_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDynamicFilter"" label=""Dynamic Filter"" imageMso=""FilterBySelection"" onAction=""ShowcptDynamicFilter_frm"" visible=""true"" size=""large"" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTextTools"" label=""Tools"" imageMso=""TextBoxInsert"" visible=""true"" size=""large"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""BulkPrepend"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""BulkAppend"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""MyReplace"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""Enumerate"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""TrimTaskNames"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process"" imageMso=""DuplicateSelectedSlides"" onAction=""ReplicateProcess"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""FindDuplicateTaskNames"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & Increment(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdvancedTextTools"" label=""Advanced (WIP)"" imageMso=""AdvancedFilterDialog"" onAction=""ShowcptTextTools_frm"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'critical path tools
  If ModuleExists("basCriticalPathTools") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCPA"" label=""Trace"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTrace"" label=""Tools"" imageMso=""TaskEntryView"" visible=""true"" size=""large"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" label=""Driving Path"" imageMso=""TaskDrivers"" onAction=""DrivingPaths"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bExport"" label=""Export to PowerPoint"" imageMso=""SlideNew"" onAction=""ExportCriticalPathSelected"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & Increment(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReset"" label=""Reset View"" imageMso=""FilterClear"" onAction=""ResetView"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'status
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gStatus"" label=""Status"" visible=""true"" >"
    If ModuleExists("cptStatusSheet_bas") And ModuleExists("cptStatusSheet_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bStatusSheet"" label=""Create Status Sheet"" imageMso=""DateAndTimeInsertOneNote"" onAction=""ShowcptStatusSheet_frm"" visible=""true""/>"
    End If
    If ModuleExists("cptSmartDur_frm") And ModuleExists("cptSmartDur_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSmartDuration"" label=""Smart Duration"" imageMso=""CalendarToolSelectDate"" onAction=""SmartDuration"" visible=""true""/>"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  
  'resource allocation
  If ModuleExists("cptResourceDemand_bas") And ModuleExists("cptResourceDemand_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gResourceDemand"" label=""Resource Demand"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResourceDemandExcel"" label=""Export to Excel"" imageMso=""Chart3DColumnChart"" onAction=""ShowFrmExportResourceDemand"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'cobra export
  If ModuleExists("cptIMSCobraExport_bas") And ModuleExists("cptIMSCobraExport_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gIntegration"" label=""Integration"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCOBRA"" label=""COBRA Export Tool"" imageMso=""Export"" onAction=""Export_IMS"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'bcr tool
  If ModuleExists("basBaselineChange") Then
  
  End If
  
  'about
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gHelp"" label=""About"" visible=""true"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAbout"" onAction=""CheckVersions""  size=""large"" visible=""true"" "
  ribbonXML = ribbonXML + "label=""About"" imageMso=""Info"" "
  ribbonXML = ribbonXML & "/>"
  If InternetIsConnected Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUpdate"" label=""Check Status"" imageMso=""AcceptTask"" onAction=""ShowcptUpgrades_frm"" size = ""large"" visible=""true"" />" 'supertip=" & Chr(34) & strSuperTip & Chr(34) & "
  End If
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  
  If Environ("UserName") = "arong" Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gAdmin"" label=""Admin"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdmin"" label=""Backup VBA Modules"" imageMso=""SaveProject"" onAction=""ShowFrmBackupVBA"" visible=""true"" size=""large"" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  ribbonXML = ribbonXML + vbCrLf & "</mso:tab>"

  cptBuildRibbonTab = ribbonXML

End Function

Function Increment(ByRef lgCleanUp As Long) As Long
  lgCleanUp = lgCleanUp + 1
  Increment = lgCleanUp
End Function

Function VersionStatus(strInstalled As String, strCurrent As String) As String
'objects
'strings
Dim strAction As String
'longs
Dim lngVersion As Long
Dim lngInstalled As Long
Dim lngCurrent As Long
'integers
'booleans
'variants
Dim aCurrent As Variant
Dim aInstalled As Variant
Dim vVersion As Variant
'dates

  'for now, assumes semantic version control - see https://semver.org/
  'useing yy.dd.mm would be easier, would show age of the release
  'and does it really matter if we 'have the lates' patch if we know we have *the latest*?

  'todo: indicate install, (major upgrade, minor upgrade, patch, downgrade) available
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'clean the versions - include all three levels
  For Each vVersion In Array(strInstalled, strCurrent)
    'following line doesn't remove non-numeric characters
    vVersion = RegEx(CStr(vVersion), "([0-9].*.?){1,3}")
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
  
  If aCurrent(0) > aInstalled(0) Then
    VersionStatus = "major upgrade"
    GoTo exit_here
  ElseIf aCurrent(1) > aInstalled(1) Then
    VersionStatus = "minor upgrade"
    GoTo exit_here
  ElseIf aCurrent(2) > aInstalled(2) Then
    VersionStatus = "upgrade patch"
    GoTo exit_here
  ElseIf aCurrent(0) = aInstalled(0) And aCurrent(1) = aInstalled(1) And aCurrent(2) = aInstalled(2) Then
    VersionStatus = "up to date"
    GoTo exit_here
  End If
  
  If aCurrent(0) < aInstalled(0) Then
    VersionStatus = "major downgrade"
    GoTo exit_here
  ElseIf aCurrent(1) < aInstalled(1) Then
    VersionStatus = "minor downgrade"
    GoTo exit_here
  ElseIf aCurrent(2) < aInstalled(2) Then
    VersionStatus = "downgrade patch"
    GoTo exit_here
  Else
    VersionStatus = "error"
  End If
    
exit_here:
  On Error Resume Next
  
  Exit Function
err_here:
  Call HandleErr("cptCore_bas", "VersionCompare", err)
  Resume exit_here
  
End Function

Function RegEx(strText As String, strRegEx As String) As String
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
    RegEx = strMatch
    
exit_here:
    On Error Resume Next
    Set RE = Nothing
    Set REMatches = Nothing
    Exit Function
err_here:
  If err.Number = 5 Then
    RegEx = ""
    err.Clear
  End If
  Resume exit_here
End Function
