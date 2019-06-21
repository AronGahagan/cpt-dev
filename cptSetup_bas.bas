Attribute VB_Name = "cptSetup_bas"
'<cpt_version>v1.3.11</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Public Const strGitHub = "https://raw.githubusercontent.com/AronGahagan/cpt-dev/master/"
'Public Const strGitHub = "https://raw.githubusercontent.com/ClearPlan/cpt/master/"
#If Win64 And VBA7 Then
  Private Declare PtrSafe Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As LongPtr, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As LongPtr) As LongPtr

#Else
  Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As Long) As Long
#End If

Sub cptSetup()
'setup only needs to be run once
'objects
Dim Project As Object
Dim vbComponent As Object 'vbComponent
Dim arrCode As Object
Dim cmThisProject As Object 'CodeModule
Dim cmCptThisProject As Object 'CodeModule
Dim oStream As Object 'ADODB.Stream
Dim xmlHttpDoc As Object
Dim xmlNode As Object
Dim xmlDoc As Object
Dim arrCore As Object
'strings
Dim strMsg As String
Dim strError As String
Dim strCptFileName As String
Dim strVersion As String
Dim strFileName As String
Dim strModule As String
Dim strURL As String
'longs
Dim lngLine As Long
Dim lngEvent As Long
'Dim lngFile As Long
'integers
'booleans
Dim blnImportModule As Boolean
Dim blnExists As Boolean
'variants
Dim vEvent As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '<issue61> ensure proper installation
  If Instr(ThisProject.FullName, "Global.MPT") = 0 Then
    MsgBox "The VBA module 'cptSetup_bas' must be installed in your Global.MPT, not in this project file.", vbCritical + vbOKOnly, "Faulty Installation"
    GoTo exit_here
  End If '</issue61>

  'prompt user for setup instructions
  strMsg = "NOTE: This procedure should only be run once." & vbCrLf & vbCrLf
  strMsg = strMsg & "Before proceeding:" & vbCrLf
  strMsg = strMsg & "1. Setup your Global.MPT: File > Options > Trust Center > Trust Center Settings..." & vbCrLf
  strMsg = strMsg & vbTab & "a. Macro Settings > Enable all macros" & vbCrLf
  strMsg = strMsg & vbTab & "b. Legacy Formats > Allow loading files with legacy or non-default file formats" & vbCrLf
  strMsg = strMsg & "2. Completely exit, then re-open, MS Project (this makes the settings above 'stick')" & vbCrLf
  strMsg = strMsg & "Have you completed the above steps?" & vbCrLf & vbCrLf
  strMsg = strMsg & "(Yes = Proceed; No = Cancel and Close)"
  If MsgBox(strMsg, vbQuestion + vbYesNo, "Before you proceed...") = vbNo Then GoTo exit_here

  'capture list of files to download
  Set arrCore = CreateObject("System.Collections.SortedList")
  Application.StatusBar = "Identifying latest core CPT modules..."
  'get CurrentVersions.xml
  'get file list in cpt\Core
  Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
  xmlDoc.async = False
  xmlDoc.validateOnParse = False
  xmlDoc.SetProperty "SelectionLanguage", "XPath"
  xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.microsoft.com/ado/2007/08/dataservices' xmlns:m='http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'"
  strURL = strGitHub & "CurrentVersions.xml"
  If Not xmlDoc.Load(strURL) Then
    If xmlDoc.parseError.ErrorCode = -2146697210 Or -xmlDoc.parseError.ErrorCode = -2146697208 Then '</issue35>
      MsgBox "Please check your internet connection.", vbCritical + vbOKOnly, "Can't Connect"
    Else
      strMsg = "We're having trouble downloading modules:" & vbCrLf & vbCrLf  '</issue35>
      strMsg = strMsg & xmlDoc.parseError.ErrorCode & ": " & xmlDoc.parseError.reason & vbCrLf & vbCrLf '</issue35>
      strMsg = strMsg & "If the ClearPlan ribbon doesn't show up, please contact cpt@ClearPlanConsulting.com for assistance." '</issue35>
      MsgBox strMsg, vbExclamation + vbOKOnly, "XML Error" '</issue35>
    End If
    'GoTo exit_here '</issue35> redirected
    GoTo this_project '</issue35> redirected
  Else
    'download cpt/core/*.* to user's tmp directory
    arrCore.Clear
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      If xmlNode.SelectSingleNode("Directory").Text = "Core" Then
        Application.StatusBar = "Fetching " & xmlNode.SelectSingleNode("Name").Text & "..."
        arrCore.Add xmlNode.SelectSingleNode("FileName").Text, xmlNode.SelectSingleNode("Type").Text
        'get ThisProject status for later
        If xmlNode.SelectSingleNode("FileName").Text = "ThisProject.cls" Then
          strVersion = xmlNode.SelectSingleNode("Version").Text
        End If
        'build the url of the download
        strURL = strGitHub
        If Len(xmlNode.SelectSingleNode("Directory").Text) > 0 Then
          strURL = strURL & xmlNode.SelectSingleNode("Directory").Text & "/"
        End If
        strFileName = xmlNode.SelectSingleNode("FileName").Text
        strURL = strURL & strFileName
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
          GoTo next_xmlNode
        End If

        'remove if exists
        strModule = Left(strFileName, InStr(strFileName, ".") - 1)
        If strModule = "ThisProject" Then GoTo next_xmlNode
        blnExists = False
        For Each vbComponent In ThisProject.VBProject.VBComponents
          If vbComponent.Name = strModule Then
            Application.StatusBar = "Removing obsolete version of " & vbComponent.Name
            'Debug.Print Application.StatusBar
            '<issue19> revised
            vbComponent.Name = vbComponent.Name & "_" & Format(Now, "hhnnss")
            DoEvents
            ThisProject.VBProject.VBComponents.remove vbComponent 'ThisProject.VBProject.VBComponents(CStr(vbComponent.Name))
            DoEvents '</issue19>
            Exit For
          End If
        Next vbComponent

        'import the module - skip ThisProject which needs special handling
        If strModule <> "ThisProject" Then
          Application.StatusBar = "Importing " & strFileName & "..."
          'Debug.Print Application.StatusBar
          ThisProject.VBProject.VBComponents.import cptDir & "\" & strFileName
          '<issue19> added
          DoEvents '</issue19>

          '<issue24>remove the whitespace added by VBE import/export
          With ThisProject.VBProject.VBComponents(strModule).CodeModule
            For lngLine = .CountOfDeclarationLines To 1 Step -1
              If Len(.Lines(lngLine, 1)) = 0 Then .DeleteLines lngLine, 1
            Next lngLine
          End With '</issue24>

        End If

      End If
next_xmlNode:
    Next xmlNode
  End If

  Application.StatusBar = "CPT Modules imported."

this_project:

  '<issue35>
  'update user's ThisProject - if it downloaded correctly, or was copied in correctly
  strFileName = cptDir & "\ThisProject.cls"
  If Dir(strFileName) <> vbNullString Then 'it was downloaded, import it
    'rename the file and import it
    strCptFileName = Replace(strFileName, "ThisProject", "cptThisProject_cls")
    If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName
    Name strFileName As strCptFileName
    'import the module
    If cptModuleExists("cptThisProject_cls") Then
      ThisProject.VBProject.VBComponents.remove ThisProject.VBProject.VBComponents("cptThisProject_cls")
      DoEvents
    End If
    Set cmCptThisProject = ThisProject.VBProject.VBComponents.import(strCptFileName).CodeModule
  ElseIf cptModuleExists("cptThisProject_cls") Then 'it was copied in
    Set cmCptThisProject = ThisProject.VBProject.VBComponents("cptThisProject_cls").CodeModule
  Else 'ThisProject not imported or downloaded, so skip
    GoTo skip_import
  End If '</issue35>

  'avoid messy overwrites of ThisProject
  Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
  '<issue10> revised
  'If cmThisProject.Find("<cpt_version>", 1, 1, cmThisProject.CountOfLines, 1000, True, True) = True Then
  If cmThisProject.Find("<cpt_version>", 1, 1, cmThisProject.CountOfLines, 1000, False, True) = True Then
  '</issue10>
    strMsg = "Your 'ThisProject' module has already been updated to work with the ClearPlan toolbar." & vbCrLf & vbCrLf
    strMsg = strMsg & "Would you like to reset it? This will only overwrite CodeModule lines appended with '</cpt>'" & vbCrLf & vbCrLf
    strMsg = strMsg & "(Please note: if you have made modifications to your ThisProject module, you may need to review them if you proceed.)"
    If MsgBox(strMsg, vbExclamation + vbYesNo, "Danger, Will Robinson!") = vbYes Then
      For lngLine = cmThisProject.CountOfLines To 1 Step -1
        If InStr(cmThisProject.Lines(lngLine, 1), "'</cpt>") > 0 Then
          cmThisProject.DeleteLines lngLine
        End If
      Next lngLine
    Else
      GoTo skip_import
    End If
  End If

  'grab the imported code
  '<issue35>
  If Len(strVersion) = 0 Then 'grab the version
    strVersion = cptRegEx(ThisProject.VBProject.VBComponents("cptThisProject_cls").CodeModule.Lines(1, 1000), "<cpt_version>.*</cpt_version>")
    strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
  End If '</issue35>
  Set arrCode = CreateObject("System.Collections.SortedList")
  With cmCptThisProject
    For Each vEvent In Array("Project_Activate", "Project_Open")
      arrCode.Add CStr(vEvent), .Lines(.ProcStartLine(CStr(vEvent), 0) + 2, .ProcCountLines(CStr(vEvent), 0) - 3) '0 = vbext_pk_Proc
    Next
  End With
  ThisProject.VBProject.VBComponents.remove ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name)
  '<issue19> added
  DoEvents '</issue19>

  'add the events, or insert new text
  'three cases: empty or not empty (code exists or not)
  For Each vEvent In Array("Project_Activate", "Project_Open")

    'if event exists then insert code else create new event handler
    With cmThisProject
      If .CountOfLines > .CountOfDeclarationLines Then 'complications
        If .Find("Sub " & CStr(vEvent), 1, 1, .CountOfLines, 1000) = True Then
        'find its line number
          lngEvent = .ProcBodyLine(CStr(vEvent), 0) '= vbext_pk_Proc
          'import them if they *as a group* don't exist
          If .Find(arrCode(CStr(vEvent)), .ProcStartLine(CStr(vEvent), 0), 1, .ProcCountLines(CStr(vEvent), 0), 1000) = False Then 'vbext_pk_Proc
            .InsertLines lngEvent + 1, arrCode(CStr(vEvent))
          Else
            'Debug.Print CStr(vEvent) & " code exists."
          End If
        Else 'create it
          'create it, returning its line number
          lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
          'insert cpt code after line number
          .InsertLines lngEvent + 1, arrCode(CStr(vEvent))
        End If
      Else 'easy
        'create it, returning its line number
        lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
        'insert cpt code after line number
        .InsertLines lngEvent + 1, arrCode(CStr(vEvent))
      End If 'lines exist
    End With 'thisproject.codemodule

    'add version if not exists
    With cmThisProject
      If .Find("<cpt_version>", 1, 1, .CountOfLines, 1000) = False Then
        .InsertLines 1, "'<cpt_version>" & strVersion & "</cpt_version>" & vbCrLf
      End If
    End With
  Next vEvent

  'leave no trace
  'If Dir(strCptFileName, vbNormal) <> vbNullString Then Kill strCptFileName

skip_import:

  If Len(strError) > 0 Then
    strError = "The following modules did not download correctly:" & vbCrLf & strError & vbCrLf & vbCrLf & "Please contact cpt@ClearPlanConsulting.com for assistance."
    MsgBox strError, vbCritical + vbOKOnly, "Unknown Error"
    'Debug.Print strError
  End If

  'reset the toolbar
  strMsg = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
  strMsg = strMsg + "<mso:customUI "
  strMsg = strMsg + "xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >"
  strMsg = strMsg + vbCrLf & "<mso:ribbon startFromScratch=""false"" >"
  strMsg = strMsg + vbCrLf & "<mso:tabs>"
  strMsg = strMsg + cptBuildRibbonTab()
  strMsg = strMsg + vbCrLf & "</mso:tabs>"
  strMsg = strMsg + vbCrLf & "</mso:ribbon>"
  strMsg = strMsg + vbCrLf & "</mso:customUI>"
  ActiveProject.SetCustomUI (strMsg)

exit_here:
  On Error Resume Next
  Set Project = Nothing
  '<issue19> added
  Application.StatusBar = "" '</issue19>
  '<issue23> added
  Application.ScreenUpdating = True '</issue23>
  Set vbComponent = Nothing
  Set arrCode = Nothing
  Set cmThisProject = Nothing
  Set cmCptThisProject = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Set xmlNode = Nothing
  Set xmlDoc = Nothing
  Set arrCore = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptSetup", err, Erl)
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

  'backbone
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gBackbone"" label=""Backbone"" visible=""true"">"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBackbone"" imageMso=""OutlineShowDetail"" label=""Import Mil-Std-881D Appendix B"" onAction=""DrivingPaths"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bOutlineCode"" imageMso=""OutlineShowDetail"" label=""Create Outline Code"" onAction=""DrivingPaths"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDD"" imageMso=""OutlineShowDetail"" label=""Data Dictionary"" onAction=""DrivingPaths"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group> "
  
  
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
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Remove"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUninstall"" label=""Uninstall ClearPlan Toolbar"" imageMso=""TasksUnlink"" onAction=""cptUninstall"" visible=""true"" />" 'supertip=" & Chr(34) & strSuperTip & Chr(34) & "
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAbout"" onAction=""ShowcptAbout_frm""  size=""large"" visible=""true""  label=""About"" imageMso=""Info"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  ribbonXML = ribbonXML + vbCrLf & "</mso:tab>"

  'Debug.Print "<mso:customUI ""xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >" & ribbonXML
  cptBuildRibbonTab = ribbonXML

End Function

Sub cptHandleErr(strModule As String, strProcedure As String, objErr As ErrObject, Optional lngErl As Long)
'common error handling prompt
Dim strMsg As String

    strMsg = "Uh oh!" & vbCrLf & vbCrLf & "Please contact cpt@ClearPlanConsulting.com for assistance if needed." & vbCrLf & vbCrLf
    strMsg = strMsg & "Error " & err.Number & ": " & err.Description & vbCrLf
    strMsg = strMsg & "Source: " & strModule & "." & strProcedure
    If lngErl > 0 Then
      strMsg = strMsg & vbCrLf & "Line: " & lngErl
    End If
    MsgBox strMsg, vbExclamation + vbOKOnly, "Unknown Error"

End Sub

Function cptIncrement(ByRef lgCleanUp As Long) As Long
  lgCleanUp = lgCleanUp + 1
  cptIncrement = lgCleanUp
End Function

Public Function cptInternetIsConnected() As Boolean

    cptInternetIsConnected = InternetGetConnectedStateEx(0, "", 254, 0)

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

Function cptDir() As String
Dim strPath As String

  'confirm existence of cpt settings and backup modules file

  'strPath = ThisProject.FullName
  'strPath = Left(strPath, InStrRev(strPath, "MS Project\") - 1 + Len("MS Project\"))

  strPath = Environ("USERPROFILE")
  strPath = strPath & "\cpt-backup"
  If Dir(strPath, vbDirectory) = vbNullString Then
    MkDir strPath
  End If
  If Dir(strPath & "\settings", vbDirectory) = vbNullString Then
    MkDir strPath & "\settings"
  End If
  If Dir(strPath & "\modules", vbDirectory) = vbNullString Then
    MkDir strPath & "\modules"
  End If
  cptDir = strPath

End Function

Function cptModuleExists(strModule As String)
'objects
Dim vbComponent As Object
'booleans
Dim blnExists As Boolean
'strings
Dim strError As String

  On Error Resume Next
  'Set vbComponent = ThisProject.VBProject.VBComponents(strModule)
  cptModuleExists = Not ThisProject.VBProject.VBComponents(strModule) Is Nothing
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  GoTo exit_here

  For Each vbComponent In ThisProject.VBProject.VBComponents
    If UCase(vbComponent.Name) = UCase(strModule) Then
      blnExists = True
      Exit For
    End If
  Next vbComponent

  cptModuleExists = blnExists

exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call cptHandleErr("cptSetup_bas", "cptModuleExists", err, Erl)
  Resume exit_here

End Function

Sub cptUninstall()
'objects
Dim vEvent As Object
Dim Project As Object
Dim vbComponent As Object
Dim cmThisProject As Object
'strings
Dim strMsg As String
'longs
Dim lngLine As Long
'integers
'doubles
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If MsgBox("Are you sure?", vbCritical + vbYesNo, "Uninstall CPT") = vbNo Then GoTo exit_here

  strMsg = "1. Please delete the module 'cptSetup_bas' manually after this process completes." & vbCrLf & vbCrLf
  strMsg = strMsg & "2. If you have made modifications to the ThisProject codemodule, you may need to review it." & vbCrLf & vbCrLf
  strMsg = strMsg & "Alternatively, if you would like to reinstall, re-run cptSetup() and then install updates."
  If MsgBox(strMsg, vbInformation + vbOKCancel, "Thank You!") = vbCancel Then GoTo exit_here

  'remove cpt-related lines from ThisProject
  Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
  With cmThisProject
    'delete the version
    For lngLine = .CountOfDeclarationLines To 1 Step -1
      If InStr(.Lines(lngLine, 1), "<cpt_version>") > 0 Then
        .DeleteLines lngLine, 1
        DoEvents
      End If
    Next lngLine
    For lngLine = .CountOfLines To 1 Step -1
      'comment out existing lines to avoid immediate errors
      If InStr(.Lines(lngLine, 1), "Sub") > 0 Then
        'do nothing
      ElseIf InStr(.Lines(lngLine, 1), "</cpt") > 0 Then
        If .ProcOfLine(lngLine, 1) = "Project_Activate" Then
          'holding next line in case we decide to comment out instead of delete
          '.ReplaceLine lngLine, "'" & .Lines(lngLine, 1)
          .DeleteLines lngLine, 1
          DoEvents
        ElseIf .ProcOfLine(lngLine, 1) = "Project_Open" Then
          .DeleteLines lngLine, 1
          DoEvents
        End If
      End If
    Next lngLine
  End With

  'reset the toolbar
  ActiveProject.SetCustomUI "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui""><mso:ribbon></mso:ribbon></mso:customUI>"

  'remove all cpt modules
  For Each vbComponent In ThisProject.VBProject.VBComponents
    If Left(vbComponent.Name, 3) = "cpt" And vbComponent.Name <> "cptSetup_bas" Then
      If vbComponent.Name = "cptAdmin_bas" Then GoTo next_component
      Application.StatusBar = "Purging module " & vbComponent.Name & "..."
      If Dir(cptDir & "\modules\", vbDirectory) = vbNullString Then MkDir cptDir & "\modules"
      vbComponent.Export cptDir & "\modules\" & vbComponent.Name
      ThisProject.VBProject.VBComponents.remove vbComponent
    End If
next_component:
  Next vbComponent

  MsgBox "Thank you for using the ClearPlan Toolbar.", vbInformation + vbOKOnly, "Uninstall Complete"

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set vEvent = Nothing
  Set Project = Nothing
  Set vbComponent = Nothing
  Set cmThisProject = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptUninstall", err, Erl)
  Resume exit_here
End Sub
