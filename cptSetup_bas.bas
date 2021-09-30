Attribute VB_Name = "cptSetup_bas"
'<cpt_version>v1.5.5</cpt_version>
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
Dim rstCode As Object 'ADODB.Recordset
Dim cmThisProject As Object 'CodeModule
Dim cmCptThisProject As Object 'CodeModule
Dim oStream As Object 'ADODB.Stream
Dim xmlHttpDoc As Object
Dim xmlNode As Object
Dim xmlDoc As Object
Dim rstCore As Object 'ADODB.Recordset
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
'integers
'booleans
Dim blnImportModule As Boolean
Dim blnExists As Boolean
'variants
Dim vEvent As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  '<issue61> ensure proper installation
  If InStr(ThisProject.FullName, "Global.MPT") = 0 Then
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
  'why?
  On Error Resume Next
  Set rstCore = CreateObject("ADODB.Recordset")
  rstCore.Fields.Append "FileName", 200, 255 '200=adVarChar
  rstCore.Fields.Append "FileType", 3 '3=adInteger
  rstCore.Open
  
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
    GoTo this_project
  Else
    'download cpt/core/*.* to user's tmp directory
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      If xmlNode.SelectSingleNode("Directory").Text = "Core" Then
        Application.StatusBar = "Fetching " & xmlNode.SelectSingleNode("Name").Text & "..."
        rstCore.AddNew Array(0, 1), Array(xmlNode.SelectSingleNode("FileName").Text, xmlNode.SelectSingleNode("Type").Text)
        rstCore.Update
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
            ThisProject.VBProject.VBComponents.Remove vbComponent 'ThisProject.VBProject.VBComponents(CStr(vbComponent.Name))
            DoEvents '</issue19>
            Exit For
          End If
        Next vbComponent

        'import the module - skip ThisProject which needs special handling
        If strModule <> "ThisProject" Then
          Application.StatusBar = "Importing " & strFileName & "..."
          'Debug.Print Application.StatusBar
          ThisProject.VBProject.VBComponents.Import cptDir & "\" & strFileName
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
      ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents("cptThisProject_cls")
      DoEvents
    End If
    Set cmCptThisProject = ThisProject.VBProject.VBComponents.Import(strCptFileName).CodeModule
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
  Set rstCode = CreateObject("ADODB.Recordset")
  rstCode.Fields.Append "EVENT", 200, 255 '200=adVarChar
  rstCode.Fields.Append "LINES", 201, 1 '201=adLongVarChar;1=adParamInput
  rstCode.Open
  With cmCptThisProject
    For Each vEvent In Array("Project_Activate", "Project_Open")
      rstCode.AddNew Array(0, 1), Array(CStr(vEvent), .Lines(.ProcStartLine(CStr(vEvent), 0) + 2, .ProcCountLines(CStr(vEvent), 0) - 3)) '0 = vbext_pk_Proc
      rstCode.Update
    Next vEvent
  End With
  ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name)
  If cptModuleExists("ThisProject1") Then
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents("ThisProject1")
  End If
  If cptModuleExists("cptThisProject_cls") Then
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents("cptThisProject_cls")
  End If
  '<issue19> added
  DoEvents '</issue19>

  'add the events, or insert new text
  'three cases: empty or not empty (code exists or not)
  For Each vEvent In Array("Project_Activate", "Project_Open")

    'if event exists then insert code else create new event handler
    With cmThisProject
      If .CountOfLines > .CountOfDeclarationLines Then 'complications
        rstCode.MoveFirst
        rstCode.Find "EVENT='" & vEvent & "'"
        If .Find("Sub " & CStr(vEvent), 1, 1, .CountOfLines, 1000) = True Then
          'find its line number
          lngEvent = .ProcBodyLine(CStr(vEvent), 0) '= vbext_pk_Proc
          'import them if they *as a group* don't exist
          If .Find(rstCode(1), .ProcStartLine(CStr(vEvent), 0), 1, .ProcCountLines(CStr(vEvent), 0), 1000) = False Then  'vbext_pk_Proc
            .InsertLines lngEvent + 1, rstCode(1)
          End If
          rstCode.Filter = ""
        Else 'create it
          'create it, returning its line number
          lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
          'insert cpt code after line number
          .InsertLines lngEvent + 1, rstCode(1)
        End If
      Else 'easy
        'create it, returning its line number
        lngEvent = .CreateEventProc(Replace(CStr(vEvent), "Project_", ""), "Project")
        'insert cpt code after line number
        .InsertLines lngEvent + 1, rstCode(1)
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
  Set rstCode = Nothing
  Set cmThisProject = Nothing
  Set cmCptThisProject = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Set xmlNode = Nothing
  Set xmlDoc = Nothing
  Set rstCore = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptSetup_bas", "cptSetup", Err, Erl)
  Resume exit_here
End Sub

Public Function cptBuildRibbonTab()
Dim ribbonXML As String
Dim lngCleanUp As Long

  'build ClearPlan Ribbon Tab XML
  ribbonXML = ribbonXML + vbCrLf & "<mso:tab id=""tCommon"" label=""ClearPlan"" >" 'insertBeforeQ=""mso:TabTask"">"

  'common tools
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""custom_view"" label=""View"" visible=""true"">"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:OutlineSymbolsShow"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:SummaryTasks"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:NameIndent"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:AutoFilterProject"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:FilterClear"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:SplitViewCreate"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbResetAll"" size=""large"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetAll"" label=""Reset All"" imageMso=""FilterClear"" onAction=""cptResetAll"" screentip=""Reset All"" supertip=""Reset the current view based on your saved settings.""/>"  'in basCore_bas 'visible=""true""
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mResetAll"">"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Settings"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetAllSettings"" label=""Settings"" imageMso=""AdministrationHome"" onAction=""cptShowResetAll_frm"" visible=""true"" screentip=""Reset All Settings"" supertip=""Tweak your saved 'Reset All' settings.""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
  
  ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbWrapItUp"" >" 'size=""large""
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bWrapItUp"" label=""WrapItUp"" imageMso=""CollapseAll"" onAction=""cptWrapItUp"" supertip=""Collapse summary tasks starting from lowest level up to level 2. Defaults to your saved setting from Reset All or 2 if you don't have a saved setting yet."" />"   'in basCore_bas;visible=""true"" size=""large""
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mWrapItUp"">"
  'ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""WrapItUp to Level:"" />"
  'ribbonXML = ribbonXML + vbCrLf & "<mso:control idQ=""mso:OutlineShowAllTasks"" visible=""true""/>"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevelAll"" label=""All Subtasks"" imageMso=""OutlineTasksShowAll"" onAction=""cptWrapItUpAll"" visible=""true"" screentip=""Show All Subtasks"" supertip=""Show All Subtasks""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel1"" label=""Level 1"" imageMso=""_1"" onAction=""cptWrapItUp1"" visible=""true"" screentip=""WrapItUp to Level 1"" supertip=""WrapItUp to Level 1""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel2"" label=""Level 2"" imageMso=""_2"" onAction=""cptWrapItUp2"" visible=""true"" screentip=""WrapItUp to Level 2"" supertip=""WrapItUp to Level 2""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel3"" label=""Level 3"" imageMso=""_3"" onAction=""cptWrapItUp3"" visible=""true"" screentip=""WrapItUp to Level 3"" supertip=""WrapItUp to Level 3""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel4"" label=""Level 4"" imageMso=""_4"" onAction=""cptWrapItUp4"" visible=""true"" screentip=""WrapItUp to Level 4"" supertip=""WrapItUp to Level 4""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel5"" label=""Level 5"" imageMso=""_5"" onAction=""cptWrapItUp5"" visible=""true"" screentip=""WrapItUp to Level 5"" supertip=""WrapItUp to Level 5""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel6"" label=""Level 6"" imageMso=""_6"" onAction=""cptWrapItUp6"" visible=""true"" screentip=""WrapItUp to Level 6"" supertip=""WrapItUp to Level 6""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel7"" label=""Level 7"" imageMso=""_7"" onAction=""cptWrapItUp7"" visible=""true"" screentip=""WrapItUp to Level 7"" supertip=""WrapItUp to Level 7""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel8"" label=""Level 8"" imageMso=""_8"" onAction=""cptWrapItUp8"" visible=""true"" screentip=""WrapItUp to Level 8"" supertip=""WrapItUp to Level 8""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bLevel9"" label=""Level 9"" imageMso=""_9"" onAction=""cptWrapItUp9"" visible=""true"" screentip=""WrapItUp to Level 9"" supertip=""WrapItUp to Level 9""/>"  'in basCore_bas
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
  
  'ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bGroupReapply""  label=""ReGroup"" imageMso=""RefreshWebView"" onAction=""cptGroupReapply"" visible=""true"" supertip=""Reapply Group"" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFilterReapply"" label=""ReFilter"" imageMso=""RefreshWebView"" onAction=""cptFilterReapply"" visible=""true"" supertip=""Reapply Filter"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'task counters
  If cptModuleExists("cptCountTasks_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCount"" label=""Count"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountSelected"" label=""Selected"" imageMso=""NumberInsert"" onAction=""cptCountTasksSelected"" visible=""true""/>" 'SelectTaskCell
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountVisible"" label=""Visible"" imageMso=""NumberInsert"" onAction=""cptCountTasksVisible"" visible=""true""/>" 'SelectRows
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCountAll"" label=""All"" imageMso=""NumberInsert"" onAction=""cptCountTasksAll"" visible=""true""/>" 'SelectWholeLayout
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'text tools
  If cptModuleExists("cptText_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gTextTools"" label=""Text"" visible=""true"" >"
    If cptModuleExists("cptFilterByClipboard_bas") And cptModuleExists("cptFilterByClipboard_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bClipboard"" label=""Filter by Clipboard"" imageMso=""PasteOption"" onAction=""cptShowFilterByClipboard_frm"" visible=""true"" supertip=""Paste a list of Unique IDs or IDs from text, email, Excel, etc. to filter the current schedule. Accepts strings delimited by commas, tabs, or semicolons--or even tables, as long as the Unique ID (or ID) is the left-most column."" />"
    End If
    If cptModuleExists("cptDynamicFilter_bas") And cptModuleExists("cptDynamicFilter_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDynamicFilter"" label=""Dynamic Filter"" imageMso=""FilterBySelection"" onAction=""cptShowDynamicFilter_frm"" visible=""true"" supertip=""Find-as-you-type. Example: Keep Selected task, filter the rest of the schedule for a predecessor, add a link, CTRL+BACKSPACE to return to task you kept. Then do the next one. Filter or Highlight filter, include summaries in the search, or include related summaries. Oh, and you can Undo. Pure awesomeness."" />"
    End If
    If cptModuleExists("cptText_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbText"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAdvancedTextTools"" label=""Advanced"" imageMso=""AdvancedFilterDialog"" onAction=""cptShowText_frm"" supertip=""Bulk prefix, append, real find/replace, enumeration, everyting you could want. Oh, and Undo. Go ahead, give it a try."" />" 'visible=""true""
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mText"">"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Utilities"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""cptBulkPrepend"" visible=""true"" supertip=""Just what it sounds like."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""cptBulkAppend"" visible=""true"" supertip=""Just what it sounds like."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""cptMyReplace"" visible=""true"" supertip=""Find/Replace only on selected tasks, in the selected field."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""cptEnumerate"" visible=""true"" supertip=""Select a group of tasks, and then enumerate them."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""cptTrimTaskNames"" visible=""true"" supertip=""For the 'Type A' folks out there, this trims leading and trailing spaces (and multiple spaces) in your task names (e.g., after pasting them in from Excel--cool, right?)."" />"
      'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process (WIP)"" imageMso=""DuplicateSelectedSlides"" onAction=""cptReplicateProcess"" visible=""true"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""cptFindDuplicateTaskNames"" visible=""true"" supertip=""Clearly worded tasks represent well-defined tasks and are important for estimating and providing status. Click to find duplicate task names and create a report in Excel. Remember: Noun and Verb!"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetRowHeight"" label=""Reset Row Height"" imageMso=""RowHeight"" onAction=""cptResetRowHeight"" visible=""true"" supertip=""Another one for our fellow 'Type A' folks out there--reset all row heights after they get all jacked up. Give it a go; you'll like it."" />"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
    Else
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTextTools"" label=""Tools"" imageMso=""TextBoxInsert"" visible=""true"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPrepend"" label=""Bulk Prepend"" imageMso=""RightArrow2"" onAction=""cptBulkPrepend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAppend"" label=""Bulk Append"" imageMso=""LeftArrow2"" onAction=""cptBulkAppend"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bMyReplace"" label=""MyReplace"" imageMso=""ReplaceDialog"" onAction=""cptMyReplace"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bEnumerate"" label=""Enumerate"" imageMso=""NumberingRestart"" onAction=""cptEnumerate"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrimText"" label=""Trim Task Names"" imageMso=""TextEffectsClear"" onAction=""cptTrimTaskNames"" visible=""true""/>"
      'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bReplicateProcess"" label=""Replicate A Process"" imageMso=""DuplicateSelectedSlides"" onAction=""cptReplicateProcess"" visible=""true"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFindDuplicates"" label=""Find Duplicate Task Names"" imageMso=""RemoveDuplicates"" onAction=""cptFindDuplicateTaskNames"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResetRowHeight"" label=""Reset Row Height"" imageMso=""RowHeight"" onAction=""cptResetRowHeight"" visible=""true""/>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'trace tools
  If cptModuleExists("cptCriticalPathTools_bas") Or cptModuleExists("cptCriticalPath_bas") Or cptModuleExists("cptNetworkBrowser_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCPA"" label=""Trace and Mark"" visible=""true"">"
    If cptModuleExists("cptCriticalPathTools_bas") And cptModuleExists("cptCriticalPath_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:splitButton id=""sbTrace"" size=""large"" >"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" imageMso=""TaskDrivers"" label=""Driving Paths"" onAction=""DrivingPaths"" supertip=""Select a target task, get the primary, secondary, and tertiary driving paths to that task."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mTrace"">"
      ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Export"" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bPowerPoint"" label="">> PowerPoint"" imageMso=""SlideNew"" onAction=""cptExportCriticalPathSelected"" supertip=""Select a target task, get the primary, secondary, and tertiary driving paths to that task--and export them to PowerPoint."" />"
      ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
      ribbonXML = ribbonXML + vbCrLf & "</mso:splitButton>"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSinglePath"" label=""Driving Path"" imageMso=""TaskDrivers"" onAction=""cptDrivingPath"" visible=""true"" size=""large"" supertip=""Select a target task, get the driving path."" />"
    Else
      If cptModuleExists("cptCriticalPath_bas") Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bTrace"" label=""Driving Path"" imageMso=""TaskDrivers"" onAction=""DrivingPaths"" visible=""true"" size=""large"" />"
      End If
      If cptModuleExists("cptCriticalPathTools_bas") Then
        ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bExport"" label="">> PowerPoint"" imageMso=""SlideNew"" onAction=""cptExportCriticalPathSelected"" visible=""true"" size=""large"" />"
      End If
    End If
    If cptModuleExists("cptNetworkBrowser_bas") And cptModuleExists("cptNetworkBrowser_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bNetworkBrowser"" label=""Network Browser"" imageMso=""ViewPredecessorsSuccessorsShow"" onAction=""cptShowNetworkBrowser_frm"" visible=""true"" size=""large"" supertip=""Jump to, and/or trace, predecessors and successors using the Network Diagram view in full screen or in the details pane."" />"
    End If
    If cptModuleExists("cptSaveMarked_bas") And cptModuleExists("cptSaveMarked_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:separator id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""mark_selected"" label=""Mark"" imageMso=""ApproveApprovalRequest"" onAction=""cptMarkSelected"" visible=""true"" supertip=""Mark selected task(s)"" />" 'size=""large""
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""unmark_selected"" label=""Unmark"" imageMso=""RejectApprovalRequest"" onAction=""cptUnmarkSelected"" visible=""true"" supertip=""Unmark selected task(s)"" />" 'size=""large""
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""btnMarkedApply"" label=""Filter"" imageMso=""FilterToggleFilter"" onAction=""cptMarked"" visible=""true"" supertip=""Filter Marked Tasks"" />" 'size=""large""
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""btnSaveMarked"" label=""Save"" imageMso=""Archive"" onAction=""cptSaveMarked"" visible=""true"" supertip=""Save currently marked tasks for later import."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""btnImportMarked"" label=""Import"" imageMso=""ApproveApprovalRequest"" onAction=""cptShowSaveMarked_frm"" visible=""true"" supertip=""Import saved sets of marked tasks."" />"
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""clear_marked"" label=""Clear"" imageMso=""FilterClear"" onAction=""cptClearMarked"" visible=""true"" supertip=""Clear all currently marked tasks."" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'status
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gStatus"" label=""Status"" visible=""true"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mStatus"" label=""Status"" imageMso=""UpdateAsScheduled"" visible=""true"" size=""large"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Before Status"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  'todo: update project status date
  'todo: age dates [settings required]
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cpt_bAgeDates"" label=""Age Dates""  imageMso=""CalendarToolSelectDate"" onAction=""cptShowAgeDates_frm"" visible=""true"" supertip=""Keep a rolling history of the current schedule.""  />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Status Export &amp;&amp; Import"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  If cptModuleExists("cptStatusSheet_bas") And cptModuleExists("cptStatusSheet_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bStatusSheet"" label=""Create Status Sheet(s)"" imageMso=""ExportExcel"" onAction=""cptShowStatusSheet_frm"" visible=""true"" supertip=""Just what it sounds like. Include any fields you like. Settings are saved between sessions."" />" 'DateAndTimeInsertOneNote
  End If
  If cptModuleExists("cptStatusSheetImport_bas") And cptModuleExists("cptStatusSheetImport_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bStatusSheetImport"" label=""Import Status Sheet(s)"" imageMso=""ImportExcel"" onAction=""cptShowStatusSheetImport_frm"" visible=""true"" supertip=""Just what it sounds like. (Note: Assignment ETC is at the Assignment level, so use the Task Usage view to review after import.)"" />"
  End If
  If cptModuleExists("cptSmartDuration_frm") And cptModuleExists("cptSmartDuration_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSmartDuration"" label=""Smart Duration"" imageMso=""CalendarToolSelectDate"" onAction=""SmartDuration"" visible=""true"" supertip=""We've all been there: how many days between Time Now and the finish date the CAM just gave me? No more guess work: click here and improve your life."" />"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""After Status"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBlameReport"" label=""The Blame Report"" imageMso=""ContactProperties"" onAction=""cptBlameReport"" visible=""true"" supertip=""Find out which tasks slipped from last period."" />"
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCaptureWeek"" label=""Capture Week"" imageMso=""RefreshWebView"" onAction=""cptCaptureWeek"" visible=""true"" supertip=""Capture the Current Schedule to compare against past and future weeks during execution."" />"
  'todo: account for EV Tool in cptValidateEVP
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bValidateEVT"" enabled=""false"" label=""Validate EVT"" imageMso=""RefreshWebView"" onAction=""cptCaptureWeek"" visible=""true"" supertip=""Validate EVT - e.g., ensure incomplete 50/50 tasks with Actual Start are marked as 50% EV % complete."" />"
  'todo: changes from last week
  ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Status Settings"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
  'todo: age dates settings
  ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'snapshots

  'calendars
  If (cptModuleExists("cptFiscal_frm") And cptModuleExists("cptFiscal_bas")) Or (cptModuleExists("cptCalendarExceptions_frm") And cptModuleExists("cptCalendarExceptions_bas")) Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCalendars"" label=""Calendars"" visible=""true"" >"
    If cptModuleExists("cptFiscal_frm") And cptModuleExists("cptFiscal_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFiscal"" label=""Fiscal"" imageMso=""MonthlyView"" onAction=""cptShowFiscal_frm"" visible=""true"" supertip=""Maintain a fiscal calendar for various reports."" />"
    End If
    If cptModuleExists("cptCalendarExceptions_frm") And cptModuleExists("cptCalendarExceptions_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCalDetails"" label=""Details"" imageMso=""MonthlyView"" onAction=""cptShowCalendarExceptions_frm"" visible=""true"" supertip=""Export Calendar Exceptions, WorkWeeks, and settings."" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'resource allocation
  If cptModuleExists("cptResourceDemand_bas") And cptModuleExists("cptResourceDemand_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gResourceDemand"" label=""FTE"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bResourceDemandExcel"" label=""FTE"" imageMso=""Chart3DColumnChart"" onAction=""cptShowExportResourceDemand_frm"" visible=""true"" size=""large"" supertip=""Export timephased assignment remaining work, baseline work, costs (any or all rate sets), and your choice of extra fields. Settings are saved between sessions."" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If

  'allocation scenarios

  'compare

  'metrics
  If cptModuleExists("cptMetrics_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gMetrics"" label=""Metrics"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mSchedule"" label=""Schedule"" imageMso=""UpdateAsScheduled"" visible=""true"" size=""large"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Schedule Metrics"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    If Application.Version >= 12 Then 'CPLI only available in versions after 2010
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCPLI"" label=""Critical Path Length Index (CPLI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetCPLI"" visible=""true"" supertip=""Select a target task, clik to get the CPLI. Raw calculation based on time now and total slack; Schedule Margin not considered."" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBEI"" label=""Baseline Execution Index (BEI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetBEI"" visible=""true"" supertip=""Just what it sounds like."" />"
    'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCEI"" label=""Current Execution Index (CEI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetCEI"" visible=""true""/>"
    'todo: TFCI
    'todo: Earned Schedule
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bHit"" label=""Hit Task %"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetHitTask"" visible=""true"" supertip=""Because it's (still) on the Gold Card."" />"
    
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mEVish"" label=""EVish"" imageMso=""UpdateAsScheduled"" visible=""true"" size=""large"" supertip=""EV-ish metrics, based in hours. (Assumes schedule is resource-loaded using real assignments, rather than custom fields.)"" >"
    If cptModuleExists("cptMetricsSettings_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""cptMetricsSettings"" label=""Settings"" imageMso=""Settings"" onAction=""cptShowMetricsSettings_frm"" visible=""true"" supertip=""Settings required for some EV-ish metrics."" />"
    End If
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Earned Value-ish (in hrs)"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSPI"" label=""Schedule Performance Index (SPI)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetSPI"" visible=""true"" supertip=""Relies on timephased baseline work and Physical % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bSV"" label=""Schedule Variance (SV)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetSV"" visible=""true"" supertip=""Relies on timephased baseline work and Physical % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBCWS"" label=""Budgeted Cost of Work Scheduled (BCWS)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetBCWS"" visible=""true"" supertip=""Timephased BCWS/PV (in hours)."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBCWP"" label=""Budgeted Cost of Work Performed (BCWP)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetBCWP"" visible=""true"" supertip=""Timephased BCWP/EV (in hours)--relies on baseline work and Physical % Complete."" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBAC"" label=""Budget at Complete (BAC)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetBAC"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bETC"" label=""Estimate to Complete (ETC)"" imageMso=""ApplyPercentageFormat"" onAction=""cptGetETC"" visible=""true""/>"
    'ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator title=""Export"" id=""cleanup_" & cptIncrement(lngCleanUp) & """ />"
    'ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bExportMetrics"" label="">> Excel"" imageMso=""ExportExcel"" onAction=""cptExportMetricsExcel"" visible=""true""/>"
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
'    If cptModuleExists("cptGraphics_bas") And cptModuleExists("cptGraphics_frm") Then
'      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bGraphics"" label=""Quick Look"" imageMso=""PivotChartInsert"" onAction=""cptShowFrmGraphics"" visible=""true"" size=""large"" />"
'    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'outline codes
  If cptModuleExists("cptBackbone_frm") And cptModuleExists("cptBackbone_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gWBS"" label=""Backbone"" visible=""true"" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBackbone"" label=""Outline Codes"" imageMso=""WbsMenu"" onAction=""cptShowBackbone_frm"" visible=""true"" size=""large"" supertip=""Quickly create or edit Outline Codes (CWBS, IMP, etc.); import and/or export; create DI-MGMT-81334D, etc."" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'integration
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gIntegration"" label=""Integration"" visible=""true"" >"
  If cptModuleExists("cptIMSCobraExport_bas") And cptModuleExists("cptIMSCobraExport_frm") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCOBRA"" label=""COBRA Export Tool"" imageMso=""Export"" onAction=""Export_IMS"" visible=""true"" supertip=""Validate that your IMS is ready for integration; create CSV transaction files for COBRA. Baseline, forecast, status, etc."" />"
  End If
  If cptModuleExists("cptCheckAssignments_bas") Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bCheckAssignments"" label=""Check Assignments"" imageMso=""SynchronizationStatus"" onAction=""cptCheckAssignments"" visible=""true"" supertip=""Reconcile task vs assignment work, baselines, etc."" />"
  End If
  'mpm
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  'bcr

  'custom fields
  If (cptModuleExists("cptDataDictionary_frm") And cptModuleExists("cptDataDictionary_bas")) Or (cptModuleExists("cptSaveLocal_bas") And cptModuleExists("cptSaveLocal_frm")) Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gCustomFields"" label=""Custom Fields"" visible=""true"">"
    If cptModuleExists("cptDataDictionary_frm") And cptModuleExists("cptDataDictionary_bas") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bDataDictionary"" imageMso=""ReadingMode"" label=""Data Dictionary"" onAction=""cptShowDataDictionary_frm"" supertip=""Provide a description of each custom field; create export in Excel for deliverables; share dictionary. Settings are saved between sessions."" />" 'size=""large""
    End If
    If cptModuleExists("cptFieldBuilder_bas") And cptModuleExists("cptFieldBuilder_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bBuilder"" imageMso=""CustomFieldDialog"" label=""Field Builder"" onAction=""cptShowFieldBuilder_frm"" supertip=""A little help building common custom field pick lists, etc."" />" 'size=""large""
    End If
    If cptModuleExists("cptSaveLocal_bas") And cptModuleExists("cptSaveLocal_frm") Then
      ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bECFtoLCF"" imageMso=""CustomFieldDialog"" label=""ECF to LCF"" onAction=""cptShowSaveLocal_frm"" supertip=""Save Enterprise Custom Fields (ECF) and data to Local Custom Fields (LCF). Settings are saved (by project) between sessions."" />" 'size=""large""
    End If
    ribbonXML = ribbonXML + vbCrLf & "</mso:group>"
  End If
  
  'about
  ribbonXML = ribbonXML + vbCrLf & "<mso:group id=""gHelp"" label=""Help"" visible=""true"" >"
  If cptInternetIsConnected Then
    ribbonXML = ribbonXML + vbCrLf & "<mso:menu id=""mHelp"" label=""Help"" imageMso=""Help"" visible=""true"" size=""large"" supertip=""Submit a bug report, feature request, or general feedback. Upgrade modules from the InterWebs."" >"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Upgrades"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUpdate"" label=""Check for Upgrades"" imageMso=""PreviousUnread"" onAction=""cptShowUpgrades_frm"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Contribute"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bIssue"" label=""Submit an Issue"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitIssue"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bRequest"" label=""Submit a Feature Request"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitRequest"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bFeedback"" label=""Submit Other Feedback"" imageMso=""SubmitFormInfoPath"" onAction=""cptSubmitFeedback"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:menuSeparator id=""cleanup_" & cptIncrement(lngCleanUp) & """ title=""Remove"" />"
    ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bUninstall"" label=""Uninstall ClearPlan Toolbar"" imageMso=""TasksUnlink"" onAction=""cptUninstall"" visible=""true"" />"
    ribbonXML = ribbonXML + vbCrLf & "</mso:menu>"
  End If
  ribbonXML = ribbonXML + vbCrLf & "<mso:button id=""bAbout"" onAction=""cptShowAbout_frm""  size=""large"" visible=""true""  label=""About"" imageMso=""Info"" />"
  ribbonXML = ribbonXML + vbCrLf & "</mso:group>"

  ribbonXML = ribbonXML + vbCrLf & "</mso:tab>"

  'Debug.Print "<mso:customUI ""xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >" & ribbonXML
  cptBuildRibbonTab = ribbonXML

End Function

Sub cptHandleErr(strModule As String, strProcedure As String, objErr As ErrObject, Optional lngErl As Long)
'common error handling prompt
Dim strMsg As String

    strMsg = "Please contact cpt@ClearPlanConsulting.com for assistance if needed." & vbCrLf & vbCrLf
    strMsg = strMsg & "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf
    strMsg = strMsg & "Source: " & strModule & "." & strProcedure
    If lngErl > 0 Then
      strMsg = strMsg & ":" & lngErl
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
  If Err.Number = 5 Then
    cptRegEx = ""
    Err.Clear
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
  Call cptHandleErr("cptSetup_bas", "cptModuleExists", Err, Erl)
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
      ThisProject.VBProject.VBComponents.Remove vbComponent
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
  Call cptHandleErr("cptSetup_bas", "cptUninstall", Err, Erl)
  Resume exit_here
End Sub
