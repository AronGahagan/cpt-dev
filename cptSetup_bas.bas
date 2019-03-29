Attribute VB_Name = "cptSetup_bas"
'<cpt_version>v1.1</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Public Const strGitHub = "https://raw.githubusercontent.com/AronGahagan/cpt-dev/master/"
'Public Const strGitHub = "https://raw.githubusercontent.com/ClearPlan/cpt/master/"

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As Long) As Long

Sub cptSetup()
'setup only needs to be run once
'objects
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
Dim strCode As String
Dim strFileName As String
Dim strModule As String
Dim strURL As String
'longs
Dim lngLine As Long
Dim lngEvent As Long
Dim lngActivate As Long
Dim lngFile As Long
'integers
'booleans
Dim blnExists As Boolean
Dim blnSuccess As Boolean
'variants
Dim vEvent As Variant
Dim vLine As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
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
    MsgBox xmlDoc.parseError.ErrorCode & ": " & xmlDoc.parseError.reason, vbExclamation + vbOKOnly, "XML Error"
    GoTo exit_here
  Else
    'download cpt/core/*.* to user's tmp directory
    arrCore.Clear
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      If xmlNode.SelectSingleNode("Directory").Text = "Core" Then
        Application.StatusBar = "Fetching " & xmlNode.SelectSingleNode("Name").Text & "..."
        'Debug.Print Application.StatusBar
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
          strError = strError & "- " & arrCore.getKey(lngFile) & vbCrLf
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
            ThisProject.VBProject.VBComponents.remove ThisProject.VBProject.VBComponents(CStr(vbComponent.Name))
            Exit For
          End If
        Next vbComponent
        
        'import the module - skip ThisProject which needs special handling
        If strModule <> "ThisProject" Then
          Application.StatusBar = "Importing " & strFileName & "..."
          'Debug.Print Application.StatusBar
          ThisProject.VBProject.VBComponents.import cptDir & "\" & strFileName
        End If
        
      End If
next_xmlNode:
    Next xmlNode
  End If
  
  Application.StatusBar = "CPT Modules imported."
  
  'update user's ThisProject - if it downloaded correctly
  strFileName = cptDir & "\ThisProject.cls"
  
  If Dir(strFileName) <> vbNullString Then 'the file exists, proceed
    
    'avoid messy overwrites of ThisProject
    Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
    If cmThisProject.Find("<cpt_version>", 1, 1, cmThisProject.CountOfLines, 1000, True, True) = True Then
      strMsg = "Your 'ThisProject' module has already been updated to work with the ClearPlan toolbar." & vbCrLf
      strMsg = strMsg & "Would you like to reset it? This will only overwrite CodeModule lines appended with '</cpt>'"
      If MsgBox(strMsg, vbExclamation + vbYesNo, "Danger, Will Robinson!") = vbYes Then
        For lngLine = cmThisProject.CountOfLines To 1 Step -1
          If InStr(cmThisProject.Lines(lngLine, 1), "'</cpt>") > 0 Then
            cmThisProject.DeleteLines lngLine
            Debug.Print "DELETED: " & lngLine & ": " & cmThisProject.Lines(lngLine)
          End If
        Next lngLine
      End If
    End If
    
    'rename the file and import it
    strCptFileName = Replace(strFileName, "ThisProject", "cptThisProject")
    If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName
    Name strFileName As strCptFileName
    Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
    Set cmCptThisProject = ThisProject.VBProject.VBComponents.import(strCptFileName).CodeModule
    
    'grab the imported code
    Set arrCode = CreateObject("System.Collections.SortedList")
    With cmCptThisProject
      For Each vEvent In Array("Project_Activate", "Project_Open")
        arrCode.Add CStr(vEvent), .Lines(.ProcStartLine(CStr(vEvent), 0) + 2, .ProcCountLines(CStr(vEvent), 0) - 3) '0 = vbext_pk_Proc
      Next
    End With
    ThisProject.VBProject.VBComponents.remove ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name)
    
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
    If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName
    
  End If 'ThisProject.cls exists in tmp folder
  
  If Len(strError) > 0 Then
    strError = "The following modules did not download correctly:" & vbCrLf & strError & vbCrLf & vbCrLf & "Please contact cpt@ClearPlanConsulting.com for assistance."
    MsgBox strError, vbCritical + vbOKOnly, "Unknown Error"
    'Debug.Print strError
  End If

exit_here:
  On Error Resume Next
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
  'Call cptHandleErr("cptCore_bas", "cptSetup", err)
  strError = err.Number & ": " & err.Description & vbCrLf
  strError = strError & "Module: cptSetup_bas" & vbCrLf
  strError = strError & "Procedure: cptSetup"
  MsgBox strError, vbExclamation + vbOKOnly, "CPT Setup Error"
  Resume exit_here
End Sub

Public Function cptInternetIsConnected() As Boolean
 
    cptInternetIsConnected = InternetGetConnectedStateEx(0, "", 254, 0)
 
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

Function cptModuleExists(strModule)
Dim vbComponent As Object
Dim blnExists As Boolean
Dim strError As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  blnExists = False
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
  'Call cptHandleErr("cptSetup_bas", "cptModuleExists", err)
  strError = err.Number & ": " & err.Description & vbCrLf
  strError = strError & "Module: cptSetup_bas" & vbCrLf
  strError = strError & "Procedure: cptModuleExists"
  MsgBox strError, vbExclamation + vbOKOnly, "CPT Setup Error"
  
  Resume exit_here
  
End Function


