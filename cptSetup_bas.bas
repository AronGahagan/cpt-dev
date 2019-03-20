Attribute VB_Name = "cptSetup_bas"
'<cpt_version>v1.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Public Const strGitHub = "https://raw.githubusercontent.com/AronGahagan/cpt-dev/master/"
'Public Const strGitHub = "https://raw.githubusercontent.com/AronGahagan/cpt/master/"

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, _
                                                                        ByVal lpszConnectionName As String, _
                                                                        ByVal dwNameLen As Integer, _
                                                                        ByVal dwReserved As Long) As Long

Sub cptSetup()
'objects
Dim vbComponent As Object
Dim arrCode As Object
Dim cmThisProject As CodeModule
Dim cmCptThisProject As CodeModule
Dim oStream As Object
Dim xmlHttpDoc As Object
Dim xmlNode As Object
Dim xmlDoc As Object
Dim arrCore As Object
'strings
Dim strError As String
Dim strCptFileName As String
Dim strVersion As String
Dim strCode As String
Dim strFileName As String
Dim strModule As String
Dim strURL As String
'longs
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

  'before running:
  '1. enable macros
  '2. trust access to the vbproject object model
  '3. save global.mpt, completely exit msproject and restart (to make the settings 'stick')
  '4. come back and run cptSetup
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
'  If Not InternetIsConnected Then
'    MsgBox "If you cannot access the internet, then please pull the Core.mpp file from Teams and follow instructions there.", vbExclamation + vbOKOnly, "No Internet"
'    GoTo exit_here
'  End If
    
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
            ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(CStr(vbComponent.Name))
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
  
  If Dir(strFileName) <> vbNullString Then
    
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
    ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name)
    
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
  'Call HandleErr("cptCore_bas", "cptSetup", err)
  strError = err.Number & ": " & err.Description & vbCrLf
  strError = strError & "Module: cptSetup_bas" & vbCrLf
  strError = strError & "Procedure: cptSetup"
  MsgBox strError, vbExclamation + vbOKOnly, "CPT Setup Error"
  Resume exit_here
End Sub

Public Function InternetIsConnected() As Boolean
 
    InternetIsConnected = InternetGetConnectedStateEx(0, "", 254, 0)
 
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
  Call HandleErr("cptSetup_bas", "ModuleExists", err)
  Resume exit_here
  
End Function

