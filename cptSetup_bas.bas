Attribute VB_Name = "cptSetup_bas"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub cptSetup()
'objects
Dim arrCode As Object
Dim cmThisProject As CodeModule
Dim cmCptThisProject As CodeModule
Dim oStream As Object
Dim xmlHttpDoc As Object
Dim xmlNode As Object
Dim xmlDoc As Object
Dim arrCore As Object
'strings
Dim strCptFileName As String
Dim strVersion As String
Dim strCode As String
Dim strFileName As String
Dim strModule As String
Dim strError As String
Dim strURL As String
'longs
Dim lngEvent As Long
Dim lngActivate As Long
Dim lngFile As Long
'integers
'booleans
Dim blnSuccess As Boolean
'variants
Dim vEvent As Variant
Dim vLine As Variant
'dates

  'before running:
  '1. enable macros
  '2. trust access to the vbproject object model
  '3. save global.mpt, completely exist msproject and restart (to make the settings 'stick')
  '4. come back and run cptSetup
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  If Not InternetIsConnected Then
    MsgBox "If you cannot access the internet, then please pull the Core.mpp file from Teams and follow instructions there.", vbExclamation + vbOKOnly, "No Internet"
    GoTo exit_here
  End If
    
  'capture list of files to download
  Set arrCore = CreateObject("System.Collections.SortedList")
  
  'get CurrentVersions.xml
  'get file list in cpt\Core
  strURL = "https://raw.githubusercontent.com/AronGahagan/test/master/CurrentVersions.xml"
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
    arrCore.Clear
    For Each xmlNode In xmlDoc.SelectNodes("/Modules/Module")
      If xmlNode.SelectSingleNode("Directory").Text = "Core" Then
        arrCore.Add xmlNode.SelectSingleNode("FileName").Text, xmlNode.SelectSingleNode("Type").Text
        If xmlNode.SelectSingleNode("FileName").Text = "ThisProject.cls" Then
          strVersion = xmlNode.SelectSingleNode("Version").Text
        End If
        Debug.Print "found " & xmlNode.SelectSingleNode("FileName").Text
      End If
    Next
  End If
  
  'download cpt/core/*.* to user's tmp directory
  For lngFile = 0 To arrCore.count - 1
frx:
    strURL = "https://raw.githubusercontent.com/AronGahagan/test/master/Core/" & arrCore.getKey(lngFile)
    Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
    xmlHttpDoc.Open "GET", strURL, False
    xmlHttpDoc.Send

    If xmlHttpDoc.Status = 200 Then
      Set oStream = CreateObject("ADODB.Stream")
      oStream.Open
      oStream.Type = 1 'adTypeBinary
      oStream.Write xmlHttpDoc.responseBody
      strFileName = arrCore.getKey(lngFile)
      If Dir(Environ("tmp") & "\" & strFileName) <> vbNullString Then Kill Environ("tmp") & "\" & strFileName
      oStream.SaveToFile Environ("tmp") & "\" & strFileName
      oStream.Close
    Else
      strError = strError & "- " & arrCore.getKey(lngFile) & vbCrLf
      GoTo next_file
    End If
    If Right(strFileName, 4) = ".frm" Then
      strFileName = Replace(strFileName, ".frm", ".frx")
      GoTo frx
    ElseIf Right(strFileName, 4) = ".frx" Then
      strFileName = Replace(strFileName, ".frx", ".frm")
    End If
    strModule = Left(strFileName, InStr(strFileName, ".") - 1)
    If ModuleExists(strModule) Then
      'ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(strModule)
    End If
    'skip ThisProject which needs special handling
    If strModule <> "ThisProject" Then
      'ThisProject.VBProject.VBComponents.Import Environ("tmp") & "\" & strFileName
    End If
next_file:
  Next lngFile
  
  'update user's ThisProject - if it downloaded correctly
  strFileName = Environ("tmp") & "\ThisProject.cls"
  
  If Dir(strFileName) <> vbNullString Then
    
    'rename the file and import it
    strCptFileName = Replace(strFileName, "ThisProject", "cptThisProject")
    If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName
    Name strFileName As strCptFileName
    Set cmThisProject = ThisProject.VBProject.VBComponents("ThisProject").CodeModule
    Set cmCptThisProject = ThisProject.VBProject.VBComponents.Import(strCptFileName).CodeModule
    
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
              Debug.Print CStr(vEvent) & " code exists."
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
    Debug.Print strError
  End If

  If ModuleExists("cptUpgrades_frm") And ModuleExists("cptUpgrades_bas") Then
    Call ShowCptUpgrades_frm
  End If
  

exit_here:
  On Error Resume Next
  Set arrCode = Nothing
  SpeedOFF
  Set cmThisProject = Nothing
  Set cmCptThisProject = Nothing
  Set oStream = Nothing
  Set xmlHttpDoc = Nothing
  Set xmlNode = Nothing
  Set xmlDoc = Nothing
  Set arrCore = Nothing

  Exit Sub
err_here:
  Call HandleErr("cptCore_bas", "cptSetup", err)
  Resume exit_here
End Sub
