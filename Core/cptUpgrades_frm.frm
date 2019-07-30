VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptUpgrades_frm 
   Caption         =   "Installation Status"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   OleObjectBlob   =   "cptUpgrades_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptUpgrades_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.3.7</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdUpgradeAll_Click()
Dim lngItem As Long

  For lngItem = 0 To Me.lboModules.ListCount - 1
    If Me.lboModules.List(lngItem, 2) <> Me.lboModules.List(lngItem, 3) Then
      Me.lboModules.Selected(lngItem) = True
    End If
  Next lngItem
  
  Call cmdUpgradeSelected_Click
  
End Sub

Private Sub cmdUpgradeSelected_Click()
'objects
Dim arrCode As Object
Dim cmCptThisProject As Object
Dim cmThisProject As Object
Dim Project As Object
Dim vbComponent As Object
Dim xmlHttpDoc As Object
Dim oStream As Object 'ADODB.Stream
Dim arrCurrent As Object
Dim arrInstalled As Object
Dim arrTypes As Object
'strings
Dim lngEvent As String
Dim strVersion As String
Dim strMsg As String
Dim strCptFileName As String
Dim strDirectory As String
Dim strModule As String, strFileName As String, strURL As String
'longs
Dim lngLine As Long
Dim lngItem As Long
'integers
'booleans
'variants
Dim vEvent As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set arrTypes = CreateObject("System.Collections.SortedList")
  arrTypes.Add 1, ".bas"
  arrTypes.Add 2, ".cls"
  arrTypes.Add 3, ".frm"
  arrTypes.Add 100, ".cls"

  For lngItem = 0 To Me.lboModules.ListCount - 1

    If Me.lboModules.Selected(lngItem) Then
      
      '<issue33> trap invalid use of null error?
      If IsNull(Me.lboModules.List(lngItem, 0)) Then
        MsgBox "Unable to download upgrades.", vbExclamation + vbOKOnly, "Can't Connect"
        GoTo exit_here
      End If '</issue33>
      
      Me.lboModules.List(lngItem, 3) = "<installing...>"

      strModule = Me.lboModules.List(lngItem, 0)

      'get the module name
      'get the repo directory
      'get the filename
      Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
      strFileName = strModule & arrTypes.Item(CInt(cptUpgrades_frm.lboModules.List(lngItem, 5)))
      strDirectory = cptUpgrades_frm.lboModules.List(lngItem, 1)
get_frx:
      strURL = strGitHub & strDirectory & "/" & strFileName
      xmlHttpDoc.Open "GET", strURL, False
      xmlHttpDoc.Send
      
      'strURL = xmlHttpDoc.responseBody
      If xmlHttpDoc.Status = 200 And xmlHttpDoc.readyState = 4 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'adTypeBinary
        oStream.Write xmlHttpDoc.responseBody
        If Dir(cptDir & "\" & strFileName) <> vbNullString Then Kill cptDir & "\" & strFileName
        oStream.SaveToFile cptDir & "\" & strFileName
        oStream.Close
      Else
        MsgBox "Download failed. Please contact cpt@ClearPlanConsulting.com for help.", vbCritical + vbOKOnly, "XML Error"
        Me.lboModules.List(lngItem, 3) = "<failed>"
        GoTo exit_here
      End If
      If Right(strFileName, 4) = ".frm" Then
        strFileName = Replace(strFileName, ".frm", ".frx")
        GoTo get_frx
      ElseIf Right(strFileName, 4) = ".frx" Then
        strFileName = Replace(strFileName, ".frx", ".frm")
      End If

      '<issue15> added
      If strModule = "ThisProject" Then GoTo next_module 'handle separately </issue25>

      If cptModuleExists(strModule) Then
        '<issue19>
        Set vbComponent = ProjectGlobal.ThisProject.VBProject.VBComponents(strModule) '<issue61>
        vbComponent.Name = vbComponent.Name & "_" & Format(Now, "hhnnss")
        DoEvents
        ProjectGlobal.ThisProject.VBProject.VBComponents.remove vbComponent 'ThisProject.VBProject.VBComponents(strModule) '<issue61>
        DoEvents '</issue19>
      End If
      ProjectGlobal.ThisProject.VBProject.VBComponents.import cptDir & "\" & strFileName '<issue61>
      
      '<issue24> remove the whitespace added by VBE import/export
      With ProjectGlobal.ThisProject.VBProject.VBComponents(strModule).CodeModule '<issue61>
        For lngLine = .CountOfDeclarationLines To 1 Step -1
          If Len(.Lines(lngLine, 1)) = 0 Then .DeleteLines lngLine, 1
        Next lngLine
      End With '</issue24>
      
      Me.lboModules.List(lngItem, 3) = Me.lboModules.List(lngItem, 2)
      Me.lboModules.List(lngItem, 4) = "<updated>"
    End If
next_module:     '</issue25>
  Next lngItem

  '<issue25> added
  'update ThisProject
  strFileName = cptDir & "\ThisProject.cls"
  If Dir(strFileName) <> vbNullString Then 'the file was downloaded, proceed

    'notify user that modifications are about to be made to the ThisProject module
    strMsg = "This upgrade requires a revision to your ThisProject module. "
    strMsg = strMsg & "If you have made modifications, your code will not be lost, but it may need to be rearanged." & vbCrLf & vbCrLf
    strMsg = strMsg & "Please contact cpt@ClearPlanConsulting.com if you require assistance."
    MsgBox strMsg, vbInformation + vbOKOnly, "Notice"
    'ideally this would prompt user to proceed or rollback...

    'clear out existing lines of cpt-related code
    Set cmThisProject = ProjectGlobal.ThisProject.VBProject.VBComponents("ThisProject").CodeModule '<issue61>
    For lngLine = cmThisProject.CountOfLines To 1 Step -1
      'cover both '</cpt_version> and '</cpt>
      If InStr(cmThisProject.Lines(lngLine, 1), "</cpt") > 0 Then
        cmThisProject.DeleteLines lngLine, 1
        DoEvents
      End If
    Next lngLine

    'rename file and import it
    strCptFileName = Replace(strFileName, "ThisProject", "cptThisProject")
    Name strFileName As strCptFileName
    Set cmCptThisProject = ProjectGlobal.ThisProject.VBProject.VBComponents.import(strCptFileName).CodeModule '<issue61>
    'grab and insert the updated version
    strVersion = cptRegEx(cmCptThisProject.Lines(1, cmCptThisProject.CountOfLines), "<cpt_version>.*</cpt_version>")
    cmThisProject.InsertLines 1, "'" & strVersion

    'grab the imported code
    Set arrCode = CreateObject("System.Collections.SortedList")
    With cmCptThisProject
      For Each vEvent In Array("Project_Activate", "Project_Open")
        arrCode.Add CStr(vEvent), .Lines(.ProcStartLine(CStr(vEvent), 0) + 2, .ProcCountLines(CStr(vEvent), 0) - 3) '0 = vbext_pk_Proc
      Next vEvent
    End With
    ProjectGlobal.ThisProject.VBProject.VBComponents.remove ProjectGlobal.ThisProject.VBProject.VBComponents(cmCptThisProject.Parent.Name) '<issue61>
    '<issue19> added
    DoEvents '</issue19>

    'add the events, or insert new text
    'three cases: empty or not (code exists or not)
    For Each vEvent In Array("Project_Activate", "Project_Open")

      'if event exists then insert code else create new event handler
      With cmThisProject
        If .CountOfLines > .CountOfDeclarationLines Then 'complications
          If .Find("Sub " & CStr(vEvent), 1, 1, .CountOfLines, 1000) = True Then
          'find its line number
            lngEvent = .ProcBodyLine(CStr(vEvent), 0)  '= vbext_pk_Proc
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

      'leave no trace
      If Dir(strCptFileName) <> vbNullString Then Kill strCptFileName

    Next vEvent
  End If '</issue25>

  'reset the ribbon
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
  Set arrCode = Nothing
  Set cmCptThisProject = Nothing
  Set cmThisProject = Nothing
  Application.ScreenUpdating = True
  Set Project = Nothing
  Set vbComponent = Nothing
  Application.StatusBar = ""
  Set arrTypes = Nothing
  Set xmlHttpDoc = Nothing
  Set arrCurrent = Nothing
  Set arrInstalled = Nothing
  Set oStream = Nothing
  Exit Sub
err_here:
  Call cptHandleErr("cptUpgrades_frm", "cmdUpdate_Click", err, Erl)
  Me.lboModules.List(lngItem - 1, 3) = "<error>" '</issue25>
  Resume exit_here

End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptUpgrades_frm", "lblURL_Click", err, Erl)
  Resume exit_here
End Sub
