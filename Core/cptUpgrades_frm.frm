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
'<cpt_version>v1.2</cpt_version>
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
Dim xmlHttpDoc As Object, oStream As ADODB.Stream ' Object
Dim arrCurrent As Object, arrInstalled As Object
Dim arrTypes As Object
'strings
Dim strDirectory As String
Dim strModule As String, strFileName As String, strURL As String
'longs
Dim lngItem As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Set arrTypes = CreateObject("System.Collections.SortedList")
  arrTypes.Add 1, ".bas"
  arrTypes.Add 2, ".cls"
  arrTypes.Add 3, ".frm"
  arrTypes.Add 100, ".cls"

  For lngItem = 0 To Me.lboModules.ListCount - 1
  
    If Me.lboModules.Selected(lngItem) Then
  
      Me.lboModules.List(lngItem, 3) = "<installing...>"
      strModule = Me.lboModules.List(lngItem, 0)
            
      'get the module name
      'get the repo directory
      'get the filename
      'strFileName = replace(".bas","_bas.bas")
      Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
      strFileName = strModule & arrTypes.Item(CInt(cptUpgrades_frm.lboModules.List(lngItem, 5)))
      'strFileName = Replace(strFileName, RegEx(strFileName, "_frm|_bas|_cls"), "")
      strDirectory = cptUpgrades_frm.lboModules.List(lngItem, 1)
get_frx:
      strURL = strGitHub & strDirectory & "/" & strFileName
      xmlHttpDoc.Open "GET", strURL, False
      xmlHttpDoc.Send
    
      'strURL = xmlHttpDoc.responseBody
      If xmlHttpDoc.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'adTypeBinary
        oStream.Write xmlHttpDoc.responseBody
        'strFileName = cptDir & strFileName
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
      
      If ModuleExists(strModule) Then
        ThisProject.VBProject.VBComponents.remove ThisProject.VBProject.VBComponents(strModule)
      End If
      ThisProject.VBProject.VBComponents.import cptDir & "\" & strFileName
      
      Me.lboModules.List(lngItem, 3) = Me.lboModules.List(lngItem, 2)
      Me.lboModules.List(lngItem, 4) = "<updated>"
    End If
  Next lngItem

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Set arrTypes = Nothing
  Set xmlHttpDoc = Nothing
  Set arrCurrent = Nothing
  Set arrInstalled = Nothing
  Set oStream = Nothing
  Exit Sub
err_here:
  Call HandleErr("frmUpdates", "cmdUpdate_Click", err)
  Me.lboModules.List(lngItem, 3) = "<error>"
  Resume exit_here

End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If InternetIsConnected Then Application.OpenBrowser "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptUpgrades_frm", "lblURL_Click", err)
  Resume exit_here
End Sub
