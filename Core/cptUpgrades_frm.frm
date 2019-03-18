VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptUpgrades_frm 
   Caption         =   "Installation Status"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "cptUpgrades_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptUpgrades_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.1</cpt_version>

Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdUpgradeAll_Click()
Dim lgItem As Long

  For lgItem = 0 To Me.lboModules.ListCount - 1
    If Me.lboModules.List(lgItem, 2) <> Me.lboModules.List(lgItem, 2) Then
      Me.lboModules.Selected(lgItem) = True
    End If
  Next lgItem
  
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
Dim lgItem As Long
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

  For lgItem = 0 To Me.lboModules.ListCount - 1
  
    If Me.lboModules.Selected(lgItem) Then
  
      Me.lboModules.List(lgItem, 3) = "<installing...>"
      strModule = Me.lboModules.List(lgItem, 0)
            
      'get the module name
      'get the repo directory
      'get the filename
      'strFileName = replace(".bas","_bas.bas")
      Set xmlHttpDoc = CreateObject("Microsoft.XMLHTTP")
      strFileName = strModule & arrTypes.Item(CInt(cptUpgrades_frm.lboModules.List(lgItem, 4)))
      'strFileName = Replace(strFileName, RegEx(strFileName, "_frm|_bas|_cls"), "")
      strDirectory = Left(strFileName, InStr(strFileName, ".") - 1)
get_frx:
      strURL = "https://raw.githubusercontent.com/AronGahagan/test/master/" & Replace(GetDirectory(strFileName), "\", "") & "/" & strFileName
      xmlHttpDoc.Open "GET", strURL, False
      xmlHttpDoc.Send
    
      'strURL = xmlHttpDoc.responseBody
      If xmlHttpDoc.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1 'adTypeBinary
        oStream.Write xmlHttpDoc.responseBody
        'strFileName = Environ("tmp") & strFileName
        If Dir(Environ("tmp") & "\" & strFileName) <> vbNullString Then Kill Environ("tmp") & "\" & strFileName
        oStream.SaveToFile Environ("tmp") & "\" & strFileName
        oStream.Close
      Else
        MsgBox "Download failed. Please contact our Help Desk at...", vbCritical + vbOKOnly, "XML Error"
        Me.lboModules.List(lgItem, 3) = "<failed>"
        GoTo exit_here
      End If
      If Right(strFileName, 4) = ".frm" Then
        strFileName = Replace(strFileName, ".frm", ".frx")
        GoTo get_frx
      ElseIf Right(strFileName, 4) = ".frx" Then
        strFileName = Replace(strFileName, ".frx", ".frm")
      End If
      
      If ModuleExists(strModule) Then
        ThisProject.VBProject.VBComponents.Remove ThisProject.VBProject.VBComponents(strModule)
      End If
      ThisProject.VBProject.VBComponents.Import Environ("tmp") & "\" & strFileName
      
      Me.lboModules.List(lgItem, 2) = Me.lboModules.List(lgItem, 1)
      Me.lboModules.List(lgItem, 3) = "<updated>"
    End If
  Next lgItem

exit_here:
  On Error Resume Next
  Set arrTypes = Nothing
  Set xmlHttpDoc = Nothing
  Set arrCurrent = Nothing
  Set arrInstalled = Nothing
  Set oStream = Nothing
  Exit Sub
err_here:
  Call HandleErr("frmUpdates", "cmdUpdate_Click", err)
  Me.lboModules.List(lgItem, 3) = "<error>"
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
