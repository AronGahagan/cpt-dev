Attribute VB_Name = "cptGA"
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Function cptGetUserCustomizations(strGetWhat As String) As String
  'objects
  Dim oStream As Scripting.TextStream
  Dim oFile As Scripting.File
  Dim oFSO As Scripting.FileSystemObject
  'strings
  Dim strFile As String
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strFile = Environ("USERPROFILE") & "\Project Customizations.exportedUI"

  Set oFSO = CreateObject("Scripting.FileSystemObject")
  If Not oFSO.FileExists(strFile) Then
    MsgBox "'Project Customizations.exportedUI'" & vbCrLf & "does not exist in location:" & vbCrLf & Environ("USERPROFILE"), vbExclamation + vbOKOnly, "File Not Found"
    cptGetUserCustomizations = ""
    GoTo exit_here
  End If
  
  Set oFile = oFSO.GetFile(strFile)
  Set oStream = oFile.OpenAsTextStream
  
  If strGetWhat = "all" Then
    cptGetUserCustomizations = oStream.ReadAll
  ElseIf strGetWhat = "xmlns" Then
    cptGetUserCustomizations = Replace(cptRegEx(oStream.ReadAll, "<mso:customUI.*><mso:ribbon"), "<mso:ribbon", "") 'ribbon might have attributes
  ElseIf strGetWhat = "qat" Then
    cptGetUserCustomizations = cptRegEx(oStream.ReadAll, "<mso:qat>.*</mso:qat>")
  ElseIf strGetWhat = "tabs" Then
    cptGetUserCustomizations = Replace(Replace(cptRegEx(oStream.ReadAll, "<mso:tabs>.*</mso:tabs>", True), "<mso:tabs>", ""), "</mso:tabs>", "")
  ElseIf strGetWhat = "contextualTabs" Then
    cptGetUserCustomizations = cptRegEx(oStream.ReadAll, "<mso:contextualTabs>.*</mso:contextualTabs>")
  End If
  
  oStream.Close
  
exit_here:
  On Error Resume Next
  Set oStream = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing

  Exit Function
err_here:
  'Call HandleErr("cptGA", "GetUserCustomizations", Err)
  MsgBox Err.Number & ": " & Err.Description, vbInformation + vbOKOnly, "Error"
  Resume exit_here
End Function
