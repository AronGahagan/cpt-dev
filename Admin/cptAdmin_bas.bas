Attribute VB_Name = "cptAdmin_bas"
'>no cpt version - not for release<
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Sub CreateCurrentVersionsXML(strRepo As String)
'objects
Dim arrTypes As Object
Dim oStream As Object, vbComponent As Object 'adodb.stream
'strings
Dim strModule As String
Dim strDirectory As String
Dim strFile As String
Dim strXML As String, strVersion As String, strFileName As String
'longs
Dim lngFile As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'use arrTypes
  Set arrTypes = CreateObject("System.Collections.SortedList")
  arrTypes.Add 1, ".bas"
  arrTypes.Add 2, ".cls"
  arrTypes.Add 3, ".frm"
  arrTypes.Add 100, ".cls"

  'write xml
  strXML = "<?xml version=""1.0"" encoding=""utf-8"" ?>" & vbCrLf
  strXML = strXML & "<Modules>" & vbCrLf
  For Each vbComponent In ThisProject.VBProject.VBComponents
    If vbComponent.Name = "cptAdmin_bas" Then GoTo next_vbComponent
    If vbComponent.CodeModule.Find("<cpt_version>", 1, 1, vbComponent.CodeModule.CountOfLines, 25) = True Then
      strVersion = RegEx(vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines), "<cpt_version>.*</cpt_version>")
      strVersion = Replace(Replace(strVersion, "<cpt_version>", ""), "</cpt_version>", "")
      strXML = strXML & String(1, vbTab) & "<Module>" & vbCrLf
      strModule = Replace(vbComponent.Name, RegEx(vbComponent.Name, "_frm|_bas|_cls"), "")
      strXML = strXML & String(2, vbTab) & "<Name>" & vbComponent.Name & "</Name>" & vbCrLf
      strXML = strXML & String(2, vbTab) & "<FileName>" & vbComponent.Name & arrTypes(CInt(vbComponent.Type)) & "</FileName>" & vbCrLf
      strXML = strXML & String(2, vbTab) & "<Version>" & strVersion & "</Version>" & vbCrLf
      strXML = strXML & String(2, vbTab) & "<Type>" & vbComponent.Type & "</Type>" & vbCrLf
      strDirectory = Replace(vbComponent.Name, RegEx(vbComponent.Name, "_frm|_bas|_cls"), "")
      strXML = strXML & String(2, vbTab) & "<Directory>" & Replace(SetDirectory(CStr(vbComponent.Name)), "\", "") & "</Directory>" & vbCrLf
      strXML = strXML & String(1, vbTab) & "</Module>" & vbCrLf
    End If
next_vbComponent:
  Next vbComponent
  strXML = strXML & "</Modules>" & vbCrLf
  
  'write to the file
  Set oStream = CreateObject("ADODB.Stream")
  oStream.Type = 2 'adTypeText
  oStream.Charset = "utf-8"
  strFileName = strRepo & "CurrentVersions.xml"
  oStream.Open
  oStream.WriteText strXML
  oStream.SaveToFile strFileName, 2 'adSaveCreateOverWrite
  oStream.Close
  Set oStream = Nothing
  
  'stage the updated xml for next git commit/push
  git "add", strFileName

  MsgBox "CurrentVersions.xml created and staged." & vbCrLf & vbCrLf & "(Don't forget to push!)", vbInformation + vbOKOnly, "Complete"

exit_here:
  On Error Resume Next
  Set arrTypes = Nothing
  Set vbComponent = Nothing
  If oStream.State <> adStateClosed Then oStream.Close
  Set oStream = Nothing
  Exit Sub
  
err_here:
  Call HandleErr("cptAdmin_bas", "CreateCurrentVersionXML", err)
  Resume exit_here

End Sub

Sub Document()
'objects
Dim vbComponent As vbComponent
Dim xlApp As Object, Workbook As Object, Worksheet As Object
'strings
Dim strModule As String
Dim strProcName As String
'longs
Dim lngSLOC As Long
Dim lngLines As Long
Dim lngLine As Long
Dim lngRow As Long
Dim lngCountDecl As Long
'integers
'booleans
'variants
Dim arrHeader As Variant
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'get excel
  Set xlApp = CreateObject("Excel.Application")
  xlApp.Visible = True
  Set Workbook = xlApp.Workbooks.Add
  Set Worksheet = Workbook.Sheets(1)
  
  xlApp.ActiveWindow.Zoom = 85
  Worksheet.[A2].Select
  xlApp.ActiveWindow.FreezePanes = True
  
  'set the header
  arrHeader = Array("Ribbon Group", "Module", "SLOC", "Procedure", "SLOC", "Directory", "HelpDoc", "Author")
  Worksheet.Range(Worksheet.[A1], Worksheet.[A1].Offset(0, UBound(arrHeader))) = arrHeader
  Worksheet.Columns.AutoFit
  
  lngRow = 2
  
  For Each vbComponent In ThisProject.VBProject.VBComponents
    strModule = vbComponent.Name
    Debug.Print "working on " & strModule & "..."
    If strModule = "ThisProject" Or Left(strModule, 3) = "cpt" Then
      With vbComponent.CodeModule
        lngCountDecl = .CountOfDeclarationLines
        lngLines = .CountOfLines
        Worksheet.Cells(lngRow, 2) = .Name
        Worksheet.Cells(lngRow, 3) = .CountOfLines
        strProcName = .ProcOfLine(lngCountDecl + 1, 0) '0 = vbext_pk_Proc
        Worksheet.Cells(lngRow, 4) = strProcName
        Worksheet.Cells(lngRow, 5) = .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
        lngSLOC = lngSLOC + .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
        Worksheet.Columns.AutoFit
        For lngLine = lngCountDecl + 1 To lngLines
          If .ProcOfLine(lngLine, 0) <> strProcName Then '0 = vbext_pk_Proc
            strProcName = .ProcOfLine(lngLine, 0) '0 = vbext_pk_Proc
            lngRow = lngRow + 1
            Worksheet.Cells(lngRow, 2) = strModule
            Worksheet.Cells(lngRow, 4) = strProcName
            Worksheet.Cells(lngRow, 5) = .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
            lngSLOC = lngSLOC + .ProcCountLines(strProcName, 0) '0 = vbext_pk_Proc
            Worksheet.Columns.AutoFit
            If lngRow > 10 Then xlApp.ActiveWindow.ScrollRow = lngRow - 10
          End If
        Next
      End With
      lngRow = lngRow + 1
      If lngRow > 10 Then xlApp.ActiveWindow.ScrollRow = lngRow - 10
    End If
  Next vbComponent
  
  xlApp.ActiveWindow.ScrollRow = 2
  
  MsgBox "Documented." & vbCrLf & vbCrLf & "(" & Format(lngSLOC, "#,##0") & " SLOC)", vbInformation + vbOKOnly, "Documenter"
  
exit_here:
  On Error Resume Next
  Set vbComponent = Nothing
  Set xlApp = Nothing
  Set Workbook = Nothing
  Set Worksheet = Nothing
  Set xlApp = Nothing
  Exit Sub
err_here:
  Call HandleErr("cptAdmin_bas", "Document", err)
  Resume exit_here
End Sub

Sub CheckAllVersions()
Dim vbComponent As vbComponent

  For Each vbComponent In ThisProject.VBProject.VBComponents
    If Left(vbComponent.Name, 3) = "cpt" Then
      Debug.Print vbComponent.Name & ": " & vbComponent.CodeModule.Lines(1, 1)
    End If
  Next vbComponent
  Set vbComponent = Nothing
  
End Sub

Function SetDirectory(strComponentName As String) As String
'strings
Dim strDirectory As String
      
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  
  'remove the prefix
  strDirectory = Replace(strComponentName, "cpt", "")
  'remove the suffix
  If InStr(strDirectory, "_") > 0 Then strDirectory = Left(strDirectory, InStr(strDirectory, "_") - 1)
  Select Case strDirectory
    'Setup
    Case "Setup"
      strDirectory = ""
    'Core
    Case "BrowseFolder"
      strDirectory = "Core"
    Case "Logo"
      strDirectory = "Core"
    Case "Upgrades"
      strDirectory = "Core"
    Case "ThisProject"
      strDirectory = "Core"
    Case "Events"
      strDirectory = "Core"
    'count
    Case "CountTasks"
      strDirectory = "Count"
    'Integration
    Case "IMSCobraExport"
      strDirectory = "Integration"
    'TextTools
    Case "DynamicFilter"
      strDirectory = "TextTools"
    'Status
    Case "SmartDur"
      strDirectory = "Status"
    Case "StatusSheet"
      strDirectory = "Status"
    'Trace
    Case "CriticalPath"
      strDirectory = "Trace"
    Case "CriticalPathTools"
      strDirectory = "Trace"
    Case Else
      
        
  End Select
  
  SetDirectory = strDirectory & "\"
  
exit_here:
  On Error Resume Next

  Exit Function
err_here:
  Call HandleErr("cptAdmin_bas", "SetDirectory()", err)
  Resume exit_here

End Function
