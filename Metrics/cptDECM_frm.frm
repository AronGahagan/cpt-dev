VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDECM_frm 
   Caption         =   "DECM v6.0"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760.001
   OleObjectBlob   =   "cptDECM_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptDECM_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v0.0.4</cpt_version>
Option Explicit

Private Sub cmdDone_Click()
  Unload Me
  'then clean up after yourself
  Dim vFile As Variant
  Dim strFile As String
  For Each vFile In Split("Schema.ini,tasks.csv,assignments.csv,links.csv,wp-ims.csv,wp-ev.csv,wp-not-in-ims.csv,wp-not-in-ev.csv,10A302b-x.csv,10A303a-x.csv,fiscal.csv,cpt-cei.csv,06A506c-x.csv,06A504a.csv", ",")
    strFile = Environ("tmp") & "\" & vFile
    If Dir(strFile) <> vbNullString Then Kill strFile
  Next vFile
  cptResetAll
End Sub

Private Sub cmdExport_Click()
  cptDECM_EXPORT
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_frm", "lblURL_Click", Err, Erl)
  Resume exit_here

End Sub

Public Sub lboMetrics_AfterUpdate()
  'objects
  Dim oFile As Scripting.TextStream  'Object
  Dim oFSO As Scripting.FileSystemObject  'Object
  'strings
  Dim strMsg As String
  Dim strDir As String
  Dim strDescription As String
  Dim strMetric As String
  Dim strTitle As String
  Dim strTarget As String
  Dim strScore As String
  'longs
  Dim lngX As Long
  Dim lngY As Long
  'integers
  'doubles
  Dim dblScore As Double
  'booleans
  Dim blnUpdateView As Boolean
  'variants
  'dates
  
  If Me.lboMetrics.ListIndex = -1 Then Exit Sub
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  strMetric = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0)
  strTitle = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 1)
  strTarget = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 2)
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3)) Then
    lngX = CLng(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3))
  Else
    lngX = 0
  End If
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4)) Then 'todo: not all metrics have Y
    lngY = CLng(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4))
  Else
    lngY = 0
  End If
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)) Then
    strScore = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)
  Else
    strScore = "-"
  End If
  strDescription = strMetric & vbCrLf
  strDescription = strDescription & strTitle & vbCrLf & vbCrLf
  strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
  strDescription = strDescription & "X: " & lngX & vbCrLf
  strDescription = strDescription & "Y: " & lngY & vbCrLf
  
  Select Case strMetric
    Case "06A101a"
      strDescription = "needed: wp-ims.csv [+]" & vbCrLf
      strDescription = strDescription & "needed: wp-ev.csv  <?>" & vbCrLf
      Me.txtTitle.Value = strDescription
      If MsgBox("Has the EV Analyst sent you the list of discrete, incomplete WPs in the EV Tool?", vbQuestion + vbYesNo, "06A101a - WP Mismatches") = vbNo Then
        MsgBox "Please send the following query to your EV Analyst...", vbOKOnly + vbInformation, "Data Needed"
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        strDir = Environ("tmp")
        Set oFile = oFSO.CreateTextFile(strDir & "\wp-ev.sql.txt", True)
        strMsg = "Hi [person]," & vbCrLf & vbCrLf
        strMsg = strMsg & "I'm running DECM metric 06A101a which compares the list of discrete, incomplete WPs in the IMS vs what's in the EV Tool. " & vbCrLf
        strMsg = strMsg & "Could you please provide the list of discrete, incomplete WPs currently in the EV Tool?" & vbCrLf & vbCrLf
        strMsg = strMsg & "An example query for COBRA would be:" & vbCrLf
        strMsg = strMsg & String(25, "-") & vbCrLf
        
        strMsg = strMsg & "DECLARE @MyProj VARCHAR(MAX) " & vbCrLf
        strMsg = strMsg & "SET @MyProj=inputbox('Project Name:') " & vbCrLf
        strMsg = strMsg & "SELECT DISTINCT WP " & vbCrLf
        strMsg = strMsg & "FROM CAWP " & vbCrLf
        strMsg = strMsg & "WHERE PROGRAM=@MyProj " & vbCrLf
        strMsg = strMsg & "AND WP<>'' " & vbCrLf
        strMsg = strMsg & "AND PMT NOT IN ('A','J','M') " & vbCrLf
        strMsg = strMsg & "AND BCWP<(BAC-100) " & vbCrLf
        
        strMsg = strMsg & String(25, "-") & vbCrLf & vbCrLf
        strMsg = strMsg & "I appreciate your assistance. Please let me know if you have any questions."
        oFile.Write strMsg
        oFile.Close
        Shell "C:\Windows\notepad.exe '" & strDir & "\wp-ev.sql.txt" & "'", vbNormalFocus
        GoTo exit_here
      Else
        Me.txtTitle.Value = Me.txtTitle.Text & vbCrLf & "please paste data here (w/o headers):" & vbCrLf
        Me.txtTitle.SetFocus
        Me.txtTitle.SelStart = 0
        Me.txtTitle.CurLine = Me.txtTitle.LineCount - 2
        Me.txtTitle.SelLength = 65535
        Me.txtTitle.CurLine = Me.txtTitle.LineCount - 3
        Me.txtTitle.SelLength = 65535
        
        GoTo exit_here
      End If
    
    Case "06A208a"
      strDescription = strDescription & "SCORE: " & strScore
    Case "06A504a"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "...requires CPT > Status > Capture Week, two periods"
    Case "06A504b"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "...requires CPT > Status > Capture Week, two periods"
    Case "06A506b"
      strDescription = strDescription & "SCORE: " & strScore
    Case "06A506c"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "...requires CPT > Status > Capture Week, two periods"
    Case "06A212a"
      strDescription = strDescription & vbCrLf & "...pairs exported to Excel" & vbCrLf & "...select to filter"
    Case "10A103a"
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
      If lngX > 0 Then strDescription = strDescription & vbCrLf & vbCrLf & "...details exported to Excel"
    Case "06I201a"
      strDescription = strDescription & "SCORE: " & strScore
      strDescription = strDescription & vbCrLf & vbCrLf & "Task Name contains 'SVT' and has resource assignments"
    Case Else
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
  End Select
  
  Me.txtTitle.Value = strDescription
  blnUpdateView = True 'todo: make this a checkbox on the DECM form
  If blnUpdateView And Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 8)) Then
    cptDECM_UPDATE_VIEW Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0), Me.lboMetrics.List(Me.lboMetrics.ListIndex, 8)
  End If
  
exit_here:
  On Error Resume Next
  Set oFile = Nothing
  Set oFSO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_frm", "lboMetrics_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub txtTitle_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oFile As Scripting.TextStream
  Dim oFSO As Scripting.FileSystemObject
  'strings
  Dim strDescription As String
  Dim strPass As String
  Dim strFail As String
  Dim strScore As String
  Dim strTarget As String
  Dim strTitle As String
  Dim strMetric As String
  Dim strDir As String
  Dim strCon As String
  Dim strSQL As String
  'longs
  Dim lngX As Long
  Dim lngY As Long
  Dim lngRecord As Long
  'integers
  'doubles
  Dim dblScore As Double
  'booleans
  'variants
  Dim vData As Variant
  'dates
  
  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0
  
  Me.txtTitle.Text = Me.txtTitle.Text & "validating..." & vbCrLf
  vData = Split(Data.GetText, vbCrLf)
  
  strDir = Environ("tmp")
  strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & strDir & "';Extended Properties='text;HDR=Yes;FMT=Delimited';"
  Set oRecordset = CreateObject("ADODB.Recordset")

  'populate lboFilter
  If UBound(vData) > 1 Then 'user pasted a column of data
    Cancel = True 'cancel paste operation
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = oFSO.CreateTextFile(strDir & "\wp-ev.csv", True)
    oFile.Write Join(vData, vbCrLf)
    oFile.Close
    'ensure distinct
    strSQL = "SELECT DISTINCT WP FROM [wp-ev.csv]"
    oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
    lngY = oRecordset.RecordCount
    cptDECM_frm.lboMetrics.List(cptDECM_frm.lboMetrics.ListCount - 1, 4) = lngY
    Set oFile = oFSO.CreateTextFile(strDir & "\wp-ev-distinct.csv", True)
    oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
    oFile.Close
    oRecordset.Close
    Kill strDir & "\wp-ev.csv"
    Name strDir & "\wp-ev-distinct.csv" As strDir & "\wp-ev.csv"
  Else
    'user pasted nothing
  End If
  
  Me.txtTitle.Text = Replace(Me.txtTitle.Text, "validating...", "validating...ok")
  Me.txtTitle.Text = Me.txtTitle.Text & lngY & " WPs loaded..." & vbCrLf
  Me.txtTitle.Text = Replace(Me.txtTitle.Text, "<?>", "[+]")
  Me.txtTitle.Text = Me.txtTitle.Text & "analyzing..." & vbCrLf
  
  'get count of WPs in IMS but not in EV
  strSQL = "SELECT wp_ims.WP,wp_ev.WP "
  strSQL = strSQL & "FROM [wp-ims.csv] AS wp_ims "
  strSQL = strSQL & "LEFT JOIN [wp-ev.csv] AS wp_ev ON wp_ev.WP=wp_ims.WP "
  strSQL = strSQL & "WHERE wp_ev.WP IS NULL"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  lngX = oRecordset.RecordCount
  Me.txtTitle.Text = Me.txtTitle.Text & lngX & " in IMS, not in EV" & vbCrLf
  Set oFile = oFSO.CreateTextFile(strDir & "\wp-not-in-ev.csv", True)
  oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  oFile.Close
  oRecordset.Close
  
  'get count of WPs in EV but not in IMS
  strSQL = "SELECT wp_ev.WP,wp_ims.WP "
  strSQL = strSQL & "FROM [wp-ev.csv] AS wp_ev "
  strSQL = strSQL & "LEFT JOIN [wp-ims.csv] AS wp_ims ON wp_ims.WP=wp_ev.WP "
  strSQL = strSQL & "WHERE wp_ims.WP IS NULL"
  oRecordset.Open strSQL, strCon, adOpenKeyset, adLockReadOnly
  Me.txtTitle.Text = Me.txtTitle.Text & oRecordset.RecordCount & " in EV, not in IMS" & vbCrLf
  lngX = lngX + oRecordset.RecordCount
  Set oFile = oFSO.CreateTextFile(strDir & "\wp-not-in-ims.csv", True)
  oFile.Write oRecordset.GetString(adClipString, , ",", vbCrLf, vbNullString)
  oFile.Close
  oRecordset.Close
  Me.txtTitle.Text = Me.txtTitle.Text & lngX & " total"
  
  'update the listbox
  Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3) = lngX
  dblScore = Round(lngX / lngY, 2)
  Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5) = Format(dblScore, "0%")
  strPass = "[+]"
  strFail = "<!>"
  If dblScore = 0 Then
    Me.lboMetrics.List(Me.lboMetrics.ListIndex, 6) = strPass
  Else
    Me.lboMetrics.List(Me.lboMetrics.ListIndex, 6) = strFail
  End If

  'update the description
  strMetric = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0)
  strTitle = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 1)
  strTarget = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 2)
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)) Then
    strScore = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)
  Else
    strScore = "-"
  End If
  strDescription = strMetric & vbCrLf
  strDescription = strDescription & strTitle & vbCrLf & vbCrLf
  strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
  strDescription = strDescription & "X: " & lngX & vbCrLf
  strDescription = strDescription & "Y: " & lngY & vbCrLf
  strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
  Me.txtTitle.Text = strDescription
  
  'todo: user can re-paste data
    
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing
  Set oFile = Nothing
  Set oFSO = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptDECM_frm", "txtTitle_BeforeDropOrPaste", Err, Erl)
  Resume exit_here
End Sub
