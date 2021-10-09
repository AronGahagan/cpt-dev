VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSettings_frm 
   Caption         =   "cpt Settings"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "cptSettings_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSettings_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = False
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdDone_Click()
  Unload Me
End Sub

Private Sub cmdEdit_Click()
Dim strMsg As String
  strMsg = "...unless you *really* know what you're doing." & vbCrLf & vbCrLf
  strMsg = strMsg & "Contact cpt@ClearPlanConsulting.com if you need help." & vbCrLf & vbCrLf
  strMsg = strMsg & "Do you still wish to venture forth?"
  If MsgBox(strMsg, vbCritical + vbYesNo, "Do Not Attempt This...") = vbYes Then
    MsgBox "...you've been warned.", vbInformation + vbOKOnly, "OK"
    Shell "C:\Windows\notepad.exe '" & cptDir & "\settings\cpt-settings.ini" & "'", vbNormalFocus
  End If
End Sub

Private Sub cmdSetProgramAcronym_Click()
  'objects
  Dim oRecordset As ADODB.Recordset
  Dim oDocProp As DocumentProperty
  Dim oDocProps As DocumentProperties
  'strings
  Dim strFile As String
  Dim strOld As String, strNew As String
  'longs
  Dim lngUpdated As Long
  Dim lngResponse As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates

  Set oDocProps = ActiveProject.CustomDocumentProperties
  On Error Resume Next
  Set oDocProp = oDocProps("cptProgramAcronym")
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
  If oDocProp Is Nothing Then
    oDocProps.Add "cptProgramAcronym", False, msoPropertyTypeString, Me.txtProgramAcronym.Value, False
   Else
    strOld = oDocProp.Value
    strNew = Me.txtProgramAcronym
    oDocProp.Value = strNew
    lngResponse = MsgBox("Also replace all instances of '" & strOld & "' with '" & strNew & "' in stored program data and settings (CEI, Metrics, Marked, Data Dictionary)?" & vbCrLf & vbCrLf & "Hit Cancel to ignore all updates.", vbQuestion + vbYesNoCancel, "Confirm Overwrite")
    Select Case lngResponse
      Case vbYes
        Set oRecordset = CreateObject("ADODB.Recordset")
        '\settings\cpt-cei.adtg
        strFile = cptDir & "\settings\cpt-cei.adtg"
        lngUpdated = 0
        If Dir(strFile) <> vbNullString Then
          With oRecordset
            .Open strFile
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROJECT") = strOld Then
                .Fields("PROJECT") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFile, adPersistADTG
            .Close
          End With
          MsgBox Format(lngUpdated, "#,##0") & " record(s) updated in cpt-cei.adtg.", vbInformation + vbOKOnly, "CEI Updated"
        End If
        '\settings\cpt-metrics.adtg
        strFile = cptDir & "\settings\cpt-metrics.adtg"
        lngUpdated = 0
        If Dir(strFile) <> vbNullString Then
          With oRecordset
            .Open strFile
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROGRAM") = strOld Then
                .Fields("PROGRAM") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFile, adPersistADTG
            .Close
          End With
          MsgBox Format(lngUpdated, "#,##0") & " record(s) updated in cpt-metrics.adtg.", vbInformation + vbOKOnly, "Metrics Updated"
        End If
        '\cpt-marked.adtg
        strFile = cptDir & "\cpt-marked.adtg"
        lngUpdated = 0
        If Dir(strFile) <> vbNullString Then
          With oRecordset
            .Open strFile
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROJECT_ID") = strOld Then
                .Fields("PROJECT_ID") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFile, adPersistADTG
            .Close
          End With
          MsgBox Format(lngUpdated, "#,##0") & " record(s) updated in cpt-marked.adtg.", vbInformation + vbOKOnly, "Marked Updated"
        End If
        'settings\cpt-data-dictionary.adtg
        strFile = cptDir & "\settings\cpt-data-dictionary.adtg"
        lngUpdated = 0
        If Dir(strFile) <> vbNullString Then
          With oRecordset
            .Open strFile
            .MoveFirst
            Do While Not .EOF
              If .Fields("PROJECT_NAME") = strOld Then
                .Fields("PROJECT_NAME") = strNew
                lngUpdated = lngUpdated + 1
              End If
              .MoveNext
            Loop
            .Save strFile, adPersistADTG
            .Close
          End With
          MsgBox Format(lngUpdated, "#,##0") & " record(s) updated in cpt-data-dictionary.adtg.", vbInformation + vbOKOnly, "Data Dictionary Updated"
        End If
        
      Case vbNo
        If MsgBox("CEI, Metrics, Marked tasks, and Data Dictionary entries associated with '" & strOld & "' will be disconnected from this project file. Are you sure you wish to proceed?", vbQuestion + vbYesNo) = vbNo Then
          oDocProp.Value = strOld
          Me.txtProgramAcronym.Value = strOld
        End If
      Case vbCancel
        oDocProp.Value = strOld
        Me.txtProgramAcronym.Value = strOld
    End Select
  End If
  
exit_here:
  On Error Resume Next
  If oRecordset.State = 1 Then oRecordset.Close
  Set oRecordset = Nothing
  Set oDocProp = Nothing
  Set oDocProps = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSettings_frm", "cmdSetProgramAcronym_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lblURL_Click()
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptBackbone_frm", "lblURL_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub lboFeatures_AfterUpdate()
  'objects
  Dim oRecordset As ADODB.Recordset
  'strings
  'longs
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboSettings.Clear
  Set oRecordset = CreateObject("ADODB.Recordset")
  With oRecordset
    .Open cptDir & "\settings\cpt-settings.adtg"
    .Sort = "Setting"
    .MoveFirst
    Do While Not .EOF
      If .Fields(0) = Me.lboFeatures.Value And .Fields(1) <> "" Then
        Me.lboSettings.AddItem .Fields(1)
      End If
      .MoveNext
    Loop
    .Close
  End With
  
exit_here:
  On Error Resume Next
  Set oRecordset = Nothing

  Exit Sub
err_here:
  Call cptHandleErr("cptSettings_frm", "lboFeatures_AfterUpdate", Err, Erl)
  Resume exit_here
End Sub

Private Sub UserForm_Terminate()
  Dim strFile As String
  strFile = cptDir & "\settings\cpt-settings.adtg"
  If Dir(strFile) <> vbNullString Then Kill strFile
End Sub
