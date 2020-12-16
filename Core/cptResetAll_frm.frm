VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptResetAll_frm 
   Caption         =   "How would you like to Reset All?"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   OleObjectBlob   =   "cptResetAll_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptResetAll_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.0</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Sub cmdDoIt_Click()
  'objects
  Dim rstSettings As ADODB.Recordset
  'strings
  Dim strFile As String
  'longs
  Dim lngSettings As Long
  Dim lngOutlineLevel As Long
  Dim lngLevel As Long
  'integers
  'doubles
  'booleans
  'variants
  'dates
  
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  cptSpeed True

  strFile = cptDir & "\settings\cpt-reset-all.adtg"
  If Dir(strFile) <> vbNullString Then Kill strFile
  Set rstSettings = CreateObject("ADODB.Recordset")
  rstSettings.Fields.Append "SETTINGS", adInteger
  rstSettings.Fields.Append "OUTLINE_LEVEL", adInteger
  rstSettings.Open
  
  'capture bitwise value
  If Me.chkFilter Then
    FilterClear
    lngSettings = 1
  End If
  If Me.chkGroup Then
    GroupClear
    lngSettings = lngSettings + 2
  End If
  If Me.chkSummaries Then
    OptionsViewEx displaysummarytasks:=True
    lngSettings = lngSettings + 4
  End If
  If Me.optShowAllTasks Then
    OptionsViewEx displaysummarytasks:=True
    OutlineShowAllTasks
    If Not Me.chkSummaries Then
      OptionsViewEx displaysummarytasks:=False
    End If
    lngSettings = lngSettings + 8
  ElseIf Me.optOutlineLevel Then
    OptionsViewEx displaysummarytasks:=True
    OutlineShowAllTasks
    lngOutlineLevel = Me.cboOutlineLevel
    OutlineShowTasks pjTaskOutlineShowLevelMax
    For lngLevel = 20 To lngOutlineLevel Step -1
      OutlineShowTasks lngLevel
    Next lngLevel
  End If
  If Me.chkSort Then
    Sort "ID"
    lngSettings = lngSettings + 16
  End If
  If Me.chkActiveOnly Then
    SetAutoFilter "Active", pjAutoFilterFlagYes
    lngSettings = lngSettings + 32
  End If
  If Me.chkIndent Then
    OptionsViewEx displaynameindent:=True
    lngSettings = lngSettings + 64
  End If
  If Me.chkOutlineSymbols Then
    OptionsViewEx displayoutlinesymbols:=True
    lngSettings = lngSettings + 128
  End If
  rstSettings.AddNew Array(0, 1), Array(lngSettings, lngOutlineLevel)

  rstSettings.Save strFile, adPersistADTG

exit_here:
  On Error Resume Next
  If rstSettings.State Then rstSettings.Close
  Set rstSettings = Nothing
  cptSpeed False
  Unload Me
  Exit Sub
err_here:
  Call cptHandleErr("cptResetAll_frm", "cmdDoIt_Click", Err, Erl)
  Resume exit_here
End Sub

Private Sub optOutlineLevel_Click()
  Me.cboOutlineLevel.Enabled = True
End Sub

Private Sub optShowAllTasks_Click()
  Me.cboOutlineLevel.Enabled = False
End Sub
