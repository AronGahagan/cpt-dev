VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptResourceDemand_frm
   Caption         =   "Export Resource Demand"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9465
   OleObjectBlob   =   "cptResourceDemand_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptResourceDemand_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.5</cpt_version>
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0
Private Const adVarChar As Long = 200

Private Sub cmdAdd_Click()
Dim lgField As Long, lgExport As Long, lgExists As Long
Dim blnExists As Boolean

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lgField = 0 To Me.lboFields.ListCount - 1
    If Me.lboFields.Selected(lgField) Then
      'ensure doesn't already exist
      blnExists = False
      For lgExists = 0 To Me.lboExport.ListCount - 1
        If Me.lboExport.List(lgExists, 0) = Me.lboFields.List(lgField) Then
          GoTo next_item
        End If
      Next lgExists
      Me.lboExport.AddItem
      lgExport = Me.lboExport.ListCount - 1
      Me.lboExport.List(lgExport, 0) = Me.lboFields.List(lgField, 0)
      Me.lboExport.List(lgExport, 1) = Me.lboFields.List(lgField, 1)
    End If
next_item:
  Next lgField

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdAdd_Click", err, Erl)
  Resume exit_here

End Sub

Private Sub cmdCancel_Click()
Dim strFileName As String

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  If Dir(strFileName) <> vbNullString Then Kill strFileName
  Unload Me

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdCancel_Click", err, Erl)
  Resume exit_here

End Sub

Private Sub cmdExport_Click()
  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Call cptExportResourceDemand

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdExport_Click", err, Erl)
  Resume exit_here

End Sub

Private Sub cmdRemove_Click()
Dim lgExport As Long

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  For lgExport = Me.lboExport.ListCount - 1 To 0 Step -1
    If Me.lboExport.Selected(lgExport) Then
      Me.lboExport.RemoveItem lgExport
    End If
  Next lgExport

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "cmdRemove_Click", err, Erl)
  Resume exit_here
End Sub

Private Sub lblURL_Click()

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink "http://www.ClearPlanConsulting.com"

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "lblURL_Click", err, Erl)
  Resume exit_here

End Sub

Private Sub stxtSearch_Change()
'objects
'strings
Dim strFileName As String
'longs
Dim lngItem As Long
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  Me.lboFields.Clear

  strFileName = Environ("tmp") & "\cpt-resource-demand-search.adtg"
  With CreateObject("ADODB.Recordset")
    .Open strFileName
    If Len(Me.stxtSearch.Text) > 0 Then
      .Filter = "[Custom Field Name] LIKE '*" & cptRemoveIllegalCharacters(Me.stxtSearch.Text) & "*'"
    Else
      .Filter = 0
    End If
    If .RecordCount > 0 Then .MoveFirst
    lngItem = 0
    Do While Not .EOF
      Me.lboFields.AddItem
      Me.lboFields.List(lngItem, 0) = .Fields(0)
      Me.lboFields.List(lngItem, 1) = .Fields(1)
      .MoveNext
      lngItem = lngItem + 1
    Loop
    .Close
    Me.lblStatus.Caption = lngItem & " record" & IIf(lngItem = 1, "", "s") & " found."
  End With


exit_here:
  On Error Resume Next
  Exit Sub
err_here:
  Call cptHandleErr("cptResourceDemand_frm", "stxtSearch_Change", err, Erl)
  Resume exit_here

End Sub
