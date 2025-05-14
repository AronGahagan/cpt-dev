VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAbout_frm 
   Caption         =   "The ClearPlan Toolbar - v1.7.4"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "cptAbout_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAbout_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.9.0</cpt_version>
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAbout_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub
