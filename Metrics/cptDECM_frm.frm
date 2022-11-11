VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptDECM_frm 
   Caption         =   "DECM v5.0"
   ClientHeight    =   4680
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
Option Explicit

Private Sub cmdExport_Click()
  cptDECM_EXPORT
End Sub

Private Sub lboMetrics_AfterUpdate()
  Dim strDescription As String
  If IsNull(Me.lboMetrics.ListIndex) Then Exit Sub
  strDescription = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0) & vbCrLf
  strDescription = strDescription & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 1) & vbCrLf & vbCrLf
  strDescription = strDescription & "TARGET: " & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 2) & vbCrLf
  strDescription = strDescription & "X: " & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3) & vbCrLf
  strDescription = strDescription & "Y: " & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4) & vbCrLf
  strDescription = strDescription & "SCORE: " & Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)
  Me.txtTitle.Value = strDescription
End Sub
