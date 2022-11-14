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
  Dim blnUpdateView As Boolean
  If Me.lboMetrics.ListIndex = -1 Then Exit Sub
  
  Dim strMetric As String
  Dim strTitle As String
  Dim strTarget As String
  Dim strScore As String
  Dim lngX As Long
  Dim lngY As Long
  Dim dblScore As Double
  
  strMetric = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0)
  strTitle = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 1)
  strTarget = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 2)
  lngX = CLng(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 3))
  If Not IsNull(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4)) Then
    lngY = CLng(Me.lboMetrics.List(Me.lboMetrics.ListIndex, 4))
  Else
    lngY = 0
  End If
  strScore = Me.lboMetrics.List(Me.lboMetrics.ListIndex, 5)
  strDescription = strMetric & vbCrLf
  strDescription = strDescription & strTitle & vbCrLf & vbCrLf
  strDescription = strDescription & "TARGET: " & strTarget & vbCrLf
  strDescription = strDescription & "X: " & lngX & vbCrLf
  strDescription = strDescription & "Y: " & lngY & vbCrLf
  
  Select Case strMetric
    Case "06A208a"
      strDescription = strDescription & "SCORE: " & strScore
    Case "06A506b"
      strDescription = strDescription & "SCORE: " & strScore
    Case Else
      strDescription = strDescription & "SCORE: " & lngX & "/" & lngY & " = " & strScore
  End Select
  
  Me.txtTitle.Value = strDescription
  blnUpdateView = True
  If blnUpdateView Then
    cptDECM_UPDATE_VIEW Me.lboMetrics.List(Me.lboMetrics.ListIndex, 0), Me.lboMetrics.List(Me.lboMetrics.ListIndex, 8)
  End If
End Sub
