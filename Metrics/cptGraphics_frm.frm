VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptGraphics_frm 
   Caption         =   "Graphics"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "cptGraphics_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptGraphics_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboMetric_Change()
  Call cptGetChart(Me.cboMetric.Value)
End Sub
