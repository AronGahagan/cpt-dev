VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptGraphics_frm 
   Caption         =   "Metrics Quick Look"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420
   OleObjectBlob   =   "cptGraphics_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptGraphics_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLN_TRAP_ERRORS As Boolean = True
'If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

Private Sub cboMetric_Change()

  Select Case Me.cboMetric
    Case "Bow Wave"
      cptGraphics_frm.imgGraph.Picture = LoadPicture(cptDir & "\metrics\bow_wave.jpg", 481, 288)
      
    Case "Current Execution Index (CEI)"
    
    Case "DCMA14.01 Logic"
    
    Case "DCMA14.02 Leads"
    
    Case "DCMA14.03 Lags"
    
    Case "DCMA14.04 Relationship Types"
      cptGraphics_frm.imgGraph.Picture = LoadPicture(cptDir & "\metrics\dcma14-04.jpg", 481, 288)
      
    Case "DCMA14.05 Hard Constraints"
    
    Case "DCMA14.06 High Float"
    
    Case "DCMA14.07 Negative Float"
    
    Case "DCMA14.08 High Duration"
    
    Case "DCMA14.09 Invalid Dates"
    
    Case "DCMA14.10 Resources"
    
    Case "DCMA14.11 Missed Tasks"
    
    Case "DCMA14.12 Critical Path Test"
    
    Case "DCMA14.13 Critical Path Length Index (CPLI)"
    
    Case "DCMA14.14 Baseline Execution Index (BEI)"
    
    Case "-----------------------"
      Me.cboMetric.Value = "PLEASE CHOOSE A METRIC:"
      Me.imgGraph.Picture = LoadPicture("")
      
    Case Else
      Me.imgGraph.Picture = LoadPicture("")
      
  End Select
  
End Sub

Private Sub cmdAnalyze_Click()
  cptGetCharts
End Sub

Private Sub cmdDone_Click()
  Unload Me
End Sub
