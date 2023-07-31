VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCalendarExceptions_frm 
   Caption         =   "Calendar Details"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   OleObjectBlob   =   "cptCalendarExceptions_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptCalendarExceptions_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0.7</cpt_version>
Option Explicit

Private Sub cmdExport_Click()
  cptSaveSetting "CalendarDetails", "optDetailed", IIf(Me.optDetailed, 1, 0)
  Call cptExportCalendarExceptionsMain(Me.optDetailed)
End Sub
