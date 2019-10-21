VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptFiscal_frm 
   Caption         =   "Fiscal Calendar"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   OleObjectBlob   =   "cptFiscal_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptFiscal_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExportExceptions_Click()
  Call cptExportCalendarExceptions
End Sub

Private Sub cmdImportExceptions_Click()
  Call cptImportCalendarExceptions
End Sub

Private Sub cmdTemplateExceptions_Click()
  Call cptExportExceptionsTemplate
End Sub
