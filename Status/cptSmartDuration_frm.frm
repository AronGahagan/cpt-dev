VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptSmartDuration_frm 
   Caption         =   "Smart Duration"
   ClientHeight    =   1185
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3345
   OleObjectBlob   =   "cptSmartDuration_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptSmartDuration_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.0</cpt_version>

Public dateError As Boolean
Public finDate As Date
Public startDate As Date

Private Sub cancelBtn_Click()

    Me.Tag = "Cancel"
    Me.hide
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Tag = "Cancel"
        Me.hide
    End If
End Sub

Private Sub okBtn_Click()

    If dateError = True Or SmartDatePicker.Text = "" Then
        MsgBox "Enter a valid Finish Date"
        Exit Sub
    End If
    
    finDate = CDate(Month(finDate) & "/" & Day(finDate) & "/" & Year(finDate) & " 5:00 PM")
    
    Me.Tag = "OK"
    Me.hide
    
End Sub

Private Sub SmartDatePicker_AfterUpdate()

    dateError = False
    If Not IsDate(SmartDatePicker.Text) Then
        MsgBox "Please enter a valid date."
        dateError = True
        Exit Sub
    End If
    finDate = Format(SmartDatePicker.Text, "MM/DD/YY")
    If finDate <= startDate Then
        MsgBox "Please enter a Finish Date that is greater than the Start Date"
        dateError = True
        Exit Sub
    End If
    
    SmartDatePicker.Text = finDate
    weekDayLbl = Format(finDate, "DDD")
    
End Sub
