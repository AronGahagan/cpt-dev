VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptCritPathFields_frm 
   Caption         =   "cpt Driving Paths"
   ClientHeight    =   2412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4068
   OleObjectBlob   =   "cptCritPathFields_frm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "cptCritPathFields_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v2.9.0</cpt_version>

Private Sub RunBtn_Click()
        
    If PathField_Combobox.Text = "" Or GroupField_Combobox.Text = "" Then
        MsgBox "Please complete the required field mapping."
        Exit Sub
    End If
    
    StoreCustomFieldName "Driving Paths", "CP Driving Paths", FieldNameToFieldConstant(PathField_Combobox.Text)
    StoreCustomFieldName "Driving Path Group", "CP Driving Path Group ID", FieldNameToFieldConstant(GroupField_Combobox.Text)
    
    'Store Field Names
    CustomFieldRename FieldID:=FieldNameToFieldConstant(PathField_Combobox.Text), newname:="CP Driving Paths"
    CustomFieldRename FieldID:=FieldNameToFieldConstant(GroupField_Combobox.Text), newname:="CP Driving Path Group ID"
    
    Me.Tag = "run"
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Dim drivingPathField As String
    Dim groupPathField As String

    drivingPathField = GetCustomFieldName("Driving Paths")
    groupPathField = GetCustomFieldName("Driving Path Group")
    
    DisplayUserCustomFields drivingPathField, groupPathField
    
End Sub

Private Sub DisplayUserCustomFields(ByVal drivingPathField As String, ByVal groupPathField As String)

    Dim nameTest As Double

    On Error GoTo MissingDrivingPathsField
    
    nameTest = FieldNameToFieldConstant(drivingPathField)
    PathField_Combobox.Value = drivingPathField
    
MissingDrivingPathsField:
    
    On Error GoTo MissingGroupField
    
    nameTest = FieldNameToFieldConstant(groupPathField)
    GroupField_Combobox.Value = groupPathField
    
MissingGroupField:

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    Me.Tag = "cancel"
    Me.Hide
  End If
End Sub


