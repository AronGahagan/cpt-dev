VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptIMSCobraExport_frm 
   Caption         =   "IMS Export Utility v3.1.1"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "cptIMSCobraExport_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptIMSCobraExport_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bcrBox_Change()

    If checkDuplicate(bcrBox) = True Then
        MsgBox "Please select a unique IMS Field."
        bcrBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(bcrBox.Value) = False And bcrBox.Value <> "" And bcrBox.Value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        bcrBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fBCR").Value = Me.bcrBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fBCR", False, msoPropertyTypeString, Me.bcrBox.Value
    Resume PropFound
End Sub

Private Function checkDuplicate(ByVal cBoxTest As ComboBox) As Boolean

    If cBoxTest.Value = "<None>" Or cBoxTest.Value = "" Then
    
        checkDuplicate = False
        Exit Function
    
    End If

    Dim cBoxOther As ComboBox
    Dim formObj As control
    
    For Each formObj In Me.TabButtons.Pages(1).Controls
    
        If TypeName(formObj) = "ComboBox" Then
        
            Set cBoxOther = formObj
            
            If cBoxOther.Name <> cBoxTest.Name Then
            
                If cBoxOther.Value = cBoxTest.Value Then
                
                    checkDuplicate = True
                    Exit Function
                
                End If
            
            End If
        
        End If
    
    Next formObj
    
    checkDuplicate = False

End Function

Private Sub BcrBtn_Change()

    If BcrBtn = True Then
        Me.BCR_ID_TextBox.Enabled = True
        Exit Sub
    Else
        Me.BCR_ID_TextBox.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub BcrBtn_Click()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing

PropFound:

    If docProps("fBCR").Value <> "<None>" Then
        Exit Sub
    End If
    
PropMissing:

    MsgBox "Please map a BCR Field before using the BCR Export function."
    Me.BcrBtn = False
    Me.TotalProjBtn = True
    Me.BCR_ID_TextBox.Enabled = False
    Exit Sub
End Sub

Private Sub BCWS_Checkbox_Change()

    If BCWS_Checkbox.Value = True Then
        Me.TotalProjBtn.Enabled = True
        Me.BcrBtn.Enabled = True
        If BcrBtn = True Then
            BCR_ID_TextBox.Enabled = True
        End If
        Me.exportDescCheckBox.Enabled = True
        Me.exportTPhaseCheckBox.Enabled = True
    Else
        Me.TotalProjBtn.Enabled = False
        Me.BcrBtn.Enabled = False
        BCR_ID_TextBox.Enabled = False
        Me.exportDescCheckBox.Enabled = False
        If Me.ETC_Checkbox.Value = False Then
            Me.exportTPhaseCheckBox.Enabled = False
        End If
    End If

End Sub


Private Sub caID1Box_Change()

    If checkDuplicate(caID1Box) = True Then
        MsgBox "Please select a unique IMS Field."
        caID1Box.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(caID1Box.Value) = False And caID1Box.Value <> "" Then
        MsgBox "Please select a valid IMS Field."
        caID1Box.Value = ""
        CAID1TxtBox.Value = ""
        Exit Sub
    End If

    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID1").Value = Me.caID1Box.Value
    If Me.Tag = "Loaded" And Me.CAID1TxtBox.Value = "" Then Me.CAID1TxtBox.Value = Me.caID1Box.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID1", False, msoPropertyTypeString, Me.caID1Box.Value
    Resume PropFound

End Sub

Private Sub CAID1TxtBox_Change()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID1t").Value = Me.CAID1TxtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID1t", False, msoPropertyTypeString, Me.CAID1TxtBox.Value
    Resume PropFound
End Sub

Private Sub CAID1TxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID1t").Value = Me.CAID1TxtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID1t", False, msoPropertyTypeString, Me.CAID1TxtBox.Value
    Resume PropFound
End Sub

Private Sub caID2Box_Change()

    If checkDuplicate(caID2Box) = True Then
        MsgBox "Please select a unique IMS Field."
        caID2Box.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(caID2Box.Value) = False And caID2Box.Value <> "" And caID2Box.Value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        caID2Box.Value = ""
        CAID2TxtBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID2").Value = Me.caID2Box.Value
    If Me.Tag = "Loaded" And Me.CAID2TxtBox.Value = "" Then Me.CAID2TxtBox.Value = Me.caID2Box.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    If Me.caID2Box.Value = "<None>" Then
        Me.CAID2TxtBox.Enabled = False
        Me.CAID2TxtBox.Visible = False
    Else
        Me.CAID2TxtBox.Enabled = True
        Me.CAID2TxtBox.Visible = True
        If Me.Tag = "Loaded" And Me.CAID2TxtBox.Value = "" Then Me.CAID2TxtBox.Value = Me.caID2Box.Value
    End If
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID2", False, msoPropertyTypeString, Me.caID2Box.Value
    Resume PropFound
End Sub

Private Sub CAID2TxtBox_Change()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID2t").Value = Me.CAID2TxtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID2t", False, msoPropertyTypeString, Me.CAID2TxtBox.Value
    Resume PropFound
End Sub

Private Sub CAID2TxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID2t").Value = Me.CAID2TxtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID2t", False, msoPropertyTypeString, Me.CAID2TxtBox.Value
    Resume PropFound
End Sub

Private Sub caID3Box_Change()

    If checkDuplicate(caID3Box) = True Then
        MsgBox "Please select a unique IMS Field."
        caID3Box.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(caID3Box.Value) = False And caID3Box.Value <> "" And caID3Box.Value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        caID3Box.Value = ""
        CAID3TxtBox = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID3").Value = Me.caID3Box.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    If Me.caID3Box.Value = "<None>" Then
        Me.CAID3TxtBox.Enabled = False
        Me.CAID3TxtBox.Visible = False
    Else
        Me.CAID3TxtBox.Enabled = True
        Me.CAID3TxtBox.Visible = True
        If Me.Tag = "Loaded" And Me.CAID3TxtBox.Value = "" Then Me.CAID3TxtBox.Value = Me.caID3Box.Value
    End If
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID3", False, msoPropertyTypeString, Me.caID3Box.Value
    Resume PropFound
End Sub

Private Sub CAID3TxtBox_Change()
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID3t").Value = Me.CAID3TxtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID3t", False, msoPropertyTypeString, Me.CAID3TxtBox.Value
    Resume PropFound
End Sub

Private Sub CAID3TxtBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAID3t").Value = Me.CAID3TxtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAID3t", False, msoPropertyTypeString, Me.CAID3TxtBox.Value
    Resume PropFound
End Sub

Private Sub camBox_Change()

    If checkDuplicate(camBox) = True Then
        MsgBox "Please select a unique IMS Field."
        camBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(camBox.Value) = False And camBox.Value <> "" Then
        MsgBox "Please select a valid IMS Field."
        camBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fCAM").Value = Me.camBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fCAM", False, msoPropertyTypeString, Me.camBox.Value
    Resume PropFound
End Sub

Private Sub cancelBtn_Click()
    Me.Tag = "Cancel"
    Me.Hide
End Sub

Private Sub cptLinkLabel_Click()
    Application.OpenBrowser "http://www.ClearPlanConsulting.com"
End Sub

Private Sub CSVBtn_Change()

    If CSVBtn.Value = True Then
        Me.BCWS_Checkbox.Enabled = True
        Me.BCWP_Checkbox.Enabled = True
        Me.ETC_Checkbox.Enabled = True
        Me.ResExportCheckbox.Enabled = True
        If Me.ResExportCheckbox.Value = True Then
            Me.exportTPhaseCheckBox.Enabled = True
        Else
            Me.exportTPhaseCheckBox.Enabled = False
        End If
        If Me.BCWS_Checkbox = True Then
            Me.TotalProjBtn.Enabled = True
            Me.BcrBtn.Enabled = True
            Me.exportDescCheckBox.Enabled = True
            If Me.BcrBtn = True Then Me.BCR_ID_TextBox.Enabled = True
        End If
    Else
        Me.BCWS_Checkbox.Enabled = False
        Me.BCWP_Checkbox.Enabled = False
        Me.ETC_Checkbox.Enabled = False
        Me.TotalProjBtn.Enabled = False
        Me.ResExportCheckbox.Enabled = False
        Me.exportTPhaseCheckBox.Enabled = False
        Me.BcrBtn.Enabled = False
        Me.BCR_ID_TextBox.Enabled = False
        Me.exportDescCheckBox.Enabled = False
    End If
    
    If BCWS_Checkbox.Value = False And BCWP_Checkbox.Value = False And ETC_Checkbox.Value = False Then
    
        BCWS_Checkbox.Value = True
        Me.TotalProjBtn.Enabled = True
        Me.BcrBtn.Enabled = True
        If BcrBtn = True Then
            BCR_ID_TextBox.Enabled = True
        End If
    
    End If

End Sub


Private Sub ETC_Checkbox_Click()
    If Me.ETC_Checkbox = True Then
        Me.exportTPhaseCheckBox.Enabled = True
    Else
        If Me.BCWS_Checkbox = False Then
            Me.exportTPhaseCheckBox.Enabled = False
        End If
    End If
End Sub

Private Sub evtBox_Change()

    If checkDuplicate(evtBox) = True Then
        MsgBox "Please select a unique IMS Field."
        evtBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(evtBox.Value) = False And evtBox.Value <> "" Then
        MsgBox "Please select a valid IMS Field."
        evtBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fEVT").Value = Me.evtBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fEVT", False, msoPropertyTypeString, Me.evtBox.Value
    Resume PropFound
End Sub

Private Sub ExportBtn_Click()

    If CSVBtn.Value = True And BCWS_Checkbox.Value = False And BCWP_Checkbox.Value = False And ETC_Checkbox.Value = False Then
    
        MsgBox "You must select at least one CSV export file type."
        Exit Sub
        
    End If
    
    If BCR_ID_TextBox.Enabled = True Then
        If BCR_ID_TextBox.Value = "Enter BCR ID" Or BCR_ID_TextBox.Value = "" Then
            MsgBox "You must enter a valid BCR ID."
            BCR_ID_TextBox.Value = "Enter BCR ID"
            Exit Sub
        End If
        If Me.bcrBox.Value = "<None>" Then
            MsgBox "You must map a BCR ID Field."
            Exit Sub
        End If
    End If

    Me.Tag = "Export"
    Me.Hide
    
End Sub

Private Sub msidBox_Change()

    If checkDuplicate(msidBox) = True Then
        MsgBox "Please select a unique IMS Field."
        msidBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(msidBox.Value) = False And msidBox.Value <> "" And msidBox.Value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        msidBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fMSID").Value = Me.msidBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fMSID", False, msoPropertyTypeString, Me.msidBox.Value
    Resume PropFound
End Sub

Private Sub mswBox_Change()

    If checkDuplicate(mswBox) = True Then
        MsgBox "Please select a unique IMS Field."
        mswBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(mswBox.Value) = False And mswBox.Value <> "" And mswBox.Value <> "<None>" Then
        MsgBox "Please select a valid IMS Field."
        mswBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fMSW").Value = Me.mswBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fMSW", False, msoPropertyTypeString, Me.mswBox.Value
    Resume PropFound
End Sub


Private Sub PercentBox_Change()

    If checkDuplicate(PercentBox) = True Then
        MsgBox "Please select a unique IMS Field."
        PercentBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(PercentBox.Value) = False And PercentBox.Value <> "" Then
        MsgBox "Please select a valid IMS Field."
        PercentBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fPCNT").Value = Me.PercentBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fPCNT", False, msoPropertyTypeString, Me.PercentBox.Value
    Resume PropFound
End Sub

Private Sub resBox_Change()
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fResID").Value = Me.resBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fResID", False, msoPropertyTypeString, Me.resBox.Value
    Resume PropFound
    
End Sub

Private Sub ResExportCheckbox_Click()

    If ResExportCheckbox.Value = True Then
        exportTPhaseCheckBox.Enabled = True
    Else
        exportTPhaseCheckBox.Enabled = False
    End If

End Sub

Private Sub RunDataBtn_Click()
    
    Me.Tag = "DataCheck"
    Me.Hide
    
End Sub

Private Sub TabButtons_Click(ByVal Index As Long)
    If Index <> 1 And Me.TabButtons(1).Tag = False Then
        Me.TabButtons.Value = 1
        Exit Sub
    End If
    If Index <> 1 And VerifyTitles = False Then
        Me.TabButtons.Value = 1
        MsgBox "Complete CA ID Titles"
        Exit Sub
    End If
End Sub

Private Sub UserForm_Activate()

    If Me.TabButtons(1).Tag = False Then
        Me.TabButtons.Value = 1
        MsgBox "Please complete the Custom Field Configuration"
    End If

End Sub

Private Sub UserForm_Initialize()

    Me.MPPBtn.Value = True
    Me.TabButtons.Value = 0
    Me.ExportBtn.SetFocus
    
    If CSVBtn.Value = True Then
        Me.BCWS_Checkbox.Enabled = True
        Me.BCWP_Checkbox.Enabled = True
        Me.ETC_Checkbox.Enabled = True
        Me.TotalProjBtn.Enabled = True
        Me.BcrBtn.Enabled = True
        Me.ResExportCheckbox.Enabled = True
    Else
        Me.BCWS_Checkbox.Enabled = False
        Me.BCWP_Checkbox.Enabled = False
        Me.ETC_Checkbox.Enabled = False
        Me.TotalProjBtn.Enabled = False
        Me.BcrBtn.Enabled = False
        Me.ResExportCheckbox.Enabled = False
    End If
    
    If Me.TotalProjBtn = False And Me.BcrBtn = False Then
        Me.TotalProjBtn = True
    End If
    
    Me.Tag = "Loading"
    Me.TabButtons(1).Tag = PopulateCustFieldUsage
    Me.Tag = "Loaded"

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = 0 Then
    Cancel = True
    cancelBtn_Click
  End If
End Sub
Private Function VerifyCustFieldUsage() As Boolean

    Dim fCAID1, fCAID2, fCAID3, fWP, fCAM, fEVT, fPCNT, fResID As Boolean
    
    If Me.caID1Box.Value <> "" Then fCAID1 = True
    If CAID2TxtBox.Value <> "<None>" Then
        If Me.caID2Box.Value <> "" Then fCAID2 = True
    Else
        fCAID2 = False
    End If
    If CAID3TxtBox.Value <> "<None>" Then
        If Me.caID3Box.Value <> "" Then fCAID3 = True
    Else
        fCAID3 = False
    End If
    If Me.resBox.Value <> "" Then fResID = True
    If Me.wpBox.Value <> "" Then fWP = True
    If Me.camBox.Value <> "" Then fCAM = True
    If Me.evtBox.Value <> "" Then fEVT = True
    If Me.PercentBox.Value <> "" Then fPCNT = True
    
    If fCAID1 And fCAID2 And fCAID3 And fWP And fCAM And fEVT And fPCNT And fResID Then
    
        VerifyCustFieldUsage = True
    
    Else
    
        VerifyCustFieldUsage = False
    
    End If

End Function

Private Function VerifyTitles() As Boolean

    Dim TitlesComplete As Boolean
    
    TitlesComplete = True
    
    If Me.CAID1TxtBox.Value = "" Then
        Me.CAID1TxtBox.BackColor = RGB(255, 255, 0)
        TitlesComplete = False
    Else
        Me.CAID1TxtBox.BackColor = RGB(255, 255, 255)
    End If
    
    If Me.caID2Box.Value <> "<None>" Then
        If Me.CAID2TxtBox.Value = "" Then
            Me.CAID2TxtBox.BackColor = RGB(255, 255, 0)
            TitlesComplete = False
        Else
            Me.CAID2TxtBox.BackColor = RGB(255, 255, 255)
        End If
    End If
    
    If Me.caID3Box.Value <> "<None>" Then
        If Me.CAID3TxtBox.Value = "" Then
            Me.CAID3TxtBox.BackColor = RGB(255, 255, 0)
            TitlesComplete = False
        Else
            Me.CAID3TxtBox.BackColor = RGB(255, 255, 255)
        End If
    End If
    
    VerifyTitles = TitlesComplete

End Function
Private Function PopulateCustFieldUsage() As Boolean

    Dim curProj As Project
    Dim docProp As DocumentProperty
    Dim docProps As DocumentProperties
    Dim fCAID1, fCAID1t, fCAID3, fCAID3t, fWP, fCAM, fEVT, fCAID2, fCAID2t, fMSID, fMSW, fBCR, fPCNT, fResID As Boolean
    Dim NameTest As Double
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo DocPropNameChange
    
    For Each docProp In docProps
    
        Select Case docProp.Name
        
            Case "fCAID1"
            
                NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                fCAID1 = True
                Me.caID1Box.Value = docProp.Value
                
            Case "fCAID1t"
            
                fCAID1t = True
                Me.CAID1TxtBox.Value = docProp.Value
                
            Case "fCAID3"
                
                If docProp.Value = "<None>" Then
                    fCAID3 = True
                    Me.caID3Box.Value = docProp.Value
                Else
                    NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                    fCAID3 = True
                    Me.caID3Box.Value = docProp.Value
                End If
                
            Case "fCAID3t"
                
                fCAID3t = True
                Me.CAID3TxtBox.Value = docProp.Value
                
            Case "fCAID2"
            
                If docProp.Value = "<None>" Then
                    fCAID2 = True
                    Me.caID2Box.Value = docProp.Value
                Else
                    NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                    fCAID2 = True
                    Me.caID2Box.Value = docProp.Value
                End If
                
            Case "fCAID2t"
            
                fCAID2t = True
                Me.CAID2TxtBox.Value = docProp.Value
                
            Case "fWP"
                
                NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                fWP = True
                Me.wpBox.Value = docProp.Value
                
            Case "fCAM"
                
                NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                fCAM = True
                Me.camBox.Value = docProp.Value
                
            Case "fEVT"
                
                NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                fEVT = True
                Me.evtBox.Value = docProp.Value
                
            Case "fCAID2"
            
                NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                fCAID2 = True
                Me.caID2Box.Value = docProp.Value
                
            Case "fCAID2t"
            
                fCAID2t = True
                Me.CAID2TxtBox.Value = docProp.Value
                
            Case "fMSID"
                
                If docProp.Value = "<None>" Then
                    fMSID = True
                    Me.msidBox.Value = docProp.Value
                Else
                    NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                    fMSID = True
                    Me.msidBox.Value = docProp.Value
                End If
                
            Case "fMSW"
                
                If docProp.Value = "<None>" Then
                    fMSW = True
                    Me.mswBox.Value = docProp.Value
                Else
                    NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                    fMSW = True
                    Me.mswBox.Value = docProp.Value
                End If
                
            Case "fBCR"
            
                If docProp.Value = "<None>" Then
                    fBCR = True
                    Me.bcrBox.Value = docProp.Value
                Else
                    NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                    fBCR = True
                    Me.bcrBox.Value = docProp.Value
                End If
                
            Case "fPCNT"
            
                NameTest = ActiveProject.Application.FieldNameToFieldConstant(docProp.Value)
                fPCNT = True
                Me.PercentBox.Value = docProp.Value
                
            Case "fResID"
            
                fResID = True
                Me.resBox.Value = docProp.Value
            
            Case Else
        
        End Select
    
NextDocProp:
    
    Next docProp
    
    Set docProps = Nothing
    Set curpro = Nothing
    
    If fCAID1 And fCAID2 And fWP And fCAM And fEVT And fCAID3 And fMSID And fMSW And fPCNT And fResID Then
    
        PopulateCustFieldUsage = True
    
    Else
    
        PopulateCustFieldUsage = False
    
    End If
    
    Exit Function
    
DocPropNameChange:

    Resume NextDocProp

End Function

Private Sub wpBox_Change()

    If checkDuplicate(wpBox) = True Then
        MsgBox "Please select a unique IMS Field."
        wpBox.Value = ""
        Exit Sub
    End If
    
    If isIMSfield(wpBox.Value) = False And wpBox.Value <> "" Then
        MsgBox "Please select a valid IMS Field."
        wpBox.Value = ""
        Exit Sub
    End If
    
    Dim docProps As DocumentProperties
    Dim curProj As Project
    
    Set curProj = ActiveProject
    Set docProps = curProj.CustomDocumentProperties
    
    On Error GoTo PropMissing
    
    docProps("fWP").Value = Me.wpBox.Value

PropFound:

    Me.TabButtons(1).Tag = VerifyCustFieldUsage
    
    Set docProps = Nothing
    Set curProj = Nothing
    
    Exit Sub
    
PropMissing:

    docProps.Add "fWP", False, msoPropertyTypeString, Me.wpBox.Value
    Resume PropFound
End Sub
Private Function isIMSfield(ByVal mspField As String) As Boolean

    On Error GoTo fieldMissing
    
    Dim curProj As Project
    Set curProj = ActiveProject
    
    If curProj.Application.FieldNameToFieldConstant(mspField) Then
    
        isIMSfield = True
        Set curProj = Nothing
        Exit Function
    
    End If
    
fieldMissing:

    isIMSfield = False
    Set curProj = Nothing

End Function
