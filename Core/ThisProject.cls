VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'<cpt_version>v1.3.0</cpt_version>

Private Sub Project_Activate(ByVal pj As Project)
Dim ribbonXML As String '</cpt>

  'standard xml opening wrapper '</cpt>
  If Dir(Environ("USERPROFILE") & "\Project Customizations.exportedUI") = vbNullString Then '</cpt>
    ribbonXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf '</cpt>
    ribbonXML = ribbonXML + "<mso:customUI " '</cpt>
    ribbonXML = ribbonXML + "xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >" '</cpt>
    ribbonXML = ribbonXML + vbCrLf & "<mso:ribbon startFromScratch=""false"" >" '</cpt>
  Else '</cpt>
    ribbonXML = ribbonXML + cptGetUserCustomizations("xmlns") '</strip> 'todo: this function exists in cptGA_bas and nowhere else; move it to cptSetup
    ribbonXML = ribbonXML + "<mso:ribbon>" '</strip>
  End If '</cpt>
  ribbonXML = ribbonXML + vbCrLf & "<mso:tabs>" '</cpt>
  '</cpt>
  'this line adds the ClearPlan tab '</cpt>
  ribbonXML = ribbonXML + cptBuildRibbonTab() '</cpt>
  '</cpt>
  '/===== enter other tab level customizatzions below =====\ '</cpt>
  ribbonXML = ribbonXML + keoBuildRibbonTab() '</strip>
  ribbonXML = ribbonXML + SRA_GetRibbon() '</strip>
  ribbonXML = ribbonXML + gitBuildRibbonTab() '</strip>
'  ribbonXML = ribbonXML + getJPLRibbonTab() '</strip>
'  ribbonXML = ribbonXML + cptGetUserCustomizations("tabs") '</strip>
  '\===== enter other tab level customizatzions above =====/ '</cpt>
  '</cpt>
  'standard xml closing wrapper '</cpt>
  ribbonXML = ribbonXML + vbCrLf & "</mso:tabs>" '</cpt>
'  ribbonXML = ribbonXML + vbCrLf & cptGetUserCustomizations("contextualTabs") '</strip>
  ribbonXML = ribbonXML + vbCrLf & "</mso:ribbon>" '</cpt>
  ribbonXML = ribbonXML + vbCrLf & "</mso:customUI>" '</cpt>
  '</cpt>
  'this actually adds the toolbar '</cpt>
  'Debug.Print ribbonXML '</strip>
  pj.SetCustomUI (ribbonXML)  '</cpt>
  '</cpt>
  'start monitoring for WindowSelectionChange events '</cpt>
  Call cptCore_bas.cptStartEvents '</cpt>

  pj.NewTasksCreatedAsManual = False '</strip>
  ViewEditSingle Name:="Task Details Form", ShowInMenu:=True '</strip>
  
End Sub

Private Sub Project_Open(ByVal pj As Project)

  'start monitoring for WindowSelectionChange events '</cpt>
  Call cptCore_bas.cptStartEvents '</cpt>
  'set references - in case something isn't late bound '</cpt>
  Call cptCore_bas.cptSetReferences '</cpt>

End Sub
