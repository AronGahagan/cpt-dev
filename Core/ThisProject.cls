VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'<cpt_version>v1.1</cpt_version>

Private Sub Project_Activate(ByVal pj As Project)
Dim ribbonXML As String

  'standard xml opening wrapper
  ribbonXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
  ribbonXML = ribbonXML + "<mso:customUI "
  ribbonXML = ribbonXML + "xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:ribbon startFromScratch=""false"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:tabs>"

  'this line adds the ClearPlan tab
  ribbonXML = ribbonXML + cptBuildRibbonTab()
  
  '/===== enter other tab level customizatzions below =====\

  '\===== enter other tab level customizatzions above =====/
  
  'standard xml closing wrapper
  ribbonXML = ribbonXML + vbCrLf & "</mso:tabs>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:ribbon>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:customUI>"

  'this actually adds the toolbar
  'method only available in the Project_Activate event
  pj.SetCustomUI (ribbonXML)
  
  'start monitoring for WindowSelectionChange events
  Call cptCore_bas.StartEvents

End Sub

Private Sub Project_Open(ByVal pj As Project)

  'cpt: start monitoring for WindowSelectionChange events
  Call cptCore_bas.StartEvents
  'cpt: set references - in case something isn't late bound
  Call cptCore_bas.SetReferences

End Sub
