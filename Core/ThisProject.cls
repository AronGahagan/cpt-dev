VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'<cpt_version>v0.1</cpt_version>

Option Explicit

Private Sub Project_Activate(ByVal pj As Project)
Dim ribbonXML As String, blnDebug As Boolean
Dim strSuperTip As String, lgCleanUp As Long
Dim blnOnLoad As Boolean, strMethod As String
    
  lgCleanUp = 0
  
  ribbonXML = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
  ribbonXML = ribbonXML + "<mso:customUI "
  ribbonXML = ribbonXML + "xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:ribbon startFromScratch=""false"" >"
  ribbonXML = ribbonXML + vbCrLf & "<mso:tabs>"
  
  ribbonXML = ribbonXML + cptBuildRibbonTab()
  
  ribbonXML = ribbonXML + vbCrLf & "</mso:tabs>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:ribbon>"
  ribbonXML = ribbonXML + vbCrLf & "</mso:customUI>"
  
  'Debug.Print ribbonXml
  pj.SetCustomUI (ribbonXML)
  Call cptCore_bas.StartEvents
End Sub

Private Sub Project_Open(ByVal pj As Project)
  Call cptCore_bas.SetReferences
  Call cptCore_bas.StartEvents
End Sub
