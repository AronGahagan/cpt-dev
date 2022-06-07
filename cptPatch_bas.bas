Attribute VB_Name = "cptPatch_bas"
'<cpt_version>21.04.10</cpt_version> this one based on date vs. SemVer
Option Explicit

Public Sub cptApplyPatch()
'objects
'strings
'longs
'integers
'booleans
'variants
'dates

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  'patch code goes here
  Application.StatusBar = "Applying patch 21.04.10..."
  If Not cptReferenceExists("VBScript_RegExp_55") Then
    ThisProject.VBProject.References.AddFromFile "C:\WINDOWS\System32\vbscript.dll\3"
  End If

exit_here:
  On Error Resume Next
  Application.StatusBar = ""
  Exit Sub
err_here:
  Call cptHandleErr("cptPatch_bas", "cptApplyPatch()", Err, Erl)
  Resume exit_here
End Sub
