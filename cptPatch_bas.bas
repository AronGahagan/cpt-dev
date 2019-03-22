Attribute VB_Name = "cptPatch_bas"
'<cpt_version>19.03.22</cpt_version> this one based on date vs. SemVer
'this file will update with code to run to apply deep code updates if necessary
Option Explicit

Private Sub cptApplyPatch()
'objects
'strings
'longs
'integers
'booleans
'variants
'dates

  If BLN_TRAP_ERRORS Then On Error GoTo err_here Else On Error GoTo 0

  'patch code goes here

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call HandleErr("cptPatch_bas", "cptApplyPatch()", err)
  Resume exit_here
End Sub
