Attribute VB_Name = "cptSmartDuration_bas"
'<cpt_version>v1.0</cpt_version>
Sub SmartDuration()

    Dim smrtForm As cptSmartDuration_frm
    Dim t As Task
    Dim curproj As Project
    Dim response As Variant
    
    Set curproj = ActiveProject
    
    If curproj.Tasks.count < 0 Then
        MsgBox "No Tasks Found"
        Exit Sub
    End If
    
    On Error GoTo NotATask
    
    Set t = curproj.Application.ActiveCell.Task
    
    If t.Summary = True Then
        MsgBox "Please select a non-Summary Task"
        Set t = Nothing
        Set curproj = Nothing
        Exit Sub
    End If
    
    If t.Milestone = True Then
        response = MsgBox("Proceed with editing a Milestone?", vbYesNo)
        If response = vbNo Then
            Set t = Nothing
            Set curproj = Nothing
            Exit Sub
        End If
    End If
    
    If t.ActualFinish <> "NA" Then
        MsgBox "Please select an incomplete Task"
        Set t = Nothing
        Set curproj = Nothing
        Exit Sub
    End If
    
    Set smrtForm = New cptSmartDuration_frm
    
    With smrtForm
    
        .StartDate = t.Start
        .SmartDatePicker.Text = t.GetField(pjTaskFinish)
        .weekDayLbl = Format(t.Finish, "DDD")
    
        .Show
        
        If .Tag = "Cancel" Then
            
            GoTo CleanUp
            
        End If
    
        If .Tag = "OK" Then
            
            If InStr(t.GetField(pjTaskDuration), "e") > 0 Then
            
                OpenUndoTransaction "Smart Duration"
                t.duration = VBA.DateDiff("n", t.Start, .finDate)
                CloseUndoTransaction
                GoTo CleanUp
            
            Else
                
                If t.Calendar = "None" Or t.Calendar = curproj.Calendar Then
                    OpenUndoTransaction "Smart Duration"
                    t.duration = Application.DateDifference(t.Start, .finDate)
                    CloseUndoTransaction
                    GoTo CleanUp
                Else
                    OpenUndoTransaction "Smart Duration"
                    t.duration = Application.DateDifference(t.Start, .finDate, t.Calendar)
                    CloseUndoTransaction
                    GoTo CleanUp
                End If
            End If
            
        End If
            
    End With
    
CleanUp:

    Set t = Nothing
    Set curproj = Nothing
    Set smrtForm = Nothing
    
    Exit Sub
    
NotATask:

    MsgBox "Please select a valid Task"
    Set curproj = Nothing
    Exit Sub

End Sub
