Attribute VB_Name = "cptSmartDur_bas"
'<cpt_version>v0</cpt_version>

Sub SmartDuration()

    Dim smrtForm As cptSmartDur_frm
    Dim t As Task
    Dim curProj As Project
    Dim response As Variant
    
    Set curProj = ActiveProject
    
    If curProj.Tasks.count < 0 Then
        MsgBox "No Tasks Found"
        Exit Sub
    End If
    
    On Error GoTo NotATask
    
    Set t = curProj.Application.ActiveCell.Task
    
    If t.Summary = True Then
        MsgBox "Please select a non-Summary Task"
        Set t = Nothing
        Set curProj = Nothing
        Exit Sub
    End If
    
    If t.Milestone = True Then
        response = MsgBox("Proceed with editing a Milestone?", vbYesNo)
        If response = vbNo Then
            Set t = Nothing
            Set curProj = Nothing
            Exit Sub
        End If
    End If
    
    If t.ActualFinish <> "NA" Then
        MsgBox "Please select an incomplete Task"
        Set t = Nothing
        Set curProj = Nothing
        Exit Sub
    End If
    
    Set smrtForm = New cptSmartDur_frm
    
    With smrtForm
    
        .startDate = t.Start
        .SmartDatePicker.Text = t.GetField(pjTaskFinish)
        .weekDayLbl = Format(t.Finish, "DDD")
    
        .Show
        
        If .Tag = "Cancel" Then
            
            GoTo CleanUp
            
        End If
    
        If .Tag = "OK" Then
            If t.Calendar = "None" Or t.Calendar = curProj.Calendar Then
                OpenUndoTransaction "Smart Duration"
                t.Duration = Application.DateDifference(t.Start, .finDate)
                CloseUndoTransaction
                GoTo CleanUp
            Else
                OpenUndoTransaction "Smart Duration"
                t.Duration = Application.DateDifference(t.Start, .finDate, t.Calendar)
                CloseUndoTransaction
                GoTo CleanUp
            End If
            
        End If
            
    End With
    
CleanUp:

    Set t = Nothing
    Set curProj = Nothing
    Set smrtForm = Nothing
    
    Exit Sub
    
NotATask:

    MsgBox "Please select a valid Task"
    Set curProj = Nothing
    Exit Sub

End Sub
