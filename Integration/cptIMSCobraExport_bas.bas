Attribute VB_Name = "cptIMSCobraExport_bas"
'<cpt_version>v3.3.4</cpt_version>
Option Explicit
Private destFolder As String
Private BCWSxport As Boolean
Private BCWPxport As Boolean
Private ETCxport As Boolean
Private WhatIfxport As Boolean 'v3.2
Private ResourceLoaded As Boolean
Private MasterProject As Boolean
Private ACTfilename As String
Private RESfilename As String
Private BCR_WP() As String
Private BCR_ID As String
Private BCRxport As Boolean
Private BCR_Error As Boolean
Private fCAID1, fCAID1t, fCAID3, fCAID3t, fWP, fCAM, fPCNT, fAssignPcnt, fEVT, fCAID2, fCAID2t, fMilestone, fMilestoneWeight, fBCR, fWhatIf, fResID As String 'v3.3.0
Private CustTextFields() As String
Private EntFields() As String
Private CustNumFields() As String
Private CustOLCodeFields() As String
Private ActFound As Boolean
Private CAID3_Used As Boolean
Private CAID2_Used As Boolean
Private Milestones_Used As Boolean
Private AssignmentPCNT_Used As Boolean 'v3.3.2
Private TimeScaleExport As Boolean
Private DescExport As Boolean
Private ErrMsg As String
Private WPDescArray() As WP_Descriptions
Private WPDescCount As Integer
Private Type WP_Descriptions
    WP_ID As String
    Desc As String
End Type
Private Type ACTrowWP
    CAID1 As String
    CAID3 As String
    CAID2 As String
    Desc As String
    CAM As String
    WP As String
    Resource As String
    ID As String
    BStart As Date
    BFinish As Date
    FStart As Date
    FFinish As Date
    AStart As Date
    AFinish As Date
    EVT As String
    sumBCWS As Double
    sumBCWP As Double
    Prog As Integer
End Type
Private Type WPDataCheck
    WP_ID As String
    ID_Test As String
    EVT_Test As String
    WP_DupError As Boolean
    EVT_Error As Boolean
End Type
Private Type CAMDataCheck
    
    ID_str As String '**CAID1/CAID2/CAID3**
    CAM_Test As String
    CAM_Error As Boolean
    
End Type
Private Type TaskDataCheck
    UID As String
    WP As String
    CAID1 As String
    CAID2 As String
    CAID3 As String
    CAM As String
    BStart As String
    BFinish As String
    FStart As String 'v3.3.0
    FFinish As String 'v3.3.0
    BWork As Double
    BCost As Double
    FWork As Double 'v3.3.0
    FCost As Double 'v3.3.0
    AssignmentBStart As String
    AssignmentBFinish As String
    AssignmentFStart As String 'v3.3.0
    AssignmentFFinish As String 'v3.3.0
    AssignmentBWork As Double
    AssignmentBCost As Double
    AssignmentFWork As Double 'v3.3.0
    AssignmentFCost As Double 'v3.3.0
    AssignmentTSVWork As Double 'v3.3.2
    AssignmentTSVCost As Double 'v3.3.2
    EVT As String
    MSID As String
    MSWeight As String
    AssignmentCount As Integer
End Type

Sub Export_IMS()

    Dim xportFrm As cptIMSCobraExport_frm
    Dim xportFormat As String
    Dim curproj As Project
    Dim i As Integer

    On Error GoTo CleanUp

    Set curproj = ActiveProject

    curproj.Application.Calculation = pjManual
    curproj.Application.DisplayAlerts = False

    If curproj.Subprojects.Count > 0 And InStr(curproj.FullName, "<>") > 0 And curproj.ReadOnly <> True Then
        MsgBox "Master Project Files with Subprojects must be opened Read Only"
        GoTo Quick_Exit
    End If

    If curproj.Subprojects.Count > 0 Then
        MasterProject = True
    Else
        MasterProject = False
    End If

    ReadCustomFields curproj

    Set xportFrm = New cptIMSCobraExport_frm

    With xportFrm

        On Error Resume Next

        .resBox.List = Split("Name,Code,Initials", ",")

        'populate listboxes
        Dim vArray As Variant
        vArray = Split(Join(CustTextFields, ",") & Join(CustOLCodeFields, ",") & Join(EntFields, ","), ",")
        Call cptQuickSort(vArray, 0, UBound(vArray))
        .caID1Box.List = Split("WBS," & Join(vArray, ","), ",")
        .caID2Box.List = Split("<None>," & Join(vArray, ","), ",")
        .caID3Box.List = Split("<None>," & Join(vArray, ","), ",")
        .wpBox.List = vArray
        .camBox.List = Split("Contact," & Join(vArray, ","), ",")
        .evtBox.List = vArray
        .mswBox.List = Split("<None>,BaselineWork,BaselineCost,Work,Cost," & Join(vArray, ","), ",")
        .bcrBox.List = Split("<None>," & Join(vArray, ","), ",")
        .whatifBox.List = Split("<None>," & Join(vArray, ","), ",")
        vArray = Split(Join(CustTextFields, ",") & Join(CustNumFields, ",") & Join(CustOLCodeFields, ",") & Join(EntFields, ","), ",")
        Call cptQuickSort(vArray, 0, UBound(vArray))
        .msidBox.List = Split("<None>,UniqueID," & Join(vArray, ","), ",")
        Call cptQuickSort(CustNumFields, 1, UBound(CustNumFields))
        .PercentBox.List = Split("Physical % Complete,% Complete," & Join(CustNumFields, ","), ",")
        .AsgnPcntBox.List = Split("<None>," & Join(CustNumFields, ","), ",")
        
        On Error GoTo CleanUp
        ErrMsg = "Please try again, or contact the developer if this message repeats."
        '********************************************
        'On Error GoTo 0 '**Used for Debugging ONLY**
        '********************************************

        .Show

        If .Tag = "Cancel" Then
            Set xportFrm = Nothing
            Set curproj = Nothing
            Exit Sub
        End If

        If .Tag = "DataCheck" Then
            CAID3_Used = .CAID3TxtBox.Enabled
            CAID2_Used = .CAID2TxtBox.Enabled
            DataChecks curproj
            Set xportFrm = Nothing
            GoTo Quick_Exit
        End If

        If .MPPBtn.Value = True Then
            Set xportFrm = Nothing
            xportFormat = "MPP"
        ElseIf .XMLBtn.Value = True Then
            Set xportFrm = Nothing
            xportFormat = "XML"
        ElseIf .CSVBtn.Value = True Then
            BCWSxport = .BCWS_Checkbox.Value
            BCWPxport = .BCWP_Checkbox.Value
            ETCxport = .ETC_Checkbox.Value
            BCRxport = .BcrBtn.Value
            WhatIfxport = .WhatIf_CheckBox.Value 'v3.2
            BCR_ID = .BCR_ID_TextBox
            ResourceLoaded = .ResExportCheckbox
            TimeScaleExport = .exportTPhaseCheckBox
            Set xportFrm = Nothing
            xportFormat = "CSV"
            CAID3_Used = .CAID3TxtBox.Enabled
            CAID2_Used = .CAID2TxtBox.Enabled
            If .msidBox.Value = "<None>" Or .mswBox.Value = "<None>" Or .msidBox.Value = "" Or .mswBox.Value = "" Then
                Milestones_Used = False
            Else
                Milestones_Used = True
            End If
            If .AsgnPcntBox = "<None>" Or .AsgnPcntBox = "" Then
                AssignmentPCNT_Used = False
            Else
                AssignmentPCNT_Used = True
            End If
            DescExport = .exportDescCheckBox.Value
        End If

    End With

    Select Case xportFormat

        Case "MPP"

            MPP_Export curproj

        Case "XML"

            XML_Export curproj

        Case "CSV"

            CSV_Export curproj

        Case Else

    End Select

    If BCR_Error = False Then
        MsgBox "IMS Export saved to " & destFolder
        Shell "explorer.exe" & " " & destFolder, vbNormalFocus
    End If

    curproj.Application.Calculation = pjAutomatic
    curproj.Application.DisplayAlerts = True
    Set curproj = Nothing

    Exit Sub

CleanUp:

    If ACTfilename <> "" Then Reset

    curproj.Application.Calculation = pjAutomatic
    curproj.Application.DisplayAlerts = True
    Set curproj = Nothing
    MsgBox "An error was encountered." & vbCr & vbCr & ErrMsg
    Exit Sub

Quick_Exit:

    curproj.Application.Calculation = pjAutomatic
    curproj.Application.DisplayAlerts = True
    Set curproj = Nothing

    Exit Sub

End Sub

Private Function get_assignment_timescalevalues(ByVal tAss As Assignment) As Double 'v3.3.2
    
    Dim tsvs As TimeScaleValues
    Dim tsv As TimeScaleValue
    Dim tempTotal As Double
    
    If tAss.Resource.Type = pjResourceTypeCost Then
        Set tsvs = tAss.TimeScaleData(tAss.BaselineStart, tAss.BaselineFinish, pjAssignmentTimescaledBaselineCost)
    Else
        Set tsvs = tAss.TimeScaleData(tAss.BaselineStart, tAss.BaselineFinish, pjAssignmentTimescaledBaselineWork)
    End If
    
    tempTotal = 0
    
    For Each tsv In tsvs
    
        If tsv.Value <> "" Then
            
            tempTotal = tempTotal + tsv.Value
        
        End If
        
    Next
    
    get_assignment_timescalevalues = tempTotal

End Function

Private Sub DataChecks(ByVal curproj As Project)

    Dim WPChecks() As WPDataCheck
    Dim wpFound As Boolean
    Dim CAMChecks() As CAMDataCheck
    Dim CAfound As Boolean
    Dim TaskChecks() As TaskDataCheck
    Dim taskFound As Boolean
    Dim t As Task
    Dim tAss As Assignment
    Dim tasses As Assignments
    Dim tAssBStart As String
    Dim tAssBFin As String
    Dim tAssBWork As String
    Dim tempID As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim wpCount As Integer
    Dim camCount As Integer
    Dim taskCount As Integer
    Dim X As Integer
    Dim i As Integer
    Dim errorStr As String
    Dim ErrorCounter As Integer
    Dim tempBValue As Double
    Dim tempBWork As Double
    Dim tempFWork As Double 'v3.3.0
    Dim tempTSVWork As Double 'v3.3.2

    Dim docProps As DocumentProperties

    Set docProps = curproj.CustomDocumentProperties

    fCAID1 = docProps("fCAID1").Value
    fCAID1t = docProps("fCAID1t").Value
    If BCRxport = True Then
        fBCR = docProps("fBCR").Value
    End If
    If CAID3_Used = True Then
        fCAID3 = docProps("fCAID3").Value
        fCAID3t = docProps("fCAID3t").Value
    End If
    fWP = docProps("fWP").Value
    fCAM = docProps("fCAM").Value
    fEVT = docProps("fEVT").Value
    If CAID2_Used = True Then
        fCAID2 = docProps("fCAID2").Value
        fCAID2t = docProps("fCAID2t").Value
    End If
    If Milestones_Used = True Then 'v3.2.6
        fMilestone = docProps("fMSID").Value
        fMilestoneWeight = docProps("fMSW").Value
    End If
    fPCNT = docProps("fPCNT").Value

    destFolder = SetDirectory(curproj.ProjectSummaryTask.Name)

    taskCount = 0
    taskFound = False

    '**Scan Task Data**

    If curproj.Subprojects.Count > 0 Then

        Set subProjs = curproj.Subprojects

        For Each subProj In subProjs

            FileOpen Name:=subProj.Path, ReadOnly:=True

            Set curSProj = ActiveProject

            For Each t In curSProj.Tasks

                If Not t Is Nothing Then

                    If t.Summary = False And t.GetField(188744959) = "Yes" And t.ExternalTask = False Then

                        taskCount = taskCount + 1
                        taskFound = True
                        ReDim Preserve TaskChecks(1 To taskCount)

                        With TaskChecks(taskCount)

                            .UID = t.UniqueID
                            .WP = t.GetField(FieldNameToFieldConstant(fWP))
                            .CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID2_Used = True Then
                                .CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True Then
                                .CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            .EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            If Milestones_Used = True Then
                                .MSID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                .MSWeight = t.GetField(FieldNameToFieldConstant(fMilestoneWeight))
                            End If
                            .BWork = t.BaselineWork / 60 'v3.3.2
                            .BCost = t.BaselineCost
                            .FWork = t.Work / 60 'v3.3.2
                            .FCost = t.Cost 'v3.3.0
                            .CAM = t.GetField(FieldNameToFieldConstant(fCAM))
                            .AssignmentBStart = "NA"
                            .AssignmentBFinish = "NA"
                            .AssignmentFStart = "NA" 'v3.3.0
                            .AssignmentFFinish = "NA" 'v3.3.0
                            .AssignmentBCost = 0
                            .AssignmentBWork = 0
                            .AssignmentFCost = 0 'v3.3.0
                            .AssignmentFWork = 0 'v3.3.0
                            .AssignmentTSVWork = 0 'v3.3.2
                            .AssignmentTSVCost = 0 'v3.3.2
                            .BStart = t.BaselineStart
                            .BFinish = t.BaselineFinish
                            .FStart = t.Start 'v3.3.0
                            .FFinish = t.Finish 'v3.3.0

                            Set tasses = t.Assignments
                            .AssignmentCount = tasses.Count

                            For Each tAss In tasses

                                If tAss.BaselineStart <> "NA" Then
                                    If .AssignmentBStart = "NA" Then
                                        .AssignmentBStart = tAss.BaselineStart
                                    Else
                                        If tAss.BaselineStart < .AssignmentBStart Then
                                            .AssignmentBStart = tAss.BaselineStart
                                        End If
                                    End If
                                End If

                                If tAss.BaselineFinish <> "NA" Then
                                    If .AssignmentBFinish = "NA" Then
                                        .AssignmentBFinish = tAss.BaselineFinish
                                    Else
                                        If tAss.BaselineFinish > .AssignmentBFinish Then
                                            .AssignmentBFinish = tAss.BaselineFinish
                                        End If
                                    End If
                                End If
                                
                                If .AssignmentFStart = "NA" Then 'v3.3.0
                                    .AssignmentFStart = tAss.Start
                                Else
                                    If tAss.FStart < .AssignmentFStart Then
                                        .AssignmentFStart = tAss.Start
                                    End If
                                End If

                                If .AssignmentFFinish = "NA" Then 'v3.3.0
                                    .AssignmentFFinish = tAss.Finish
                                Else
                                    If tAss.Finish > .AssignmentFFinish Then
                                        .AssignmentFFinish = tAss.Finish
                                    End If
                                End If

                                .AssignmentBCost = .AssignmentBCost + tAss.BaselineCost
                                If tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.2
                                    .AssignmentTSVCost = .AssignmentTSVCost + get_assignment_timescalevalues(tAss)
                                End If
                                .AssignmentFCost = .AssignmentFCost + tAss.Cost 'v3.3.0
                                
                                If tAss.BaselineWork = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.2.1
                                    tempBWork = 0
                                    tempTSVWork = 0 'v3.3.2
                                Else
                                    tempBWork = tAss.BaselineWork
                                    tempTSVWork = get_assignment_timescalevalues(tAss) 'v3.3.2
                                End If
                                .AssignmentBWork = .AssignmentBWork + tempBWork / 60 'v3.3.2
                                .AssignmentTSVWork = .AssignmentTSVWork + tempTSVWork / 60 'v3.3.2
                                
                                If tAss.Work = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.0
                                    tempFWork = 0
                                Else
                                    tempFWork = tAss.Work
                                End If
                                .AssignmentFWork = .AssignmentFWork + tempFWork / 60 'v3.3.2

                            Next tAss

                        End With

                    End If

                End If

            Next t

            FileClose pjDoNotSave

        Next subProj

    Else

        For Each t In curproj.Tasks

            If Not t Is Nothing Then

                If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                    taskCount = taskCount + 1
                    taskFound = True
                    ReDim Preserve TaskChecks(1 To taskCount)

                    With TaskChecks(taskCount)

                        .UID = t.UniqueID
                            .WP = t.GetField(FieldNameToFieldConstant(fWP))
                            .CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID2_Used = True Then
                                .CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True Then
                                .CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            .EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            If Milestones_Used = True Then
                                .MSID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                .MSWeight = t.GetField(FieldNameToFieldConstant(fMilestoneWeight))
                            End If
                            .BWork = t.BaselineWork / 60 'v3.3.2
                            .BCost = t.BaselineCost
                            .FWork = t.Work / 60 'v3.3.2
                            .FCost = t.Cost 'v3.3.0
                            .CAM = t.GetField(FieldNameToFieldConstant(fCAM))
                            .AssignmentBStart = "NA"
                            .AssignmentBFinish = "NA"
                            .AssignmentFStart = "NA" 'v3.3.0
                            .AssignmentFFinish = "NA" 'v3.3.0
                            .AssignmentBCost = 0
                            .AssignmentBWork = 0
                            .AssignmentFCost = 0 'v3.3.0
                            .AssignmentFWork = 0 'v3.3.0
                            .AssignmentTSVWork = 0 'v3.3.2
                            .AssignmentTSVCost = 0 'v3.3.2
                            .BStart = t.BaselineStart
                            .BFinish = t.BaselineFinish
                            .FStart = t.Start 'v3.3.0
                            .FFinish = t.Finish 'v3.3.0

                            Set tasses = t.Assignments
                            .AssignmentCount = tasses.Count

                            For Each tAss In tasses

                                If tAss.BaselineStart <> "NA" Then
                                    If .AssignmentBStart = "NA" Then
                                        .AssignmentBStart = tAss.BaselineStart
                                    Else
                                        If tAss.BaselineStart < .AssignmentBStart Then
                                            .AssignmentBStart = tAss.BaselineStart
                                        End If
                                    End If
                                End If

                                If tAss.BaselineFinish <> "NA" Then
                                    If .AssignmentBFinish = "NA" Then
                                        .AssignmentBFinish = tAss.BaselineFinish
                                    Else
                                        If tAss.BaselineFinish > .AssignmentBFinish Then
                                            .AssignmentBFinish = tAss.BaselineFinish
                                        End If
                                    End If
                                End If
                                
                                If .AssignmentFStart = "NA" Then 'v3.3.0
                                    .AssignmentFStart = tAss.Start
                                Else
                                    If tAss.Start < .AssignmentFStart Then
                                        .AssignmentFStart = tAss.Start
                                    End If
                                End If

                                If .AssignmentFFinish = "NA" Then 'v3.3.0
                                    .AssignmentFFinish = tAss.Finish
                                Else
                                    If tAss.Finish > .AssignmentFFinish Then
                                        .AssignmentFFinish = tAss.Finish
                                    End If
                                End If

                                .AssignmentBCost = .AssignmentBCost + tAss.BaselineCost
                                If tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.2
                                    .AssignmentTSVCost = .AssignmentTSVCost + get_assignment_timescalevalues(tAss)
                                End If
                                .AssignmentFCost = .AssignmentFCost + tAss.Cost 'v3.3.0
                                
                                If tAss.BaselineWork = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.2.1
                                    tempBWork = 0
                                    tempTSVWork = 0 'v3.3.2
                                Else
                                    tempBWork = tAss.BaselineWork
                                    tempTSVWork = get_assignment_timescalevalues(tAss) 'v3.3.2
                                End If
                                .AssignmentBWork = .AssignmentBWork + tempBWork / 60 'v3.3.2
                                .AssignmentTSVWork = .AssignmentTSVWork + tempTSVWork / 60 'v3.3.2
                                
                                If tAss.Work = "" Or tAss.ResourceType <> pjResourceTypeWork Then 'v3.3.0
                                    tempFWork = 0
                                Else
                                    tempFWork = tAss.Work
                                End If
                                .AssignmentFWork = .AssignmentFWork + tempFWork / 60 'v3.3.2

                            Next tAss

                    End With

                End If

            End If

        Next t

    End If

    ACTfilename = destFolder & "\DataChecks_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

    Open ACTfilename For Output As #1

    Print #1, "Tasks Missing Data - The following tasks are assumed to be EV Relevant Activities due to User population of one or more of the following data values: Work Package ID; Earned Value Technique; Baseline Work; Baseline Cost"

    If CAID3_Used = True And CAID2_Used = True Then
        Print #1, vbCrLf & "UID," & fCAID1t & "," & fCAID2t & "," & fCAID3t & ",CAM,WP,EVT,Baseline Value,Baseline Start,Baseline Finish,Milestone ID (As Req),Milestone Weight (As Req)"
    End If
    If CAID3_Used = False And CAID2_Used = True Then
        Print #1, vbCrLf & "UID," & fCAID1t & "," & fCAID2t & ",CAM,WP,EVT,Baseline Value,Baseline Start,Baseline Finish,Milestone ID (As Req),Milestone Weight (As Req)"
    End If
    If CAID3_Used = False And CAID2_Used = False Then
        Print #1, vbCrLf & "UID," & fCAID1t & ",CAM,WP,EVT,Baseline Value,Baseline Start,Baseline Finish,Milestone ID (As Req),Milestone Weight (As Req)"
    End If

    '**Evaluate WP and CAM data**
    wpCount = 0
    camCount = 0

    ErrorCounter = 0

    For X = 1 To taskCount

        If CAID3_Used = True And CAID2_Used = True Then
            tempID = TaskChecks(X).CAID1 & "/" & TaskChecks(X).CAID2 & "/" & TaskChecks(X).CAID3
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            tempID = TaskChecks(X).CAID1 & "/" & TaskChecks(X).CAID2
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            tempID = TaskChecks(X).CAID1
        End If

        If TaskChecks(X).CAM <> "" And TaskChecks(X).WP <> "" Then

            CAfound = False

            If camCount = 0 Then

                camCount = 1

                ReDim CAMChecks(1 To camCount)

                With CAMChecks(camCount)

                    .ID_str = tempID
                    .CAM_Test = TaskChecks(X).CAM
                    .CAM_Error = False

                End With

            Else

                For i = 1 To camCount

                    If CAMChecks(i).ID_str = tempID Then

                        CAfound = True

                        If TaskChecks(X).CAM <> CAMChecks(i).CAM_Test Then
                            CAMChecks(i).CAM_Error = True
                        End If

                        GoTo next_task

                    End If

                Next i

                If CAfound = False Then

                    camCount = camCount + 1

                    ReDim Preserve CAMChecks(1 To camCount)

                    With CAMChecks(camCount)

                        .ID_str = tempID
                        .CAM_Test = TaskChecks(X).CAM
                        .CAM_Error = False

                    End With

                End If

            End If

        End If

        If TaskChecks(X).WP <> "" Then

            wpFound = False

            If wpCount = 0 Then

                wpCount = 1

                ReDim WPChecks(1 To wpCount)

                With WPChecks(wpCount)

                    .ID_Test = tempID
                    .WP_ID = TaskChecks(X).WP
                    .EVT_Test = TaskChecks(X).EVT
                    .WP_DupError = False
                    .EVT_Error = False

                End With

            Else

                For i = 1 To wpCount

                    If WPChecks(i).WP_ID = TaskChecks(X).WP Then

                        wpFound = True

                        If tempID <> WPChecks(i).ID_Test Then

                            WPChecks(i).WP_DupError = True

                        End If

                        If TaskChecks(X).EVT <> WPChecks(i).EVT_Test Then

                            WPChecks(i).EVT_Error = True

                        End If

                        GoTo next_task

                    End If

                Next i

                If wpFound = False Then

                    wpCount = wpCount + 1

                    ReDim Preserve WPChecks(1 To wpCount)

                    With WPChecks(wpCount)

                        .ID_Test = tempID
                        .WP_ID = TaskChecks(X).WP
                        .EVT_Test = TaskChecks(X).EVT
                        .WP_DupError = False
                        .EVT_Error = False

                    End With

                End If

            End If

        End If

next_task:

        '**Report Tasks Missing Metadata**

        If TaskChecks(X).WP <> "" Or (TaskChecks(X).EVT <> "" And TaskChecks(X).EVT <> "NA" And TaskChecks(X).EVT <> "N/A") Or TaskChecks(X).BCost <> 0 Or TaskChecks(X).BWork <> 0 Then 'v3.2.2

            If TaskChecks(X).BWork = 0 Then tempBValue = TaskChecks(X).BCost Else tempBValue = TaskChecks(X).BWork

            If TaskChecks(X).WP = "" Or TaskChecks(X).EVT = "" Or TaskChecks(X).BStart = "NA" Or TaskChecks(X).BFinish = "NA" Or tempBValue = 0 Then

                ErrorCounter = ErrorCounter + 1

                If CAID3_Used = True And CAID2_Used = True Then

                    With TaskChecks(X)

                        errorStr = .UID & ","
                        If .CAID1 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID1 & ","
                        End If
                        If .CAID2 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID2 & ","
                        End If
                        If .CAID3 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID3 & ","
                        End If
                        If .CAM = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAM & ","
                        End If
                        If .WP = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .WP & ","
                        End If
                        If .EVT = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .EVT & ","
                        End If
                        If .BCost = 0 And .BWork = 0 Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            If .BWork > 0 Then
                                errorStr = errorStr & .BWork & ","
                            Else
                                errorStr = errorStr & .BCost & ","
                            End If
                        End If
                        If .BStart = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BStart & ","
                        End If
                        If .BFinish = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BFinish & ","
                        End If
                        If .EVT = "B" Or .EVT = "B Milestone" Then
                            If .MSID = "" Then
                                errorStr = errorStr & "MISSING,"
                            Else
                                errorStr = errorStr & .MSID & ","
                            End If

                            If .MSWeight = "" Then
                                errorStr = errorStr & "MISSING"
                            Else
                                errorStr = errorStr & .MSWeight
                            End If
                        Else
                            errorStr = errorStr & "N/A,N/A"
                        End If

                    End With
                End If

                If CAID3_Used = False And CAID2_Used = True Then

                    With TaskChecks(X)

                        errorStr = .UID & ","
                        If .CAID1 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID1 & ","
                        End If
                        If .CAID2 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID2 & ","
                        End If
                        If .CAM = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAM & ","
                        End If
                        If .WP = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .WP & ","
                        End If
                        If .EVT = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .EVT & ","
                        End If
                        If .BCost = 0 And .BWork = 0 Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            If .BWork > 0 Then
                                errorStr = errorStr & .BWork & ","
                            Else
                                errorStr = errorStr & .BCost & ","
                            End If
                        End If
                        If .BStart = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BStart & ","
                        End If
                        If .BFinish = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BFinish & ","
                        End If
                        If .EVT = "B" Or .EVT = "B Milestones" Then
                            If .MSID = "" Then
                                errorStr = errorStr & "MISSING,"
                            Else
                                errorStr = errorStr & .MSID & ","
                            End If

                            If .MSWeight = "" Then
                                errorStr = errorStr & "MISSING"
                            Else
                                errorStr = errorStr & .MSWeight
                            End If
                        Else
                            errorStr = errorStr & "N/A,N/A"
                        End If

                    End With

                End If

                If CAID3_Used = False And CAID2_Used = False Then

                    With TaskChecks(X)

                        errorStr = .UID & ","
                        If .CAID1 = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAID1 & ","
                        End If
                        If .CAM = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .CAM & ","
                        End If
                        If .WP = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .WP & ","
                        End If
                        If .EVT = "" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .EVT & ","
                        End If
                        If .BCost = 0 And .BWork = 0 Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            If .BWork > 0 Then
                                errorStr = errorStr & .BWork & ","
                            Else
                                errorStr = errorStr & .BCost & ","
                            End If
                        End If
                        If .BStart = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BStart & ","
                        End If
                        If .BFinish = "NA" Then
                            errorStr = errorStr & "MISSING,"
                        Else
                            errorStr = errorStr & .BFinish & ","
                        End If
                        If .EVT = "B" Or .EVT = "B Milestone" Then
                            If .MSID = "" Then
                                errorStr = errorStr & "MISSING,"
                            Else
                                errorStr = errorStr & .MSID & ","
                            End If

                            If .MSWeight = "" Then
                                errorStr = errorStr & "MISSING"
                            Else
                                errorStr = errorStr & .MSWeight
                            End If
                        Else
                            errorStr = errorStr & "N/A,N/A"
                        End If

                    End With

                End If

                Print #1, errorStr

                errorStr = ""

            End If

        End If

    Next X

    Print #1, vbCrLf & "Total Task Errors Found: " & ErrorCounter

    '**Report Multiple CAM Assignments**

    Print #1, vbCrLf & vbCrLf & "CAM Errors - The following items reflect multiple CAM assignments per Control Account as interpreted by Cobra (based on a unique CA record ID constructed from Concatenated CA ID Values"

    Print #1, vbCrLf & "CA ID String"

    ErrorCounter = 0

    For X = 1 To camCount

        If CAMChecks(X).CAM_Error = True Then

            ErrorCounter = ErrorCounter + 1

            Print #1, CAMChecks(X).ID_str

        End If

    Next X

    Print #1, vbCrLf & "Total CAM Errors Found: " & ErrorCounter

    '**Report Duplicate WP IDs & Multiple EVTs**

    Print #1, vbCrLf & vbCrLf & "Work Package Errors - The following Work Package IDs are duplicated across multiple CA ID values and/or are assigned multiple EVTs"

    Print #1, vbCrLf & "Work Package,Duplicate WP ID,Multiple EVTs"

    ErrorCounter = 0

    For X = 1 To wpCount

        If WPChecks(X).WP_DupError = True Or WPChecks(X).EVT_Error = True Then

            ErrorCounter = ErrorCounter + 1

            With WPChecks(X)
                errorStr = .WP_ID & "," & .WP_DupError & "," & .EVT_Error
            End With

            Print #1, errorStr

            errorStr = ""
        End If

    Next X

    Print #1, vbCrLf & "Total Work Package Errors Found: " & ErrorCounter

    '**Reporting Assignment Baseline Issues (Values and Dates)**

    Print #1, vbCrLf & vbCrLf & "Task Assignment Baseline Discrepancies - The following Tasks have vertical traceability errors with their Assignment Baseline Values and/or Baseline Dates. Note that tasks with mixed elements of cost (labor & material/cost) that include labor rates may result in false positives." 'v3.3.0

    Print #1, vbCrLf & "UID,Task Baseline Work,Assignment Baseline Work,Assignment TimeScaled Work,Task Baseline Cost,Assignment Baseline Cost,Assignment TimeScaled Cost,Task Baseline Start,Assignment Baseline Start,Task Baseline Finish,Assignment Baseline Finish, Assignment Count" 'v3.3.2

    ErrorCounter = 0

    For X = 1 To taskCount

        With TaskChecks(X)

            If .AssignmentCount > 0 Then 'v3.2.3

                If Round(CDec(.AssignmentTSVWork), 2) <> Round(CDec(.AssignmentBWork), 2) Or (Round(CDec(.AssignmentTSVCost), 2) <> Round(CDec(.AssignmentBCost), 2) And .AssignmentBWork = 0) Or Round(CDec(.BCost), 2) <> Round(CDec(.AssignmentBCost), 2) Or Round(CDec(.BWork), 2) <> Round(CDec(.AssignmentBWork), 2) Or .BStart <> .AssignmentBStart Or .BFinish <> .AssignmentBFinish Then 'v3.3.2

                    ErrorCounter = ErrorCounter + 1

                    errorStr = .UID & ","
                    errorStr = errorStr & .BWork & ","
                    errorStr = errorStr & .AssignmentBWork & ","
                    errorStr = errorStr & .AssignmentTSVWork & "," 'v3.3.2
                    errorStr = errorStr & .BCost & ","
                    errorStr = errorStr & .AssignmentBCost & ","
                    errorStr = errorStr & .AssignmentTSVCost & "," 'v3.3.2
                    errorStr = errorStr & .BStart & ","
                    errorStr = errorStr & .AssignmentBStart & ","
                    errorStr = errorStr & .BFinish & ","
                    errorStr = errorStr & .AssignmentBFinish & "," 'v3.2.3
                    errorStr = errorStr & .AssignmentCount 'v3.2.3

                    Print #1, errorStr
                    errorStr = ""

                End If

            End If
            
            If .AssignmentCount = 0 And (.BCost <> 0 Or .BWork <> 0) Then 'v3.2.3
            
                ErrorCounter = ErrorCounter + 1
                
                errorStr = .UID & ","
                errorStr = errorStr & .BWork & ","
                errorStr = errorStr & .AssignmentBWork & ","
                errorStr = errorStr & .AssignmentTSVWork & "," 'v3.3.2
                errorStr = errorStr & .BCost & ","
                errorStr = errorStr & .AssignmentBCost & ","
                errorStr = errorStr & .BStart & ","
                errorStr = errorStr & .AssignmentBStart & ","
                errorStr = errorStr & .AssignmentTSVCost & "," 'v3.3.2
                errorStr = errorStr & .BFinish & ","
                errorStr = errorStr & .AssignmentBFinish & "," 'v3.2.3
                errorStr = errorStr & .AssignmentCount 'v3.2.3
            
                Print #1, errorStr
                errorStr = ""
            
            End If

        End With

    Next X

    Print #1, vbCrLf & "Total Task Assignment Baseline Errors Found: " & ErrorCounter 'v3.3.0

    '**Reporting Assignment Forecast Issues (Values and Dates)**

    Print #1, vbCrLf & vbCrLf & "Task Assignment Forecast Discrepancies - The following Tasks have potential vertical traceability errors with their Assignment Forecast Values and/or Forecast Dates"

    Print #1, vbCrLf & "UID,Task Work,Assignment Work,Task Cost,Assignment Cost,Task Start,Assignment Start,Task Finish,Assignment Finish, Assignment Count"

    ErrorCounter = 0

    For X = 1 To taskCount

        With TaskChecks(X)

            If .AssignmentCount > 0 Then

                If Round(CDec(.FCost), 2) <> Round(CDec(.AssignmentFCost), 2) Or Round(CDec(.FWork), 2) <> Round(CDec(.AssignmentFWork), 2) Or .FStart <> .AssignmentFStart Or .FFinish <> .AssignmentFFinish Then

                    ErrorCounter = ErrorCounter + 1

                    errorStr = .UID & ","
                    errorStr = errorStr & .FWork & ","
                    errorStr = errorStr & .AssignmentFWork & ","
                    errorStr = errorStr & .FCost & ","
                    errorStr = errorStr & .AssignmentFCost & ","
                    errorStr = errorStr & .FStart & ","
                    errorStr = errorStr & .AssignmentFStart & ","
                    errorStr = errorStr & .FFinish & ","
                    errorStr = errorStr & .AssignmentFFinish & ","
                    errorStr = errorStr & .AssignmentCount

                    Print #1, errorStr
                    errorStr = ""

                End If

            End If
            
            If .AssignmentCount = 0 And (.FCost <> 0 Or .FWork <> 0) Then
            
                ErrorCounter = ErrorCounter + 1
                
                errorStr = .UID & ","
                errorStr = errorStr & .FWork & ","
                errorStr = errorStr & .AssignmentFWork & ","
                errorStr = errorStr & .FCost & ","
                errorStr = errorStr & .AssignmentFCost & ","
                errorStr = errorStr & .FStart & ","
                errorStr = errorStr & .AssignmentFStart & ","
                errorStr = errorStr & .FFinish & ","
                errorStr = errorStr & .AssignmentFFinish & ","
                errorStr = errorStr & .AssignmentCount
            
                Print #1, errorStr
                errorStr = ""
            
            End If

        End With

    Next X
    
    Print #1, vbCrLf & "Total Task Assignment Forecast Errors Found: " & ErrorCounter

    MsgBox "Data Check Report saved to " & destFolder

    Shell "explorer.exe" & " " & destFolder, vbNormalFocus

    Close #1

End Sub

Private Sub MPP_Export(ByVal curproj As Project)

    Dim subProj As SubProject
    Dim subProjs As Subprojects

    destFolder = SetDirectory(curproj.ProjectSummaryTask.Name)

    If curproj.Subprojects.Count > 0 Then

        Set subProjs = curproj.Subprojects

        For Each subProj In subProjs

            subProj.SourceProject.SaveAs Name:=destFolder & "\" & subProj.SourceProject.Name
            curproj.Subprojects(subProj.Index).SourceProject = destFolder & "\" & subProj.SourceProject.Name

        Next subProj

        curproj.SaveAs Name:=destFolder & "\" & curproj.ProjectSummaryTask.Name

    Else

        curproj.SaveAs Name:=destFolder & "\" & curproj.ProjectSummaryTask.Name

    End If

End Sub
Private Sub XML_Export(ByVal curproj As Project)

    Dim subProj As SubProject
    Dim subProjs As Subprojects

    destFolder = SetDirectory(curproj.ProjectSummaryTask.Name)

    If curproj.Subprojects.Count > 0 Then

        Set subProjs = curproj.Subprojects

        For Each subProj In subProjs

            subProj.SourceProject.SaveAs Name:=destFolder & "\" & subProj.SourceProject.Name, FormatID:="MSProject.XML"

        Next subProj


    Else

        curproj.SaveAs Name:=destFolder & "\" & curproj.ProjectSummaryTask.Name, FormatID:="MSProject.XML"

    End If

End Sub

Private Sub CSV_Export(ByVal curproj As Project)

    Dim docProps As DocumentProperties

    Set docProps = curproj.CustomDocumentProperties

    fCAID1 = docProps("fCAID1").Value
    fCAID1t = docProps("fCAID1t").Value
    If BCRxport = True Then
        fBCR = docProps("fBCR").Value
    End If
    If WhatIfxport = True Then 'v3.2
        fWhatIf = docProps("fWhatIf").Value
    End If
    If CAID3_Used = True Then
        fCAID3 = docProps("fCAID3").Value
        fCAID3t = docProps("fCAID3t").Value
    End If
    If CAID2_Used = True Then
        fCAID2 = docProps("fCAID2").Value
        fCAID2t = docProps("fCAID2t").Value
    End If
    fWP = docProps("fWP").Value
    fCAM = docProps("fCAM").Value
    fEVT = docProps("fEVT").Value
    If Milestones_Used Then
        fMilestone = docProps("fMSID").Value
        fMilestoneWeight = docProps("fMSW").Value
    End If
    fPCNT = docProps("fPCNT").Value
    If AssignmentPCNT_Used = True Then 'v3.3.2
        fAssignPcnt = docProps("fAssignPcnt").Value 'v3.3.0
    End If
    fResID = docProps("fResID").Value

    BCR_Error = False

    destFolder = SetDirectory(curproj.ProjectSummaryTask.Name)

    '*******************
    '****BCR Review*****
    '*******************

    If BCWSxport = True And BCRxport = True Then
        If Find_BCRs(curproj, fWP, fBCR, BCR_ID) = 0 Then
            MsgBox "BCR ID " & Chr(34) & BCR_ID & Chr(34) & " was not found in the IMS." & vbCr & vbCr & "Please try again."
            BCR_Error = True
            GoTo BCR_Error
        End If
    End If

    '*******************
    '****BCWS Export****
    '*******************

    If BCWSxport = True Then

        BCWS_Export curproj

    End If

    '*******************
    '****ETC Export****
    '*******************

    If ETCxport = True Then

        ETC_Export curproj

    End If

    '*******************
    '****BCWP Export****
    '*******************

    If BCWPxport = True Then

        BCWP_Export curproj

    End If
    
    '*******************
    '**What-if Export***
    '*******************

    If WhatIfxport = True Then 'v3.2

        WhatIf_Export curproj

    End If

    Exit Sub

BCR_Error:

    If BCR_Error = True Then
        DeleteDirectory (destFolder)
    End If

End Sub

Private Sub BCWP_Export(ByVal curproj As Project)

    '*******************
    '****BCWP Export****
    '*******************

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, ResName, MSWeight, ID, PCNT As String 'v3.3.0
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\BCWP ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete"
        End If

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    Err.Raise 1
                                End If

                                If EVT = "B" Or EVT = "N" Or EVT = "B Milestone" Or EVT = "N Earning Rules" Then

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If

                                ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = 1
                                            ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        Else
                                            ACTarray(X).sumBCWS = 1
                                            ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        End If
                                        ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
                                        ActFound = True

                                        GoTo nrBCWP_WP_Match_A

                                    End If

                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineWork <> 0 Then
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            Else
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            End If

                                            ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100

                                            GoTo nrBCWP_WP_Match_A
                                        End If
                                    Next i

                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(i).sumBCWS = 1
                                        ACTarray(i).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(i).sumBCWS = 1
                                        ACTarray(i).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    ActFound = True

                                ElseIf EVT = "E" Or EVT = "F" Or EVT = "G" Or EVT = "H" Or EVT = "E 50/50" Or EVT = "F 0/100" Or EVT = "G 100/0" Or EVT = "H User Defined" Then

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If

                                        X = X + 1
                                        ActFound = True

                                        GoTo nrBCWP_WP_Match_A

                                    End If

                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If

                                            GoTo nrBCWP_WP_Match_A
                                        End If
                                    Next i

                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
                                    ActFound = True

                                End If

                            End If

                        End If

                    End If

nrBCWP_WP_Match_A:

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            End If
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                Err.Raise 1
                            End If

                            If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earning Rules" Then

                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If

                            ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = 1
                                        ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(X).sumBCWS = 1
                                        ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    ActFound = True

                                    GoTo nrBCWP_WP_Match_B

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        Else
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + 1
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        End If
                                        ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100

                                        GoTo nrBCWP_WP_Match_B
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If
                                If t.BaselineWork <> 0 Then
                                    ACTarray(X).sumBCWS = 1
                                    ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                Else
                                    ACTarray(X).sumBCWS = 1
                                    ACTarray(X).sumBCWP = 1 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                End If
                                ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                X = X + 1
                                ActFound = True

                            ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or EVT = "H" Or EVT = "H User Defined" Then

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo nrBCWP_WP_Match_B

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If

                                        GoTo nrBCWP_WP_Match_B
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
                                ActFound = True

                            End If

                        End If

                    End If

                End If

nrBCWP_WP_Match_B:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")

                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog
                End If

            Next i
        End If

        Close #1

    Else '**Resource Loaded**

        ACTfilename = destFolder & "\BCWP ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete,Resource" 'v3.3.0
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete,Resource" 'v3.3.0
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",WP,Milestone,Forecast Start Date,Forecast Finish Date,Actual Start Date,Actual Finish Date,Percent Complete,Resource" 'v3.3.0
        End If

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                ResName = "" 'v3.3.0

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    Err.Raise 1
                                End If
                                

                                If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earned Rules" Then

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                    End If
    
                                ElseIf EVT = "L" Or EVT = "L Assignment % Complete" Then 'v3.3.0
                                
                                    'store ACT info
                                    'WP Data
                                    
                                    Set tAss = t.Assignments
                                        
                                    For Each tAssign In tAss
                                    
                                        ResName = tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource))
                                        ID = ID & "/" & ResName
                                        
                                        If X = 1 Then
    
                                            'create new WP line in ACTarrray
                                            ReDim ACTarray(1 To X)
                                            If CAID3_Used = True Then
                                                ACTarray(X).CAID3 = CAID3
                                            End If
                                            ACTarray(X).CAM = CAM
                                            ACTarray(X).Resource = ResName
                                            ACTarray(X).ID = ID
                                            ACTarray(X).CAID1 = CAID1
                                            ACTarray(X).EVT = EVT
                                            If CAID2_Used = True Then
                                                ACTarray(X).CAID2 = CAID2
                                            End If
                                            ACTarray(X).WP = WP
                                            ACTarray(X).FFinish = tAssign.Finish
                                            ACTarray(X).FStart = tAssign.Start
                                            If tAssign.ActualStart <> "NA" Then ACTarray(X).AStart = tAssign.ActualStart
                                            If tAssign.ActualFinish <> "NA" Then ACTarray(X).AFinish = tAssign.ActualFinish
                                            
                                            If tAssign.BaselineWork <> 0 Then
                                                ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                                ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                            Else
                                                ACTarray(X).sumBCWS = tAssign.BaselineCost
                                                ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                            End If
                                            
                                            If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100
    
                                            X = X + 1
                                            ActFound = True
    
                                            GoTo Next_Assign_A
    
                                        End If
    
                                        For i = 1 To UBound(ACTarray)
                                            If ACTarray(i).ID = ID Then
                                                'Found an existing matching WP line
                                                If ACTarray(i).FStart > tAssign.Start Then
                                                    ACTarray(i).FStart = tAssign.Start
                                                End If
                                                If ACTarray(i).FFinish < tAssign.Finish Then
                                                    ACTarray(i).FFinish = tAssign.Finish
                                                End If
                                                If tAssign.ActualStart <> "NA" Then
                                                    If ACTarray(i).AStart = 0 Then
                                                        ACTarray(i).AStart = tAssign.ActualStart
                                                    Else
                                                        If tAssign.ActualStart < ACTarray(i).AStart Then
                                                            ACTarray(i).AStart = tAssign.ActualStart
                                                        End If
                                                    End If
                                                End If
                                                If tAssign.ActualFinish <> "NA" Then
                                                    If ACTarray(i).AFinish = 0 Then
                                                        ACTarray(i).AFinish = tAssign.ActualFinish
                                                    Else
                                                        If tAssign.ActualFinish > ACTarray(i).AFinish Then
                                                            ACTarray(i).AFinish = tAssign.ActualFinish
                                                        End If
                                                    End If
                                                End If
                                                If tAssign.BaselineWork <> 0 Then
                                                    ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineWork / 60
                                                    ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                                Else
                                                    ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineCost
                                                    ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                                End If
    
                                                If ACTarray(i).sumBCWS > 0 Then ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100
    
                                                GoTo Next_Assign_A
                                            End If
                                        Next i
    
                                        'No match found, create new WP line in ACTarrray
                                        ReDim Preserve ACTarray(1 To X)
                                        
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).Resource = ResName
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = tAssign.Finish
                                        ACTarray(X).FStart = tAssign.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        ACTarray(X).WP = WP
                                        If tAssign.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = tAssign.ActualStart
                                        End If
                                        If tAssign.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = tAssign.ActualFinish
                                        End If
                                        If tAssign.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                            ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        Else
                                            ACTarray(X).sumBCWS = tAssign.BaselineCost
                                            ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        End If
                                        
                                        If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
                                        
                                        ActFound = True
                                        
Next_Assign_A:

                                    Next tAssign
                                
                                ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = t.BaselineWork / 60
                                            ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        Else
                                            ACTarray(X).sumBCWS = t.BaselineCost
                                            ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                        End If
                                        ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
                                        ActFound = True

                                        GoTo BCWP_WP_Match_A

                                    End If

                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineWork <> 0 Then
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineWork / 60
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            Else
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineCost
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                            End If

                                            ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100

                                            GoTo BCWP_WP_Match_A
                                        End If
                                    Next i

                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = t.BaselineWork / 60 'v3.3.0
                                        ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100 'v3.3.0
                                    Else
                                        ACTarray(X).sumBCWS = t.BaselineCost
                                        ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100 'v3.3.0
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    ActFound = True

                                ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or EVT = "H" Or EVT = "H User Defined" Then

                                    'store ACT info
                                    'WP Data
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If t.BaselineFinish <> "NA" Then
                                            ACTarray(X).BFinish = t.BaselineFinish
                                        End If
                                        If t.BaselineStart <> "NA" Then
                                            ACTarray(X).BStart = t.BaselineStart
                                        End If
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        ACTarray(X).FFinish = t.Finish
                                        ACTarray(X).FStart = t.Start
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        ACTarray(X).WP = WP
                                        If t.ActualStart <> "NA" Then
                                            ACTarray(X).AStart = t.ActualStart
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            ACTarray(X).AFinish = t.ActualFinish
                                        End If

                                        X = X + 1
                                        ActFound = True

                                        GoTo BCWP_WP_Match_A

                                    End If

                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If t.BaselineStart <> "NA" Then
                                                If ACTarray(i).BStart = 0 Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                Else
                                                    If ACTarray(i).BStart > t.BaselineStart Then
                                                        ACTarray(i).BStart = t.BaselineStart
                                                    End If
                                                End If
                                            End If
                                            If t.BaselineFinish <> "NA" Then
                                                If ACTarray(i).BFinish = 0 Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                Else
                                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                                        ACTarray(i).BFinish = t.BaselineFinish
                                                    End If
                                                End If
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                            If t.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                Else
                                                    If t.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = t.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If t.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                Else
                                                    If t.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = t.ActualFinish
                                                    End If
                                                End If
                                            End If

                                            GoTo BCWP_WP_Match_A
                                        End If
                                    Next i

                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
                                    ActFound = True

                                End If

                            End If

                        End If

                    End If

BCWP_WP_Match_A:

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                            End If
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            ResName = "" 'v3.3.0

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                Err.Raise 1
                            End If

                            If EVT = "B" Or EVT = "B Milestone" Or EVT = "N" Or EVT = "N Earned Rules" Then

                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & WP & "," & UID & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & "," & Format(t.ActualStart, "M/D/YYYY") & "," & Format(t.ActualFinish, "M/D/YYYY") & "," & PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT)))
                                End If
                                
                            ElseIf EVT = "L" Or EVT = "L Assignment % Complete" Then 'v3.3.0
                                
                                'store ACT info
                                'WP Data
                                
                                Set tAss = t.Assignments
                                    
                                For Each tAssign In tAss
                                
                                    ResName = tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource))
                                    ID = ID & "/" & ResName
                                    
                                    If X = 1 Then

                                        'create new WP line in ACTarrray
                                        ReDim ACTarray(1 To X)
                                        If CAID3_Used = True Then
                                            ACTarray(X).CAID3 = CAID3
                                        End If
                                        ACTarray(X).CAM = CAM
                                        ACTarray(X).Resource = ResName
                                        ACTarray(X).ID = ID
                                        ACTarray(X).CAID1 = CAID1
                                        ACTarray(X).EVT = EVT
                                        If CAID2_Used = True Then
                                            ACTarray(X).CAID2 = CAID2
                                        End If
                                        ACTarray(X).WP = WP
                                        ACTarray(X).FFinish = tAssign.Finish
                                        ACTarray(X).FStart = tAssign.Start
                                        If tAssign.ActualStart <> "NA" Then ACTarray(X).AStart = tAssign.ActualStart
                                        If tAssign.ActualFinish <> "NA" Then ACTarray(X).AFinish = tAssign.ActualFinish
                                        
                                        If tAssign.BaselineWork <> 0 Then
                                            ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                            ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        Else
                                            ACTarray(X).sumBCWS = tAssign.BaselineCost
                                            ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                        End If
                                        
                                        If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                        X = X + 1
                                        
                                        ActFound = True

                                        GoTo Next_Assign_B

                                    End If

                                    For i = 1 To UBound(ACTarray)
                                        If ACTarray(i).ID = ID Then
                                            'Found an existing matching WP line
                                            If ACTarray(i).FStart > tAssign.Start Then
                                                ACTarray(i).FStart = tAssign.Start
                                            End If
                                            If ACTarray(i).FFinish < tAssign.Finish Then
                                                ACTarray(i).FFinish = tAssign.Finish
                                            End If
                                            If tAssign.ActualStart <> "NA" Then
                                                If ACTarray(i).AStart = 0 Then
                                                    ACTarray(i).AStart = tAssign.ActualStart
                                                Else
                                                    If tAssign.ActualStart < ACTarray(i).AStart Then
                                                        ACTarray(i).AStart = tAssign.ActualStart
                                                    End If
                                                End If
                                            End If
                                            If tAssign.ActualFinish <> "NA" Then
                                                If ACTarray(i).AFinish = 0 Then
                                                    ACTarray(i).AFinish = tAssign.ActualFinish
                                                Else
                                                    If tAssign.ActualFinish > ACTarray(i).AFinish Then
                                                        ACTarray(i).AFinish = tAssign.ActualFinish
                                                    End If
                                                End If
                                            End If
                                            If tAssign.BaselineWork <> 0 Then
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineWork / 60
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                            Else
                                                ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + tAssign.BaselineCost
                                                ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100)
                                            End If

                                            If ACTarray(i).sumBCWS > 0 Then ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100

                                            GoTo Next_Assign_B
                                            
                                        End If
                                    Next i

                                    'No match found, create new WP line in ACTarrray
                                    ReDim Preserve ACTarray(1 To X)
                                    
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).Resource = ResName
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = tAssign.Finish
                                    ACTarray(X).FStart = tAssign.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If tAssign.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = tAssign.ActualStart
                                    End If
                                    If tAssign.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = tAssign.ActualFinish
                                    End If
                                    If tAssign.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = tAssign.BaselineWork / 60
                                        ACTarray(X).sumBCWP = tAssign.BaselineWork / 60 * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                    Else
                                        ACTarray(X).sumBCWS = tAssign.BaselineCost
                                        ACTarray(X).sumBCWP = tAssign.BaselineCost * PercentfromString(get_Assignment_Pcnt(tAssign)) / 100
                                    End If
                                    
                                    If ACTarray(X).sumBCWS > 0 Then ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    
                                    ActFound = True
                                    
Next_Assign_B:
                                
                                Next tAssign

                            ElseIf EVT = "C" Or EVT = "C % Work Complete" Then

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If
                                    If t.BaselineWork <> 0 Then
                                        ACTarray(X).sumBCWS = t.BaselineWork / 60
                                        ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    Else
                                        ACTarray(X).sumBCWS = t.BaselineCost
                                        ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                    End If
                                    ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                    X = X + 1
                                    ActFound = True

                                    GoTo BCWP_WP_Match_B

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        If t.BaselineWork <> 0 Then
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineWork / 60
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        Else
                                            ACTarray(i).sumBCWS = ACTarray(i).sumBCWS + t.BaselineCost
                                            ACTarray(i).sumBCWP = ACTarray(i).sumBCWP + (t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100)
                                        End If
                                        ACTarray(i).Prog = ACTarray(i).sumBCWP / ACTarray(i).sumBCWS * 100

                                        GoTo BCWP_WP_Match_B
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If
                                If t.BaselineWork <> 0 Then
                                    ACTarray(X).sumBCWS = t.BaselineWork / 60
                                    ACTarray(X).sumBCWP = t.BaselineWork / 60 * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                Else
                                    ACTarray(X).sumBCWS = t.BaselineCost
                                    ACTarray(X).sumBCWP = t.BaselineCost * PercentfromString(t.GetField(FieldNameToFieldConstant(fPCNT))) / 100
                                End If
                                ACTarray(X).Prog = ACTarray(X).sumBCWP / ACTarray(X).sumBCWS * 100

                                X = X + 1
                                ActFound = True

                            ElseIf EVT = "E" Or EVT = "E 50/50" Or EVT = "F" Or EVT = "F 0/100" Or EVT = "G" Or EVT = "G 100/0" Or EVT = "H" Or EVT = "H User Defined" Then

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo BCWP_WP_Match_B

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If

                                        GoTo BCWP_WP_Match_B
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
                                ActFound = True

                            End If

                        End If

                    End If

                End If

BCWP_WP_Match_B:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")

                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog & "," & ACTarray(i).Resource 'v3.3.0
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog & "," & ACTarray(i).Resource 'v3.3.0
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).WP & "," & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY") & "," & aStartString & "," & aFinishString & "," & ACTarray(i).Prog & "," & ACTarray(i).Resource 'v3.3.0
                End If

            Next i
        End If

        Close #1

    End If

End Sub

Private Sub ETC_Export(ByVal curproj As Project)

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT As String
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String

    '*******************
    '****ETC Export****
    '*******************

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\ETC ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo nrETC_WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        GoTo nrETC_WP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).ID = ID
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
                                ActFound = True

                                'Milestone Data
nrETC_WP_Match:



                            End If

                        End If

                    End If

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).ID = ID
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
                                ActFound = True

                                GoTo nrETC_WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.BaselineStart <> "NA" Then
                                        If ACTarray(i).BStart = 0 Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        Else
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                        End If
                                    End If
                                    If t.BaselineFinish <> "NA" Then
                                        If ACTarray(i).BFinish = 0 Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        Else
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                        End If
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    If t.ActualStart <> "NA" Then
                                        If ACTarray(i).AStart = 0 Then
                                            ACTarray(i).AStart = t.ActualStart
                                        Else
                                            If t.ActualStart < ACTarray(i).AStart Then
                                                ACTarray(i).AStart = t.ActualStart
                                            End If
                                        End If
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        If ACTarray(i).AFinish = 0 Then
                                            ACTarray(i).AFinish = t.ActualFinish
                                        Else
                                            If t.ActualFinish > ACTarray(i).AFinish Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            End If
                                        End If
                                    End If

                                    GoTo nrETC_WP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            If t.BaselineFinish <> "NA" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                            End If
                            If t.BaselineStart <> "NA" Then
                                ACTarray(X).BStart = t.BaselineStart
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).ID = ID
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            ACTarray(X).WP = WP
                            If t.ActualStart <> "NA" Then
                                ACTarray(X).AStart = t.ActualStart
                            End If
                            If t.ActualFinish <> "NA" Then
                                ACTarray(X).AFinish = t.ActualFinish
                            End If

                            X = X + 1
                            ActFound = True

nrETC_WP_Match_B:

                        End If

                    End If

                End If

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")

                If aFinishString = "NA" Then
                    If CAID3_Used = True And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                End If

            Next i
        End If

        Close #1

    Else '**Resource Loaded**

        ACTfilename = destFolder & "\ETC ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\ETC RES_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Forecast Start Date,Forecast Finish Date"
        End If
        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.Work > 0 Or t.Cost > 0 Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.BaselineFinish <> "NA" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                    End If
                                    If t.BaselineStart <> "NA" Then
                                        ACTarray(X).BStart = t.BaselineStart
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP
                                    If t.ActualStart <> "NA" Then
                                        ACTarray(X).AStart = t.ActualStart
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        ACTarray(X).AFinish = t.ActualFinish
                                    End If

                                    X = X + 1
                                    ActFound = True

                                    GoTo ETC_WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.BaselineStart <> "NA" Then
                                            If ACTarray(i).BStart = 0 Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            Else
                                                If ACTarray(i).BStart > t.BaselineStart Then
                                                    ACTarray(i).BStart = t.BaselineStart
                                                End If
                                            End If
                                        End If
                                        If t.BaselineFinish <> "NA" Then
                                            If ACTarray(i).BFinish = 0 Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            Else
                                                If ACTarray(i).BFinish < t.BaselineFinish Then
                                                    ACTarray(i).BFinish = t.BaselineFinish
                                                End If
                                            End If
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                        If t.ActualStart <> "NA" Then
                                            If ACTarray(i).AStart = 0 Then
                                                ACTarray(i).AStart = t.ActualStart
                                            Else
                                                If t.ActualStart < ACTarray(i).AStart Then
                                                    ACTarray(i).AStart = t.ActualStart
                                                End If
                                            End If
                                        End If
                                        If t.ActualFinish <> "NA" Then
                                            If ACTarray(i).AFinish = 0 Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            Else
                                                If t.ActualFinish > ACTarray(i).AFinish Then
                                                    ACTarray(i).AFinish = t.ActualFinish
                                                End If
                                            End If
                                        End If
                                        GoTo ETC_WP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).ID = ID
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
                                ActFound = True

                                'Milestone Data
ETC_WP_Match:


                                Set tAss = t.Assignments

                                For Each tAssign In tAss

                                    If TimeScaleExport = True Then

                                        ExportTimeScaleResources ID, t, tAssign, 2, "ETC"

                                    Else

                                        Select Case tAssign.ResourceType

                                            Case pjResourceTypeWork

                                            If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                            ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                            End If

                                        Case pjResourceTypeCost

                                            If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingCost & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                            ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                            End If

                                        Case pjResourceTypeMaterial

                                            If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                            ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                            End If

                                        End Select

                                    End If

                                Next tAssign

                            End If

                        End If

                    End If

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If t.Work > 0 Or t.Cost > 0 Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.BaselineFinish <> "NA" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                End If
                                If t.BaselineStart <> "NA" Then
                                    ACTarray(X).BStart = t.BaselineStart
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).CAM = CAM
                                ACTarray(X).ID = ID
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP
                                If t.ActualStart <> "NA" Then
                                    ACTarray(X).AStart = t.ActualStart
                                End If
                                If t.ActualFinish <> "NA" Then
                                    ACTarray(X).AFinish = t.ActualFinish
                                End If

                                X = X + 1
                                ActFound = True

                                GoTo ETC_WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.BaselineStart <> "NA" Then
                                        If ACTarray(i).BStart = 0 Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        Else
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                        End If
                                    End If
                                    If t.BaselineFinish <> "NA" Then
                                        If ACTarray(i).BFinish = 0 Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        Else
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                        End If
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    If t.ActualStart <> "NA" Then
                                        If ACTarray(i).AStart = 0 Then
                                            ACTarray(i).AStart = t.ActualStart
                                        Else
                                            If t.ActualStart < ACTarray(i).AStart Then
                                                ACTarray(i).AStart = t.ActualStart
                                            End If
                                        End If
                                    End If
                                    If t.ActualFinish <> "NA" Then
                                        If ACTarray(i).AFinish = 0 Then
                                            ACTarray(i).AFinish = t.ActualFinish
                                        Else
                                            If t.ActualFinish > ACTarray(i).AFinish Then
                                                ACTarray(i).AFinish = t.ActualFinish
                                            End If
                                        End If
                                    End If

                                    GoTo ETC_WP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            If t.BaselineFinish <> "NA" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                            End If
                            If t.BaselineStart <> "NA" Then
                                ACTarray(X).BStart = t.BaselineStart
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).ID = ID
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            ACTarray(X).WP = WP
                            If t.ActualStart <> "NA" Then
                                ACTarray(X).AStart = t.ActualStart
                            End If
                            If t.ActualFinish <> "NA" Then
                                ACTarray(X).AFinish = t.ActualFinish
                            End If


                            X = X + 1
                            ActFound = True

                            'Milestone Data
ETC_WP_Match_B:


                            Set tAss = t.Assignments

                            For Each tAssign In tAss

                                If TimeScaleExport = True Then

                                    ExportTimeScaleResources ID, t, tAssign, 2, "ETC"

                                Else

                                    Select Case tAssign.ResourceType

                                        Case pjResourceTypeWork

                                        If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        End If

                                    Case pjResourceTypeCost

                                        If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingCost & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        End If

                                    Case pjResourceTypeMaterial

                                        If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.RemainingWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        End If

                                    End Select

                                End If

                            Next tAssign

                        End If

                    End If

                End If

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If ACTarray(i).AStart = 0 Then aStartString = "NA" Else aStartString = Format(ACTarray(i).AStart, "M/D/YYYY")
                If ACTarray(i).AFinish = 0 Or ACTarray(i).AFinish < ACTarray(i).FFinish Then aFinishString = "NA" Else aFinishString = Format(ACTarray(i).AFinish, "M/D/YYYY")

                If aFinishString = "NA" Then
                    If CAID3_Used = True And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & Format(ACTarray(i).FStart, "M/D/YYYY") & "," & Format(ACTarray(i).FFinish, "M/D/YYYY")
                    End If
                End If

            Next i
        End If

        Close #1
        Close #2

    End If

End Sub
Private Sub BCWS_Export(ByVal curproj As Project)

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT As String
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim WPDescArray() As WP_Descriptions
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String

    '*******************
    '****BCWS Export****
    '*******************

    If DescExport = True Then
        Get_WP_Descriptions curproj
    End If

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\BCWS ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))

                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If

                                If Milestones_Used = False Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_nrSProj_Task
                                    End If
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    Err.Raise 1
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    ACTarray(X).CAID2 = CAID2
                                    ACTarray(X).WP = WP

                                    X = X + 1
                                    ActFound = True

                                    GoTo nrWP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If

                                        GoTo nrWP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                'Milestone Data
nrWP_Match:

                                If EVT = "B" Or EVT = "B Milestone" Then

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If

                                End If

                            End If

                        End If

                    End If
Next_nrSProj_Task:

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                Err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo Next_nrTask
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                GoTo nrWP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If ACTarray(i).BStart > t.BaselineStart Then
                                        ACTarray(i).BStart = t.BaselineStart
                                    End If
                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                        ACTarray(i).BFinish = t.BaselineFinish
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    GoTo nrWP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            ACTarray(X).BFinish = t.BaselineFinish
                            ACTarray(X).BStart = t.BaselineStart
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            ACTarray(X).CAID2 = CAID2
                            ACTarray(X).WP = WP

                            X = X + 1
                            ActFound = True

                            'Milestone Data
nrWP_Match_B:

                            If EVT = "B" Or EVT = "B Milestone" Then
                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                            End If

                        End If

                    End If

                End If
Next_nrTask:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If

                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If

            Next i
        End If

        Close #1

    Else

        ACTfilename = destFolder & "\BCWS ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\BCWS RES_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If

                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    Err.Raise 1
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_SProj_Task
                                    End If
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP

                                    X = X + 1
                                    ActFound = True

                                    GoTo WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If

                                        GoTo WP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                'Milestone Data
WP_Match:

                                If EVT = "B" Or EVT = "B Milestone" Then

                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If

                                End If

                                Set tAss = t.Assignments

                                For Each tAssign In tAss

                                    If tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0 Then

                                        If TimeScaleExport = True Then

                                            ExportTimeScaleResources ID, t, tAssign, 2, "BCWS"

                                        Else

                                            Select Case tAssign.ResourceType

                                                Case pjResourceTypeWork

                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                                Case pjResourceTypeCost

                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                                Case pjResourceTypeMaterial

                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                            End Select

                                        End If

                                    End If

                                Next tAssign

                            End If

                        End If

                    End If
Next_SProj_Task:

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If t.BaselineWork > 0 Or t.BaselineCost > 0 Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                Err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo next_task
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                GoTo WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If ACTarray(i).BStart > t.BaselineStart Then
                                        ACTarray(i).BStart = t.BaselineStart
                                    End If
                                    If ACTarray(i).BFinish < t.BaselineFinish Then
                                        ACTarray(i).BFinish = t.BaselineFinish
                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    GoTo WP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            ACTarray(X).BFinish = t.BaselineFinish
                            ACTarray(X).BStart = t.BaselineStart
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            ACTarray(X).WP = WP

                            X = X + 1
                            ActFound = True

                            'Milestone Data
WP_Match_B:

                            If EVT = "B" Or EVT = "B Milestone" Then
                                If CAID3_Used = True And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                End If
                            End If

                            Set tAss = t.Assignments

                            For Each tAssign In tAss

                                If tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0 Then

                                    If TimeScaleExport = True Then

                                        ExportTimeScaleResources ID, t, tAssign, 2, "BCWS"

                                    Else

                                        Select Case tAssign.ResourceType

                                            Case pjResourceTypeWork

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                            Case pjResourceTypeCost

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                            Case pjResourceTypeMaterial

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                        End Select

                                    End If

                                End If

                            Next tAssign

                        End If

                    End If

                End If
next_task:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If

                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If

            Next i
        End If

        Close #1
        Close #2

    End If
        
End Sub

Private Sub WhatIf_Export(ByVal curproj As Project) 'v3.2

    Dim t As Task
    Dim tAss As Assignments
    Dim tAssign As Assignment
    Dim CAID1, CAID3, WP, CAM, EVT, UID, CAID2, MSWeight, ID, PCNT As String
    Dim Milestone As String
    Dim subProj As SubProject
    Dim subProjs As Subprojects
    Dim curSProj As Project
    Dim ACTarray() As ACTrowWP
    Dim WPDescArray() As WP_Descriptions
    Dim X As Integer
    Dim i As Integer
    Dim aStartString As String
    Dim aFinishString As String

    '*******************
    '**What-if Export***
    '*******************

    If DescExport = True Then
        Get_WP_Descriptions curproj
    End If

    If ResourceLoaded = False Then

        ACTfilename = destFolder & "\WhatIf ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                            If t.GetField(FieldNameToFieldConstant(fWP)) <> "" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))

                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If

                                If Milestones_Used = False Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_nrSProj_Task
                                    End If
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    Err.Raise 1
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                        ACTarray(X).BStart = t.BaselineStart
                                    Else
                                        ACTarray(X).BFinish = t.Finish
                                        ACTarray(X).BStart = t.Start
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    ACTarray(X).CAID2 = CAID2
                                    ACTarray(X).WP = WP

                                    X = X + 1
                                    ActFound = True

                                    GoTo nrWP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        
                                        'Found an existing matching WP line
                                        If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                        Else
                                            If ACTarray(i).BStart > t.Start Then
                                                ACTarray(i).BStart = t.Start
                                            End If
                                            If ACTarray(i).BFinish < t.Finish Then
                                                ACTarray(i).BFinish = t.Finish
                                            End If
                                            If ACTarray(i).FStart > t.Start Then
                                                ACTarray(i).FStart = t.Start
                                            End If
                                            If ACTarray(i).FFinish < t.Finish Then
                                                ACTarray(i).FFinish = t.Finish
                                            End If
                                        End If
                                        GoTo nrWP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                'Milestone Data
nrWP_Match:

                                If EVT = "B" Or EVT = "B Milestone" Then
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                        End If
                                    Else
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                        End If
                                    End If

                                End If

                            End If

                        End If

                    End If
Next_nrSProj_Task:

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then
                        If t.GetField(FieldNameToFieldConstant(fWP)) <> "" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                Err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo Next_nrTask
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                ACTarray(X).CAID2 = CAID2
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                GoTo nrWP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                    Else
                                        If ACTarray(i).BStart > t.Start Then
                                            ACTarray(i).BStart = t.Start
                                        End If
                                        If ACTarray(i).BFinish < t.Finish Then
                                            ACTarray(i).BFinish = t.Finish
                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If
                                    End If
                                    GoTo nrWP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                            Else
                                ACTarray(X).BFinish = t.Finish
                                ACTarray(X).BStart = t.Start
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            ACTarray(X).CAID2 = CAID2
                            ACTarray(X).WP = WP

                            X = X + 1
                            ActFound = True

                            'Milestone Data
nrWP_Match_B:

                            If EVT = "B" Or EVT = "B Milestone" Then
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                Else
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                    End If
                                End If
                            End If

                        End If

                    End If

                End If
Next_nrTask:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If

                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If

            Next i
        End If

        Close #1

    Else

        ACTfilename = destFolder & "\WhatIf ACT_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"
        RESfilename = destFolder & "\WhatIf RES_" & RemoveIllegalCharacters(curproj.ProjectSummaryTask.Name) & "_" & Format(Now, "YYYYMMDD HHMM") & ".csv"

        Open ACTfilename For Output As #1
        Open RESfilename For Output As #2

        If CAID3_Used = True And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID3t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = True Then
            Print #1, fCAID1t & "," & fCAID2t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        If CAID3_Used = False And CAID2_Used = False Then
            Print #1, fCAID1t & ",CAM,WP,ID,Milestone,Milestone Weight,Description,Baseline Start Date,Baseline Finish Date,Progress Technique"
        End If
        Print #2, "Cobra ID,Resource,Amount,From Date,To Date"

        X = 1
        ActFound = False

        If curproj.Subprojects.Count > 0 Then

            Set subProjs = curproj.Subprojects

            For Each subProj In subProjs

                FileOpen Name:=subProj.Path, ReadOnly:=True
                Set curSProj = ActiveProject

                For Each t In curSProj.Tasks

                    If Not t Is Nothing Then

                        If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then
                        
                            If ((t.BaselineWork > 0 Or t.BaselineCost > 0) And _
                            (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D")) _
                            Or _
                            ((t.Work > 0 Or t.Cost > 0) And _
                            (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r")) Then

                                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                                If CAID3_Used = True Then
                                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                                End If
                                WP = t.GetField(FieldNameToFieldConstant(fWP))
                                EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                                If CAID2_Used = True Then
                                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                                End If
                                CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                                If CAID3_Used = True And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = True Then
                                    ID = CAID1 & "/" & CAID2 & "/" & WP
                                End If
                                If CAID3_Used = False And CAID2_Used = False Then
                                    ID = CAID1 & "/" & WP
                                End If

                                If Milestones_Used = True Then
                                    UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                    MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                                End If

                                If EVT = "B" And Milestones_Used = False Then
                                    ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                    Err.Raise 1
                                End If

                                If BCRxport = True Then
                                    If IsInArray(WP, BCR_WP) = False Then
                                        GoTo Next_SProj_Task
                                    End If
                                End If

                                'store ACT info
                                'WP Data
                                If X = 1 Then

                                    'create new WP line in ACTarrray
                                    ReDim ACTarray(1 To X)
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        ACTarray(X).BFinish = t.BaselineFinish
                                        ACTarray(X).BStart = t.BaselineStart
                                    Else
                                        ACTarray(X).BFinish = t.Finish
                                        ACTarray(X).BStart = t.Start
                                    End If
                                    If CAID3_Used = True Then
                                        ACTarray(X).CAID3 = CAID3
                                    End If
                                    ACTarray(X).ID = ID
                                    ACTarray(X).CAM = CAM
                                    ACTarray(X).CAID1 = CAID1
                                    ACTarray(X).EVT = EVT
                                    ACTarray(X).FFinish = t.Finish
                                    ACTarray(X).FStart = t.Start
                                    If CAID2_Used = True Then
                                        ACTarray(X).CAID2 = CAID2
                                    End If
                                    ACTarray(X).WP = WP

                                    X = X + 1
                                    ActFound = True

                                    GoTo WP_Match

                                End If

                                For i = 1 To UBound(ACTarray)
                                    If ACTarray(i).ID = ID Then
                                        'Found an existing matching WP line
                                        If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                            If ACTarray(i).BStart > t.BaselineStart Then
                                                ACTarray(i).BStart = t.BaselineStart
                                            End If
                                            If ACTarray(i).BFinish < t.BaselineFinish Then
                                                ACTarray(i).BFinish = t.BaselineFinish
                                            End If
                                        Else
                                            If ACTarray(i).BStart > t.Start Then
                                                ACTarray(i).BStart = t.Start
                                            End If
                                            If ACTarray(i).BFinish < t.Finish Then
                                                ACTarray(i).BFinish = t.Finish
                                            End If

                                        End If
                                        If ACTarray(i).FStart > t.Start Then
                                            ACTarray(i).FStart = t.Start
                                        End If
                                        If ACTarray(i).FFinish < t.Finish Then
                                            ACTarray(i).FFinish = t.Finish
                                        End If

                                        GoTo WP_Match
                                    End If
                                Next i

                                'No match found, create new WP line in ACTarrray
                                ReDim Preserve ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                'Milestone Data
WP_Match:

                                If EVT = "B" Or EVT = "B Milestone" Then
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                        End If
                                        
                                    Else
                                        
                                        If CAID3_Used = True And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = True Then
                                            Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                        End If
                                        If CAID3_Used = False And CAID2_Used = False Then
                                            Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                        End If
                                        
                                    End If
                                End If

                                Set tAss = t.Assignments

                                For Each tAssign In tAss

                                    If (tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0) And _
                                    (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R") Then

                                        If TimeScaleExport = True Then

                                            ExportTimeScaleResources ID, t, tAssign, 2, "BCWS"

                                        Else

                                            Select Case tAssign.ResourceType

                                                Case pjResourceTypeWork

                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                                Case pjResourceTypeCost

                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                                Case pjResourceTypeMaterial

                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                            End Select

                                        End If
                                    Else
                                    
                                        If (tAssign.Work <> 0 Or tAssign.Cost <> 0) And _
                                        (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R") Then

                                            If TimeScaleExport = True Then
    
                                                ExportTimeScaleResources ID, t, tAssign, 2, "ETC"
    
                                            Else
    
                                                Select Case tAssign.ResourceType
    
                                                    Case pjResourceTypeWork
    
                                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
    
                                                    Case pjResourceTypeCost
    
                                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
    
                                                    Case pjResourceTypeMaterial
    
                                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
    
                                                End Select
    
                                            End If
                                            
                                        End If
                                        
                                    End If

                                Next tAssign

                            End If

                        End If

                    End If
Next_SProj_Task:

                Next t

                FileClose pjDoNotSave

            Next subProj

        Else

            For Each t In curproj.Tasks

                If Not t Is Nothing Then

                    If t.GetField(188744959) = "Yes" And t.Summary = False And t.ExternalTask = False Then

                        If ((t.BaselineWork > 0 Or t.BaselineCost > 0) And _
                        (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "d" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "D")) _
                        Or _
                        ((t.Work > 0 Or t.Cost > 0) And _
                        (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r")) Then

                            CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                            If CAID3_Used = True Then
                                CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                            End If
                            WP = t.GetField(FieldNameToFieldConstant(fWP))
                            EVT = t.GetField(FieldNameToFieldConstant(fEVT))

                            If CAID2_Used = True Then
                                CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                            End If
                            CAM = CleanCamName(t.GetField(FieldNameToFieldConstant(fCAM)))
                            If CAID3_Used = True And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = True Then
                                ID = CAID1 & "/" & CAID2 & "/" & WP
                            End If
                            If CAID3_Used = False And CAID2_Used = False Then
                                ID = CAID1 & "/" & WP
                            End If

                            If Milestones_Used = True Then
                                UID = t.GetField(FieldNameToFieldConstant(fMilestone))
                                MSWeight = CleanNumber(t.GetField(FieldNameToFieldConstant(fMilestoneWeight)))
                            End If

                            If EVT = "B" And Milestones_Used = False Then
                                ErrMsg = "Error: Found EVT = B, missing Milestone Field Maps"
                                Err.Raise 1
                            End If

                            If BCRxport = True Then
                                If IsInArray(WP, BCR_WP) = False Then
                                    GoTo next_task
                                End If
                            End If

                            'store ACT info
                            'WP Data
                            If X = 1 Then

                                'create new WP line in ACTarrray
                                ReDim ACTarray(1 To X)
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                    ACTarray(X).BFinish = t.BaselineFinish
                                    ACTarray(X).BStart = t.BaselineStart
                                Else
                                    ACTarray(X).BFinish = t.Finish
                                    ACTarray(X).BStart = t.Start
                                End If
                                If CAID3_Used = True Then
                                    ACTarray(X).CAID3 = CAID3
                                End If
                                ACTarray(X).ID = ID
                                ACTarray(X).CAM = CAM
                                ACTarray(X).CAID1 = CAID1
                                ACTarray(X).EVT = EVT
                                ACTarray(X).FFinish = t.Finish
                                ACTarray(X).FStart = t.Start
                                If CAID2_Used = True Then
                                    ACTarray(X).CAID2 = CAID2
                                End If
                                ACTarray(X).WP = WP

                                X = X + 1
                                ActFound = True

                                GoTo WP_Match_B

                            End If

                            For i = 1 To UBound(ACTarray)
                                If ACTarray(i).ID = ID Then
                                    'Found an existing matching WP line
                                    If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        If ACTarray(i).BStart > t.BaselineStart Then
                                            ACTarray(i).BStart = t.BaselineStart
                                        End If
                                        If ACTarray(i).BFinish < t.BaselineFinish Then
                                            ACTarray(i).BFinish = t.BaselineFinish
                                        End If
                                    Else
                                        If ACTarray(i).BStart > t.Start Then
                                            ACTarray(i).BStart = t.Start
                                        End If
                                        If ACTarray(i).BFinish < t.Finish Then
                                            ACTarray(i).BFinish = t.Finish
                                        End If

                                    End If
                                    If ACTarray(i).FStart > t.Start Then
                                        ACTarray(i).FStart = t.Start
                                    End If
                                    If ACTarray(i).FFinish < t.Finish Then
                                        ACTarray(i).FFinish = t.Finish
                                    End If
                                    GoTo WP_Match_B
                                End If
                            Next i

                            'No match found, create new WP line in ACTarrray
                            ReDim Preserve ACTarray(1 To X)
                            If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                ACTarray(X).BFinish = t.BaselineFinish
                                ACTarray(X).BStart = t.BaselineStart
                            Else
                                ACTarray(X).BFinish = t.Finish
                                ACTarray(X).BStart = t.Start
                            End If
                            If CAID3_Used = True Then
                                ACTarray(X).CAID3 = CAID3
                            End If
                            ACTarray(X).CAM = CAM
                            ACTarray(X).CAID1 = CAID1
                            ACTarray(X).EVT = EVT
                            ACTarray(X).ID = ID
                            ACTarray(X).FFinish = t.Finish
                            ACTarray(X).FStart = t.Start
                            If CAID2_Used = True Then
                                ACTarray(X).CAID2 = CAID2
                            End If
                            ACTarray(X).WP = WP

                            X = X + 1
                            ActFound = True

                            'Milestone Data
WP_Match_B:

                            If EVT = "B" Or EVT = "B Milestone" Then
                                If t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" Then
                                        
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.BaselineStart, "M/D/YYYY") & "," & Format(t.BaselineFinish, "M/D/YYYY") & ","
                                    End If
                                    
                                Else
                                    
                                    If CAID3_Used = True And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID3 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = True Then
                                        Print #1, CAID1 & "," & CAID2 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                    End If
                                    If CAID3_Used = False And CAID2_Used = False Then
                                        Print #1, CAID1 & "," & CAM & "," & WP & "," & "," & UID & "," & MSWeight & "," & Replace(t.Name, ",", "") & "," & Format(t.Start, "M/D/YYYY") & "," & Format(t.Finish, "M/D/YYYY") & ","
                                    End If
                                    
                                End If
                            End If

                            Set tAss = t.Assignments

                            For Each tAssign In tAss

                                If (tAssign.BaselineWork <> 0 Or tAssign.BaselineCost <> 0) And _
                                (t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "r" And t.GetField(FieldNameToFieldConstant(fWhatIf)) <> "R") Then

                                    If TimeScaleExport = True Then

                                        ExportTimeScaleResources ID, t, tAssign, 2, "BCWS"

                                    Else

                                        Select Case tAssign.ResourceType

                                            Case pjResourceTypeWork

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                            Case pjResourceTypeCost

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineCost & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                            Case pjResourceTypeMaterial

                                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.BaselineWork & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                        End Select

                                    End If
                                    
                                Else
                                    
                                    If (tAssign.Work <> 0 Or tAssign.Cost <> 0) And _
                                    (t.GetField(FieldNameToFieldConstant(fWhatIf)) = "r" Or t.GetField(FieldNameToFieldConstant(fWhatIf)) = "R") Then

                                        If TimeScaleExport = True Then
    
                                            ExportTimeScaleResources ID, t, tAssign, 2, "ETC"
    
                                        Else
    
                                            Select Case tAssign.ResourceType
    
                                                Case pjResourceTypeWork
    
                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
    
                                                Case pjResourceTypeCost
    
                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Cost & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
    
                                                Case pjResourceTypeMaterial
    
                                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tAssign.Work & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")
    
                                            End Select
    
                                        End If
                                    
                                    End If

                                End If

                            Next tAssign

                        End If

                    End If

                End If
next_task:

            Next t

        End If

        If ActFound = True Then
            For i = 1 To UBound(ACTarray)

                If DescExport = True Then
                    ACTarray(i).Desc = WP_Desc(ACTarray(i).ID)
                End If

                If CAID3_Used = True And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID3 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAID2 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    Print #1, ACTarray(i).CAID1 & "," & ACTarray(i).CAM & "," & ACTarray(i).WP & "," & ACTarray(i).ID & "," & "," & "," & ACTarray(i).Desc & "," & Format(ACTarray(i).BStart, "M/D/YYYY") & "," & Format(ACTarray(i).BFinish, "M/D/YYYY") & "," & ACTarray(i).EVT
                End If

            Next i
        End If

        Close #1
        Close #2

    End If
        
End Sub

Private Function SetDirectory(ByVal ProjName As String) As String
    Dim newDir As String
    Dim pathDesktop As String

    pathDesktop = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    newDir = pathDesktop & "\" & RemoveIllegalCharacters(ProjName) & "_" & Format(Now, "YYYYMMDD HHMMSS")

    MkDir newDir
    SetDirectory = newDir
    Exit Function

End Function

Private Sub DeleteDirectory(ByVal DirName As String)

    RmDir DirName & "\"
    Exit Sub

End Sub

Private Function WP_Desc(ByVal WP_ID As String) As String

    Dim tempDesc As String
    Dim X As Integer

    tempDesc = ""

    If WPDescCount = 0 Then
        WP_Desc = tempDesc
        Exit Function
    End If

    On Error GoTo NoWPMatchFound

    For X = 1 To UBound(WPDescArray)
        If WPDescArray(X).WP_ID = WP_ID Then
            tempDesc = WPDescArray(X).Desc
            WP_Desc = tempDesc
            Exit Function
        End If
    Next X

NoWPMatchFound:

    WP_Desc = ""

End Function

Private Sub Get_WP_Descriptions(ByVal curproj As Project)

    Dim CAID1 As String
    Dim CAID2 As String
    Dim CAID3 As String
    Dim WP As String
    Dim ID As String
    Dim Desc As String
    Dim i As Integer
    Dim X As Integer
    '<issue47>
    Dim subProjs As Subprojects
    Dim subProj As SubProject
    Dim curSProj As Project
    Dim t As Task '</issue47>

    WPDescCount = 0

    i = 0

    If curproj.Subprojects.Count > 0 Then

        Set subProjs = curproj.Subprojects

        For Each subProj In subProjs

            FileOpen Name:=subProj.Path, ReadOnly:=True

            Set curSProj = ActiveProject

            For Each t In curSProj.Tasks '<issue47>

                If Not t Is Nothing Then

                    WP = t.GetField(FieldNameToFieldConstant(fWP))

                    If WP = "" Then GoTo Next_SubProj_WPtask

                    CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                    If CAID3_Used = True Then
                        CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                    End If
                    If CAID2_Used = True Then
                        CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                    End If
                    If CAID3_Used = True And CAID2_Used = True Then
                        ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                    End If
                    If CAID3_Used = False And CAID2_Used = True Then
                        ID = CAID1 & "/" & CAID2 & "/" & WP
                    End If
                    If CAID3_Used = False And CAID2_Used = False Then
                        ID = CAID1 & "/" & WP
                    End If
                    Desc = Replace(t.Name, ",", "")

                    If i = 0 Then
                        i = 1
                    End If

                    If i = 1 Then

                        ReDim WPDescArray(1 To i)
                        WPDescArray(i).WP_ID = ID
                        WPDescArray(i).Desc = Desc
                        WPDescCount = i
                        i = i + 1

                    Else

                        For X = 1 To UBound(WPDescArray)

                            If WPDescArray(X).WP_ID = ID Then
                                GoTo Next_SubProj_WPtask
                            End If
                        Next X

                        ReDim Preserve WPDescArray(1 To i)
                        WPDescArray(i).WP_ID = ID
                        WPDescArray(i).Desc = Desc
                        WPDescCount = i
                        i = i + 1

                    End If

                End If
Next_SubProj_WPtask:

            Next t

        Next subProj

    Else

        For Each t In curproj.Tasks

            If Not t Is Nothing Then
                WP = t.GetField(FieldNameToFieldConstant(fWP))

                If WP = "" Then GoTo Next_WPtask

                CAID1 = t.GetField(FieldNameToFieldConstant(fCAID1))
                If CAID3_Used = True Then
                    CAID3 = t.GetField(FieldNameToFieldConstant(fCAID3))
                End If
                If CAID2_Used = True Then
                    CAID2 = t.GetField(FieldNameToFieldConstant(fCAID2))
                End If
                If CAID3_Used = True And CAID2_Used = True Then
                    ID = CAID1 & "/" & CAID2 & "/" & CAID3 & "/" & WP
                End If
                If CAID3_Used = False And CAID2_Used = True Then
                    ID = CAID1 & "/" & CAID2 & "/" & WP
                End If
                If CAID3_Used = False And CAID2_Used = False Then
                    ID = CAID1 & "/" & WP
                End If
                Desc = Replace(t.Name, ",", "")

                If i = 0 Then
                    i = 1
                End If

                If i = 1 Then

                    ReDim WPDescArray(1 To i)
                    WPDescArray(i).WP_ID = ID
                    WPDescArray(i).Desc = Desc
                    WPDescCount = i
                    i = i + 1

                Else

                    For X = 1 To UBound(WPDescArray)

                        If WPDescArray(X).WP_ID = ID Then
                            GoTo Next_WPtask
                        End If
                    Next X

                    ReDim Preserve WPDescArray(1 To i)
                    WPDescArray(i).WP_ID = ID
                    WPDescArray(i).Desc = Desc
                    WPDescCount = i
                    i = i + 1

                End If

            End If

Next_WPtask:

        Next t

    End If

End Sub

Private Function IsInArray(ByVal stringToBeFound As String, ByVal arr As Variant) As Boolean
    On Error GoTo NullArray
    Dim testInt As Integer
    testInt = UBound(Filter(arr, stringToBeFound))
    IsInArray = (testInt > -1)
    Debug.Print "Searching for WP ID: " & stringToBeFound
    Exit Function
NullArray:
    IsInArray = False
End Function

Private Function CleanCamName(ByVal CAM As String) As String

    Dim tempCAM As String

    tempCAM = CAM

    If InStr(tempCAM, ".") > 0 Then
        tempCAM = Right(tempCAM, Len(tempCAM) - InStr(tempCAM, "."))
    End If

    If InStr(tempCAM, ",") > 0 Then
        tempCAM = Replace(tempCAM, ",", " ")
    End If

    CleanCamName = tempCAM

End Function

Private Function Find_BCRs(ByVal curproj As Project, ByVal fWP As String, ByVal fBCR As String, ByVal BCRnum As String) As Integer

    Dim t As Task
    Dim i As Integer
    Dim X As Integer
    Dim tempBCRstr As String
    Dim subProjs As Subprojects
    Dim subProj As SubProject
    Dim curSProj As Project

    i = 0

    If curproj.Subprojects.Count > 0 Then

        Set subProjs = curproj.Subprojects

        For Each subProj In subProjs

            FileOpen Name:=subProj.Path, ReadOnly:=True

            Set curSProj = ActiveProject

            For Each t In curSProj.Tasks

                If Not t Is Nothing Then

                    tempBCRstr = t.GetField(FieldNameToFieldConstant(fBCR))

                    If InStr(tempBCRstr, BCRnum) > 0 Then

                        If i = 0 Then
                            i = 1
                        End If

                        If i = 1 Then

                            ReDim BCR_WP(1 To i)
                            BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                            Debug.Print "Storing BCR ID: " & BCR_WP(i)
                            i = i + 1

                        Else

                            For X = 1 To UBound(BCR_WP)
                                If BCR_WP(X) = t.GetField(FieldNameToFieldConstant(fWP)) Then
                                    GoTo Next_SubProj_WPtask
                                End If
                            Next X

                            ReDim Preserve BCR_WP(1 To i)
                            BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                            Debug.Print "Storing BCR ID: " & BCR_WP(i)
                            i = i + 1

                        End If

                    End If
                End If
Next_SubProj_WPtask:

            Next t

        Next subProj

    Else

        For Each t In curproj.Tasks

            If Not t Is Nothing Then
                tempBCRstr = t.GetField(FieldNameToFieldConstant(fBCR))

                If InStr(tempBCRstr, BCRnum) > 0 Then

                    If i = 0 Then
                        i = 1
                    End If

                    If i = 1 Then

                        ReDim BCR_WP(1 To i)
                        BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                        Debug.Print "Storing BCR ID: " & BCR_WP(i)
                        i = i + 1

                    Else

                        For X = 1 To UBound(BCR_WP)
                            If BCR_WP(X) = t.GetField(FieldNameToFieldConstant(fWP)) Then
                                GoTo Next_WPtask
                            End If
                        Next X

                        ReDim Preserve BCR_WP(1 To i)
                        BCR_WP(i) = t.GetField(FieldNameToFieldConstant(fWP))
                        Debug.Print "Storing BCR ID: " & BCR_WP(i)
                        i = i + 1

                    End If

                End If

            End If

Next_WPtask:

        Next t

    End If

    Find_BCRs = i

End Function

Private Function RemoveIllegalCharacters(ByVal strText As String) As String

    Const cstrIllegals As String = "\,/,:,*,?,"",<,>,|"

    Dim lngCounter As Long
    Dim astrChars() As String

    astrChars() = Split(cstrIllegals, ",")

    For lngCounter = LBound(astrChars()) To UBound(astrChars())
        strText = Replace(strText, astrChars(lngCounter), vbNullString)
    Next lngCounter

    RemoveIllegalCharacters = strText

End Function

Private Sub ReadCustomFields(ByVal curproj As Project)

    Dim i As Integer
    Dim fID As Long

    'Read local Custom Text Fields
    For i = 1 To 30

        If Len(curproj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))) > 0 Then
            ReDim Preserve CustTextFields(1 To i)
            CustTextFields(i) = curproj.Application.CustomFieldGetName(FieldNameToFieldConstant("Text" & i))
        Else
            ReDim Preserve CustTextFields(1 To i)
            CustTextFields(i) = "Text" & i
        End If

    Next i
    
    'Read local Custom Number Fields
    For i = 1 To 20

        If Len(curproj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))) > 0 Then
            ReDim Preserve CustNumFields(1 To i)
            CustNumFields(i) = curproj.Application.CustomFieldGetName(FieldNameToFieldConstant("Number" & i))
        Else
            ReDim Preserve CustNumFields(1 To i)
            CustNumFields(i) = "Number" & i
        End If

    Next i

    'Read local Custom Outline Code Fields
    For i = 1 To 10

        If Len(curproj.Application.CustomFieldGetName(FieldNameToFieldConstant("OutlineCode" & i))) > 0 Then
            ReDim Preserve CustOLCodeFields(1 To i)
            CustOLCodeFields(i) = curproj.Application.CustomFieldGetName(FieldNameToFieldConstant("OutlineCode" & i))
        Else
            ReDim Preserve CustOLCodeFields(1 To i)
            CustOLCodeFields(i) = "OutlineCode" & i
        End If

    Next i
    
    'Read Enterprise Custom Fields
    i = 1

    For fID = 188776000 To 188778000

        On Error GoTo fID_Error

        If Application.CustomFieldGetName(fID) <> "" Then
            ReDim Preserve EntFields(1 To i)
            EntFields(i) = Application.CustomFieldGetName(fID)
            i = i + 1
        End If

next_fID:

    Next fID
    
    Exit Sub

fID_Error:

    Resume next_fID

End Sub

Private Function CleanNumber(ByVal NumStr As String) As String

    Dim i As Integer
    Dim newNumStr As String

    For i = 1 To Len(NumStr)

        If Mid(NumStr, i, 1) = "." Or IsNumeric(Mid(NumStr, i, 1)) Then

            newNumStr = newNumStr & Mid(NumStr, i, 1)

        End If

    Next

    CleanNumber = newNumStr

End Function
Private Function PercentfromString(ByVal inputStr As String) As Double

    Dim tempDbl As String

    'Test for % String
    If InStr(inputStr, "%") > 0 Then

        tempDbl = Left(inputStr, Len(inputStr) - 1)

    Else

        tempDbl = inputStr

    End If

    PercentfromString = CDbl(tempDbl)

End Function
Private Sub ExportTimeScaleResources(ByVal ID As String, ByVal t As Task, ByVal tAssign As Assignment, ResFile As Integer, exportType As String)

    Dim tsv As TimeScaleValue
    Dim tsvs As TimeScaleValues
    Dim tsvsa As TimeScaleValues
    Dim tsva As TimeScaleValue
    Dim tempWork As Double

    Select Case exportType

        Case "ETC"

            Select Case tAssign.ResourceType

                Case pjResourceTypeWork

                    If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                        Set tsvs = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        Set tsvsa = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleWeeks)

                        For Each tsv In tsvs

                            Set tsva = tsvsa(tsv.Index)

                            tempWork = 0

                            If tsva <> "" Then
                                tempWork = CDbl(tsv.Value) - CLng(tsva.Value)
                            ElseIf tsv.Value <> "" Then
                                tempWork = CDbl(tsv.Value)
                            End If

                            If tempWork <> 0 Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork / 60 & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                        Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        For Each tsv In tsvs

                            If tsv.Value <> "" Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    End If

            Case pjResourceTypeCost

                If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                    Set tsvs = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleWeeks)
                    Set tsvsa = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledActualCost, pjTimescaleWeeks)

                    For Each tsv In tsvs

                        Set tsva = tsvsa(tsv.Index)

                        tempWork = 0

                        If tsva <> "" Then
                            tempWork = CDbl(tsv.Value) - CLng(tsva.Value)
                        ElseIf tsv.Value <> "" Then
                            tempWork = CDbl(tsv.Value)
                        End If

                        If tempWork <> 0 Then

                            If tsvs.Count = 1 Then

                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                            Else

                                Select Case tsv.Index

                                    Case 1

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    Case tsvs.Count

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                    Case Else

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                End Select

                            End If

                        End If

                    Next tsv

                    Exit Sub

                ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                    Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledCost, pjTimescaleWeeks)
                    For Each tsv In tsvs

                        If tsv.Value <> "" Then

                            If tsvs.Count = 1 Then

                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                            Else

                                Select Case tsv.Index

                                    Case 1

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    Case tsvs.Count

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                    Case Else

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                End Select

                            End If

                        End If

                    Next tsv

                    Exit Sub

                End If

            Case pjResourceTypeMaterial

                If t.Resume <> "NA" And t.ActualFinish = "NA" And tAssign.PercentWorkComplete <> 100 Then

                        Set tsvs = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        Set tsvsa = tAssign.TimeScaleData(t.Resume, tAssign.Finish, pjAssignmentTimescaledActualWork, pjTimescaleWeeks)

                        For Each tsv In tsvs

                            Set tsva = tsvsa(tsv.Index)

                            tempWork = 0

                            If tsva <> "" Then
                                tempWork = CDbl(tsv.Value) - CLng(tsva.Value)
                            ElseIf tsv.Value <> "" Then
                                tempWork = CDbl(tsv.Value)
                            End If

                            If tempWork <> 0 Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(t.Resume, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tempWork & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    ElseIf t.Resume = "NA" And tAssign.PercentWorkComplete <> 100 Then

                        Set tsvs = tAssign.TimeScaleData(tAssign.Start, tAssign.Finish, pjAssignmentTimescaledWork, pjTimescaleWeeks)
                        For Each tsv In tsvs

                            If tsv.Value <> "" Then

                                If tsvs.Count = 1 Then

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                Else

                                    Select Case tsv.Index

                                        Case 1

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.Start, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                        Case tsvs.Count

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.Finish, "M/D/YYYY")

                                        Case Else

                                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    End Select

                                End If

                            End If

                        Next tsv

                        Exit Sub

                    End If

            End Select

        Case "BCWS"

            Select Case tAssign.ResourceType

                Case pjResourceTypeWork

                    Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                    For Each tsv In tsvs

                        If tsv.Value <> "" Then

                            If tsvs.Count = 1 Then

                                Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                            Else

                                Select Case tsv.Index

                                    Case 1

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                    Case tsvs.Count

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                    Case Else

                                        Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value / 60 & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                End Select

                            End If

                        End If

                    Next tsv

                    Exit Sub

            Case pjResourceTypeCost

                Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineCost, pjTimescaleWeeks)
                For Each tsv In tsvs

                    If tsv.Value <> "" Then

                        If tsvs.Count = 1 Then

                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                        Else

                            Select Case tsv.Index

                                Case 1

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                Case tsvs.Count

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                Case Else

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                            End Select

                        End If

                    End If

                Next tsv

                Exit Sub

            Case pjResourceTypeMaterial

                Set tsvs = tAssign.TimeScaleData(tAssign.BaselineStart, tAssign.BaselineFinish, pjAssignmentTimescaledBaselineWork, pjTimescaleWeeks)
                For Each tsv In tsvs

                    If tsv.Value <> "" Then

                        If tsvs.Count = 1 Then

                            Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                        Else

                            Select Case tsv.Index

                                Case 1

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tAssign.BaselineStart, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                                Case tsvs.Count

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tAssign.BaselineFinish, "M/D/YYYY")

                                Case Else

                                    Print #2, ID & "," & tAssign.Resource.GetField(FieldNameToFieldConstant(fResID, pjResource)) & "," & tsv.Value & "," & Format(tsv.StartDate, "M/D/YYYY") & "," & Format(tsv.EndDate - 1, "M/D/YYYY")

                            End Select

                        End If

                    End If

                Next tsv

                Exit Sub

            End Select

        Case Else

            Exit Sub

    End Select

End Sub

Private Function get_Assignment_Pcnt(ByVal tAssignment As Assignment) As String

    Dim pcntField As String

    If AssignmentPCNT_Used = False Then 'v3.3.2
    
        get_Assignment_Pcnt = tAssignment.Task.GetField(FieldNameToFieldConstant(fPCNT))
        Exit Function
    
    Else
    
        pcntField = fAssignPcnt
        
    End If

    Select Case FieldNameToFieldConstant(pcntField)
    
        Case FieldNameToFieldConstant("Number1")
        
            get_Assignment_Pcnt = tAssignment.Number1
            Exit Function
        
        Case FieldNameToFieldConstant("Number2")
        
            get_Assignment_Pcnt = tAssignment.Number2
            Exit Function
        
        Case FieldNameToFieldConstant("Number3")
        
            get_Assignment_Pcnt = tAssignment.Number3
            Exit Function
        
        Case FieldNameToFieldConstant("Number4")
        
            get_Assignment_Pcnt = tAssignment.Number4
            Exit Function
        
        Case FieldNameToFieldConstant("Number5")
        
            get_Assignment_Pcnt = tAssignment.Number5
            Exit Function
        
        Case FieldNameToFieldConstant("Number6")
        
            get_Assignment_Pcnt = tAssignment.Number6
            Exit Function
        
        Case FieldNameToFieldConstant("Number7")
        
            get_Assignment_Pcnt = tAssignment.Number7
            Exit Function
        
        Case FieldNameToFieldConstant("Number8")
        
            get_Assignment_Pcnt = tAssignment.Number8
            Exit Function
        
        Case FieldNameToFieldConstant("Number9")
        
            get_Assignment_Pcnt = tAssignment.Number9
            Exit Function
        
        Case FieldNameToFieldConstant("Number10")
        
            get_Assignment_Pcnt = tAssignment.Number10
            Exit Function
            
        Case FieldNameToFieldConstant("Number11")
        
            get_Assignment_Pcnt = tAssignment.Number11
            Exit Function
        
        Case FieldNameToFieldConstant("Number12")
        
            get_Assignment_Pcnt = tAssignment.Number12
            Exit Function
        
        Case FieldNameToFieldConstant("Number13")
        
            get_Assignment_Pcnt = tAssignment.Number13
            Exit Function
        
        Case FieldNameToFieldConstant("Number14")
        
            get_Assignment_Pcnt = tAssignment.Number14
            Exit Function
        
        Case FieldNameToFieldConstant("Number15")
        
            get_Assignment_Pcnt = tAssignment.Number15
            Exit Function
        
        Case FieldNameToFieldConstant("Number16")
        
            get_Assignment_Pcnt = tAssignment.Number16
            Exit Function
        
        Case FieldNameToFieldConstant("Number17")
        
            get_Assignment_Pcnt = tAssignment.Number17
            Exit Function
        
        Case FieldNameToFieldConstant("Number18")
        
            get_Assignment_Pcnt = tAssignment.Number18
            Exit Function
        
        Case FieldNameToFieldConstant("Number19")
        
            get_Assignment_Pcnt = tAssignment.Number19
            Exit Function
        
        Case FieldNameToFieldConstant("Number20")
        
            get_Assignment_Pcnt = tAssignment.Number20
            Exit Function
        
        Case Else
        
            get_Assignment_Pcnt = 0
            Exit Function
            
    End Select
    
    get_Assignment_Pcnt = 0

End Function

Public Sub cptQuickSort(vArray As Variant, inLow As Long, inHi As Long)
  'public domain: https://stackoverflow.com/questions/152319/vba-array-sort-function
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then cptQuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then cptQuickSort vArray, tmpLow, inHi
End Sub
