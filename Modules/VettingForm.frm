VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VettingForm 
   Caption         =   "MeasurLink Routine Vetting"
   ClientHeight    =   8205.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7260
   OleObjectBlob   =   "VettingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************************************
'
'   Vetting Form
'       1. Everything happens in _Initialize(),
'           We set the Routine, ObsReq and ObsFound in the respesctive routines
'           As we go through the routines for the part number, we set the name and make the text gray
'           When we find a routine of the same name that we ran for that job number, we fill it black to indicate a match
'           A routine being gray doesn't auto indicate a failure, it might depend on the job type
'*************************************************************************************************




Dim cellLeadAlertReq As Boolean
Dim qcManagerAlertReq As Boolean
Dim failedRoutines() As Variant




'****************************************************************************************
'               UserForm Callbacks
'****************************************************************************************

Private Sub UserForm_Initialize()
    
    Call SetActivePrinter


    'Set the required routines for the part
    'Also set the required observations for the routine
    For i = 0 To UBound(RibbonCommands.partRoutineList, 2)
        With Me.RoutineFrame.Controls(i)
            .Caption = RibbonCommands.partRoutineList(0, i)
            .ForeColor = RGB(128, 128, 128)
            .Visible = True
        End With
        With Me.ObsFound.Controls(i)
            .Caption = ""
            .Visible = False
        End With
        With Me.ObsReq.Controls(i)
        
            On Error GoTo RoutineNameErr
        
            Dim routineType As String
            routineType = Split(RibbonCommands.partRoutineList(0, i), RibbonCommands.partNum & "_" & RibbonCommands.rev & "_")(1)
            
            
            On Error GoTo RoutineSwitchErr
            
            'Given routine of a name like "DRW-00717-01_RAG_IP_SYLVAC", we're trying to grab the "IP_SYLVAC"
            Select Case (routineType)
                Case "FA_FIRST"
                    If (RibbonCommands.chkFull_Pressed) Then
                        .Caption = "2"
                        .Visible = True
                    Else
                        .Caption = "0"
                        .Visible = False
                    End If
                Case "FA_SYLVAC", "FA_CMM", "FA_RAMPROG"
                    If (RibbonCommands.chkFull_Pressed) Then
                        .Caption = "1"
                        .Visible = True
                    Else
                        .Caption = "0"
                        .Visible = False
                    End If
                Case "FA_MINI"
                    If (RibbonCommands.chkMini_Pressed) Then
                        .Caption = "2"
                        .Visible = True
                    Else
                        .Caption = "0"
                        .Visible = False
                    End If
                Case "FA_VIS"
                    If (RibbonCommands.chkNone_Pressed) Then
                        .Caption = "2"
                        .Visible = True
                    Else
                        .Caption = "0"
                        .Visible = False
                    End If
                Case "IP_1XSHIFT"
                    .Caption = DatabaseModule.Get1XSHIFTInsps(JobID:=RibbonCommands.jobNumUcase)
                    .Visible = True
                Case "IP_EDM"
                    .Caption = RibbonCommands.prodQty
                    .Visible = True
                Case "FI_VIS", "IP_LAST"
                    .Caption = "1"
                    .Visible = True
                Case "FI_DIM"
                    If DatabaseModule.IsAllAttribrute(routine:=RibbonCommands.partRoutineList(0, i)) Then
                        .Caption = "1"
                    Else
                        .Caption = ExcelHelpers.GetAQL(customer:=RibbonCommands.customer, drawNum:=RibbonCommands.drawNum, _
                                            prodQty:=RibbonCommands.prodQty)
                    End If
                    .Visible = True
                Case Else
                    'Anything not covered above should be AQL quantity
                    .Caption = ExcelHelpers.GetAQL(customer:=RibbonCommands.customer, drawNum:=RibbonCommands.drawNum, _
                                            prodQty:=RibbonCommands.prodQty)
                    .Visible = True
            End Select
        End With
        
    Next i
    
    On Error GoTo 0
    
    'Hide and reset the excess controls
    For i = i To Me.RoutineFrame.Controls.Count - 1
        Me.RoutineFrame.Controls(i).Visible = False
        Me.RoutineFrame.Controls(i).Caption = ""
        Me.ObsReq.Controls(i).Visible = False
        Me.ObsReq.Controls(i).Caption = ""
        Me.ObsFound.Controls(i).Visible = False
        Me.ObsFound.Controls(i).Caption = ""
        Me.ResultFrame.Controls(i).Visible = False
    Next i
    
    'Fill the required routines with the routines that we found in the run
    'Also fill in the observations that we have collected
    For i = 0 To UBound(RibbonCommands.runRoutineList, 2)
        For j = 0 To Me.RoutineFrame.Count - 1
            'If the control name matches the found routine name, then we should make the text black and fill in the observations found
            If (RibbonCommands.runRoutineList(0, i) = Me.RoutineFrame.Controls(j).Caption) Then
                Me.RoutineFrame.Controls(j).ForeColor = RGB(0, 0, 0)
                Me.ObsFound.Controls(j).Caption = runRoutineList(2, i)
                Me.ObsFound.Controls(j).Visible = True
                GoTo NextControl
            End If
        Next j
        'If we couldnt find a match between required routines and the run routine
        GoTo UniqueRoutineErr
NextControl:
    Next i
    
    
    Call VetInspections
    
    If (cellLeadAlertReq Or qcManagerAlertReq) Then
        Me.PrintButton.Enabled = False
        Me.EmailButton.Enabled = True
        Me.EmailButton.SetFocus
    Else
        Me.PrintButton.Enabled = True
        Me.EmailButton.Enabled = False
        Me.PrintButton.SetFocus
    End If
    
    Exit Sub
    
RoutineNameErr:
   Result = MsgBox("Could Not Parse the Routine Name " & RibbonCommands.partRoutineList(0, i) & _
                vbCrLf & "Alert a QE" & _
                vbCrLf & "Routines Must Follow the standard naming convention of [Part]_[Rev]_[OPtype]_[Routine SubType]", vbExclamation)
    Exit Sub

UniqueRoutineErr:
   Result = MsgBox("Application found this routine: " & RibbonCommands.runRoutineList(0, i) & _
                vbCrLf & "Which doesn't match any of our required routines" & _
                vbCrLf & "If a routine name changed, it could cause misalignment here", vbInformation)
   GoTo NextControl

RoutineSwitchErr:
       Result = MsgBox("Error when determining observations needed for : " & RibbonCommands.runRoutineList(0, i) & vbCrLf & _
                Err.description, vbCritical)

End Sub


Private Sub EmailButton_Click()
    Dim cellLeadEmail As String
    cellLeadEmail = DatabaseModule.GetCellLeadEmail(cell:=RibbonCommands.cell)
    
    Call ExcelHelpers.CreateEmail(qcManager:=qcManagerAlertReq, cellLead:=cellLeadAlertReq, cellLeadEmail:=cellLeadEmail, _
                                    jobNum:=RibbonCommands.jobNumUcase, machine:=RibbonCommands.machine, failInfo:=failedRoutines)

End Sub

Private Sub PrintButton_Click()
    Unload Me
    Call RibbonCommands.IterPrintRoutines
End Sub


Private Sub UserForm_Activate()
'    MsgBox (Me.Controls("RoutineFrame").Routine1.Caption)
End Sub


Private Sub ChangePrinterButton_Click()
    If (Application.Dialogs(xlDialogPrinterSetup).Show) Then
        Call SetActivePrinter
    End If
End Sub




'****************************************************************************************
'               Extra Functions
'****************************************************************************************
Private Sub VetInspections()
    
    For i = 0 To Me.RoutineFrame.Controls.Count - 1
        'If there is a Req Qty and a Found Qty
        If (Me.ObsReq.Controls(i).Visible = True And Me.ObsFound.Controls(i).Visible = True) Then
            'If our Found Qty meets our criteria
            If (CInt(Me.ObsFound.Controls(i).Caption) >= CInt(Me.ObsReq.Controls(i).Caption)) Then
                GoTo NextIter
            'If it doesn't meet our criteria
            Else
                Call setFailure(location:=i, routine:=Me.RoutineFrame.Controls(i).Caption)
                GoTo NextIter
            End If
        
        'If we have Req Qty but didn't find results for inspected Qty
        ElseIf Me.ObsReq.Controls(i).Visible = True And Me.ObsFound.Controls(i).Visible = False Then
            Call setFailure(location:=i, routine:=Me.RoutineFrame.Controls(i).Caption)
            GoTo NextIter
        'If we have a routine but no Req Qty because the setup type doesn't require it, not considered a failure
        ElseIf (Me.RoutineFrame.Controls(i).Visible = True And Me.ObsReq.Controls(i).Visible = False) Then
            Call hideResult(location:=i)
            GoTo NextIter
        'If its a hidden control
        ElseIf Me.RoutineFrame.Controls(i).Visible = False Then
        Else
            GoTo RoutineReadErr
        End If
        
NextIter:
    Next i
    
    Exit Sub
    
RoutineReadErr:

   Result = MsgBox("Could not correctly compare the quantities of " & Me.RoutineFrame.Controls(i).Caption & _
                    vbCrLf & "Please alert a QE to this.", vbExclamation)

End Sub


Private Sub setFailure(location As Variant, routine As String)
    'Set failure picture
    With Me.ResultFrame.Controls(location)
        .Picture = LoadPicture(DataSources.FAIL_IMG_PATH)
        .Height = 15
        .Width = 15
        .Top = .Top + 4
    
    End With
    
    
    'If there was a problem with and 'FI' Routine like 'FI_DIM' then we need to alert the QC manager, otherwise we need to alert the cell lead
    Dim routineSuffix As String
    routineSuffix = Split(routine, RibbonCommands.partNum & "_" & RibbonCommands.rev & "_")(1)
    If (InStr(routineSuffix, "FI") > 0) Then
        qcManagerAlertReq = True
    Else
        cellLeadAlertReq = True
    End If
    
    'Store the failed results for later passing to the email
    Dim index As Integer
        'Checks if the array has ever been initialized
    If (Not failedRoutines) = -1 Then
        index = 0
        ReDim Preserve failedRoutines(2, 0)
    Else
        index = UBound(failedRoutines, 2) + 1
        ReDim Preserve failedRoutines(2, index)
    End If
    
    failedRoutines(0, index) = Me.RoutineFrame(location).Caption
    failedRoutines(1, index) = Me.ObsReq(location).Caption
    failedRoutines(2, index) = Me.ObsFound(location).Caption
    
End Sub

Private Sub hideResult(location As Variant)
    Me.ResultFrame.Controls(location).Visible = False
End Sub

Private Sub SetActivePrinter()
    Me.ActivePrinter.Caption = Split(Application.ActivePrinter, " ")(0)

End Sub





