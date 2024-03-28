VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VettingForm 
   Caption         =   "MeasurLink Routine Vetting"
   ClientHeight    =   8260
   ClientLeft      =   -3800
   ClientTop       =   -15330
   ClientWidth     =   7270
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


Option Compare Text

Dim cellLeadAlertReq As Boolean
Dim qcManagerAlertReq As Boolean
Dim pmodManagerAlertReq As Boolean
Dim failedRoutines() As Variant
    '(0, i) -> Routine_Name
    '(1, i) -> Obs_Req
    '(2, i) -> Obs_Found
'failedRoutines(0, index) = Me.RoutineFrame(location).Caption
'    failedRoutines(1, index) = Me.ObsReq(location).Caption
'    failedRoutines(2, index) = Me.ObsFound(location).Caption


Private Sub FinalAQLHeader_Click()

End Sub

'****************************************************************************************
'               UserForm Callbacks
'****************************************************************************************

Private Sub UserForm_Initialize()
    
    Call SetActivePrinter
    
    Me.ProdQty.Caption = format(RibbonCommands.job_json("Qty Complete"), "#,###")

    'Set the AQL
    If RibbonCommands.job_json("AQL") Then Me.Controls("AQL").Caption = format(RibbonCommands.job_json("AQL"), "#.00")
    If RibbonCommands.job_json("FINAL_AQL") Then
        Me.Controls("FinalAQLHeader").Visible = True
        With Me.Controls("FinalAQL")
            .Caption = format(RibbonCommands.job_json("FINAL_AQL"), "#.00")
            .Visible = True
        End With
    End If

    'Set the required routines for the part
    'Also set the required observations for the routine
    Dim i As Integer
    For Each rt In RibbonCommands.job_json("Runs")
    

        ' Reset Controls
        With Me.RoutineFrame.Controls(i)
            .Caption = rt("RtName") 'set all of the routines that COULD be applicable
            .ForeColor = RGB(0, 0, 0)
            .Visible = True
        End With
        
        ' Set the obs results
        With Me.ObsFound.Controls(i)
            .Caption = rt("Passed Inspections")
            .Visible = True
        End With
        With Me.ObsReq.Controls(i)
            .Caption = rt("Required Inspections")
            .Visible = True
        End With
        
        ' Wasn't a required routine, but someone made it anyway
        If rt("Created") And rt("Required Inspections") = 0 Then
            Me.RoutineFrame.Controls(i).ForeColor = &H8000000D
            Me.ResultFrame.Controls(i).Visible = False
            
            GoTo UniqueRoutineErr
        End If
        
    
        'If a routine was never created, gray it out
        If rt("Created") = False Then
            With Me.RoutineFrame.Controls(i)
                .ForeColor = RGB(128, 128, 128)
            End With
        End If
        
        
        If rt("Passed") = True Then
            With Me.ObsFound.Controls(i)
                .Caption = rt("Passed Inspections")
                .Visible = True
            End With
            With Me.ObsReq.Controls(i)
                .Caption = rt("Required Inspections")
                .Visible = True
            End With
            GoTo NextRt
        ElseIf rt("Required Inspections") = 0 Then
            ' We have a routine that was never required to be inspected, make the text gray and strikeout
'             With Me.ObsReq.Controls(i)
'                .Caption = rt("Passed Inspections")
'                .Visible = True
'            End With
            Me.RoutineFrame.Controls(i).Font.Strikethrough = True
            Me.ResultFrame.Controls(i).Visible = False
'            hideResult location:=i

        Else
            ' We have an actual failure
            Call setFailure(location:=i, routine:=rt("Type"))
        End If

NextRt:
        i = i + 1
    Next rt
    
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
    
    
    If (cellLeadAlertReq Or qcManagerAlertReq Or pmodManagerAlertReq) Then
        'TODO: if we have only the 1XShift inspection and its not a true failure
            'Then we should enable everything and continue to set focus to the email button
        Me.PrintButton.Enabled = False
        Me.EmailButton.Enabled = True
        Me.EmailButton.SetFocus
    Else
        Me.PrintButton.Enabled = True
        Me.EmailButton.Enabled = False
        Me.PrintButton.SetFocus
    End If
    
    Exit Sub
    

UniqueRoutineErr:
   MsgBox "Application found this routine: " & rt("RtName") & _
                vbCrLf & "Which doesn't match any of our required routines" & _
                vbCrLf & "If a routine name changed, it could cause misalignment here", vbInformation
   GoTo NextRt
       
End Sub


Private Sub EmailButton_Click()
    Dim cells() As Variant
    Dim machines() As Variant
    Dim shiftDetails() As Variant
        '(0,i) -> OpCode
        '(1,i) -> OpNum
        '(2,i) -> ShiftDetails
    Dim shiftTraceability() As Variant
        '(0,i) -> ObsTimestamp
        '(1,i) -> EmpID
        '(2,i) -> Obs#
        '(3,i) -> Pass / Fail

    
    'If there are no machining operations required, then we dont need to query for a machine name
    If RibbonCommands.job_json Is Nothing Then
        ReDim Preserve machines(0)
        machines(0) = "[Outsourced Machining]"
        GoTo 10
    End If
    
    For i = 0 To UBound(failedRoutines, 2)
        Dim opInfo() As Variant
        opInfo = RibbonCommands.GetMachiningOpInfo(failedRoutines(0, i))
        
        'If its a 1XShift Routine, then add to our list of shift details
        If failedRoutines(0, i) Like "*1XSHIFT*" Then
            If (Not shiftDetails) = -1 Then
                ReDim Preserve shiftDetails(2, 0)
                shiftDetails(0, 0) = opInfo(0) 'OpCode
                shiftDetails(1, 0) = opInfo(1) 'OpSeq
                Set shiftDetails(2, 0) = HTTPconnections.Get1XSHIFTDetails(job_name:=RibbonCommands.jobNumUcase, op_num:=opInfo(1))
            Else
                ReDim Preserve shiftDetails(2, UBound(shiftDetails, 2) + 1)
                shiftDetails(0, 0) = opInfo(0) 'OpCode
                shiftDetails(1, 0) = opInfo(1) 'OpSeq
                Set shiftDetails(2, 0) = HTTPconnections.Get1XSHIFTDetails(job_name:=RibbonCommands.jobNumUcase, op_num:=opInfo(1))
            End If
            
            'We also want to populate the Inspections taken for the 1XSHIFT
            If (Not shiftTraceability) = -1 Then
                ReDim Preserve shiftTraceability(2, 0)
                shiftTraceability(0, 0) = RibbonCommands.jobNumUcase
                shiftTraceability(1, 0) = failedRoutines(0, i)
                Set shiftTraceability(2, 0) = HTTPconnections.GetAllFeatureTraceabilityData(job_name:=RibbonCommands.jobNumUcase, routine_name:=CStr(failedRoutines(0, i)))
'                shiftTraceability(2, 0) = DatabaseModule.GetAllFeatureTraceabilityData(jobNum:=RibbonCommands.jobNumUcase, routine:=CStr(failedRoutines(0, i)), FILL_EMP_IDS:=True)
            Else
                ReDim Preserve shiftTraceability(2, UBound(shiftTraceability, 2) + 1)
                shiftTraceability(0, UBound(shiftTraceability, 2)) = RibbonCommands.jobNumUcase
                shiftTraceability(1, UBound(shiftTraceability, 2)) = failedRoutines(0, i)
                Set shiftTraceability(2, 0) = HTTPconnections.GetAllFeatureTraceabilityData(job_name:=RibbonCommands.jobNumUcase, routine_name:=CStr(failedRoutines(0, i)))
            End If
            
        End If
                    
        'Add to our list of machines and cells responsible for the failures
        If ((Not cells) = -1 And (Not machines) = -1) Then
            ReDim Preserve cells(0)
            ReDim Preserve machines(0)
            cells(0) = opInfo(2)
            machines(0) = opInfo(3)
        Else
            If Not (IsNumeric(Application.Match(opInfo(2), cells, 0))) Then 'If the cell is not already in our list of cells, add it
                ReDim Preserve cells(UBound(cells) + 1)
                cells(UBound(cells)) = opInfo(2)
            End If
            If Not (IsNumeric(Application.Match(opInfo(3), machines, 0))) Then 'If the machines is not already in our list of machines, add it
                ReDim Preserve machines(UBound(machines) + 1)
                machines(UBound(machines)) = opInfo(3)
            End If
        End If
        
    Next i

    Dim cellLeadEmail As String
    For i = 0 To UBound(cells)
        cellLeadEmail = cellLeadEmail & HTTPconnections.GetCellLeadEmail(cell:=cells(i)) & ";"
    Next i
10
    Dim machineList As String
    For i = 0 To UBound(machines)
        machineList = machineList & machines(i)
        If i <> UBound(machines) Then machineList = machineList & ","
    Next i
        
    Call ExcelHelpers.CreateEmail(qcManager:=qcManagerAlertReq, pmodManager:=pmodManagerAlertReq, cellLead:=cellLeadAlertReq, cellLeadEmail:=cellLeadEmail, _
                                    jobNum:=RibbonCommands.jobNumUcase, machine:=machineList, failInfo:=failedRoutines, _
                                    shiftDetails:=shiftDetails, shiftTraceability:=shiftTraceability)

End Sub

Private Sub PrintButton_Click()
    Unload Me
    Call RibbonCommands.IterPrintRoutines
End Sub



Private Sub ChangePrinterButton_Click()
    If (Application.Dialogs(xlDialogPrinterSetup).Show) Then
        Call SetActivePrinter
    End If
End Sub

Private Sub ForcePrintButton_Click()
    'Override locking out our users from printing when routines fail
    Dim result As String
    result = Me.PasswordTextBox.Text
    
    If result = DataSources.ENABLE_PRINTING_PASS Then
        Me.PrintButton.Enabled = True
        Me.ForcePrintButton.Enabled = False
        MsgBox ("Printing Enabled")
    Else
        result = MsgBox("Incorrect Password", vbCritical)
        
    End If
    Me.PasswordTextBox.Text = vbNullString

End Sub




'****************************************************************************************
'               Extra Functions
'****************************************************************************************

Private Sub setFailure(location As Variant, routine As Variant)
    'Set failure picture
    With Me.ResultFrame.Controls(location)
        .Picture = LoadPicture(DataSources.FAIL_IMG_PATH)
        .Height = 15
        .Width = 15
        .Top = .Top + 4
    End With
    
    'If there was a problem with and 'FI' Routine like 'FI_DIM' then we need to alert the QC manager, otherwise we need to alert the cell lead
    If (InStr(routine, "FI") > 0) Then
        qcManagerAlertReq = True
    ElseIf (InStr(routine, "IP_ASSY") > 0) Then
        pmodManagerAlertReq = True
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





