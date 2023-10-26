Attribute VB_Name = "RibbonCommands"
'*************************************************************************************************
'
'   RibbonCommands
'       Event logic for the Custom Ribbon Controls
'       1. The JobID Field and our editText Form should be updated to be the same when chanes are applied
'       2. The RoutineSelection and our ComboBox should be updated to be the same when changes are applied
'       3. we should ask the DataBase Module to perform our check on whether a jobNumber actually exists and is valid
'*************************************************************************************************

Option Compare Text

'Epicor Universal Job Info
Public jobNumUcase As String
Public job_json As Dictionary
Public part_json As Dictionary

Public run_feature_json As Collection
Public run_data_json As Collection
Public run_traceability_json As Collection

' Public customer As String
' Public partNum As String
' Public rev As String
' Public partDesc As String
' Public drawNum As String
' Public ProdQty As Integer
' Public dateTravelerPrinted As String
' Public isShortRunEnabled As Boolean, lowerBoundCutoff As Integer, lowerBoundInspections As Integer
' Public samplingSize As String, custAQL As String
' Public parentChildSamplingSize As String, parentChildFinalAQL As String
' Public IsChildJob As Boolean  'example: NV18209-2
' Public IsParentJob As Boolean  'example: NV18209


'*************
'Features and Measurement Information, applicable to the currently selected Routine
Dim featureHeaderInfo() As Collection
Dim featureMeasuredValues() As Collection
Dim featureTraceabilityInfo As Collection


'***********Ribbon Controls**************
'   we store the Ribbon on startup and use it to "invalidate" the other controls later
'   which makes them call some of their callback functions
Dim cusRibbon As IRibbonUI

Dim lblStatus_Text As String
Dim rtCombo_TextField As String
Dim rtCombo_Enabled As Boolean

Private toggAutoForm_Pressed As Boolean
Public toggML7TestDB_Pressed As Boolean
Public toggShowAllObs_Pressed As Boolean

Public chkFull_Pressed As Boolean
Public chkMini_Pressed As Boolean
Public chkNone_Pressed As Boolean




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               UI Ribbon
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Ribbon_OnLoad(uiRibbon As IRibbonUI)
    Set cusRibbon = uiRibbon
    cusRibbon.ActivateTab "mlTab"
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Job Number EditTextField
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub jbEditText_onGetText(ByRef control As IRibbonControl, ByRef Text)
    Text = jobNumUcase
End Sub

Public Sub jbEditText_OnChange(ByRef control As Office.IRibbonControl, ByRef Text As String)
    'Reset the Variables
    Call ClearFeatureVariables
    jobNumUcase = UCase(Text)

    If Text = vbNullString Then GoTo 10
    
    On Error GoTo 10
    'Call the HTTP method and set the Variables
    Dim dhr_json As Dictionary
    Set dhr_json = HTTPconnections.ValidateDHR(job_num:=jobNumUcase)
    Set job_json = dhr_json("job_info")
    Set part_json = dhr_json("part_info")

    'Set our Ribbon Information to the first Routine in our list, invalidate this control later
    rtCombo_TextField = job_json("Runs")(1)("Name")
    
    'TODO: not currently pulling in the status of our runs...
    lblStatus_Text = job_json("Runs")(1)("Name")
    rtCombo_Enabled = True

    'Set our check boxes, displaying the setup information to the operator
    Select Case job_json("Operations")(1)("Setup Type")
        Case "Full"
            chkFull_Pressed = True
        Case "Mini"
            chkMini_Pressed = True
        Case "None"
            chkNone_Pressed = True
        Case Else
            If Not job_json("IsChildJob") Then GoTo SetupTypeUndefined
    End Select
    
    
20
    If toggAutoForm_Pressed And job_json("Qty Complete") <> 0 Then VettingForm.Show
    
'TODO: this will all eventually move to RibbonCommands.SetWorkbookInformation()
'TODO: need to be able to handle us having the attribute FI routine being the first one, this should all be done in a seperate function here
'TODO: should handle not having any operations here
'TODO: determine machining level, use THAT for Operation index and machine name
    ThisWorkbook.populateJobHeaders jobNum:=jobNumUcase, routine:=rtCombo_TextField, _
        customer:=job_json("Customer"), machine:=job_json("Operations")(1)("Machine"), partNum:=job_json("PartNum"), _
        rev:=job_json("RevisionNum"), partDesc:=job_json("PartDescription")
    
    
    
    ExcelHelpers.OpenDataValWB
    
    'User Closed the form, go and load up the inspection data for the first of the Routines
    If toggShowAllObs_Pressed Then  'Load all Observerations
        'TODO:
    
    Else  'Load only passed observations
        Set result = HTTPconnections.GetPassedInspData(jobNumUcase, rtCombo_TextField)
        Set run_feature_json = result("feature_info")
        Set run_data_json = result("insp_data")
        Set run_traceability_json = result("traceability")
    End If
    
    Call ThisWorkbook.populateReport(header_info:=run_feature_json, insp_data:=run_data_json, traceability:=run_traceability_json)
    
10
    'Still gets called if we have an invalid job, it should clean the page and exit out

    'Call SetWorkbookInformation

     'Standard updates that are always applicable, refresh the ribbon controls

    ' cusRibbon.InvalidateControl "chkFull"
    ' cusRibbon.InvalidateControl "chkMini"
    ' cusRibbon.InvalidateControl "chkNone"
    ' cusRibbon.InvalidateControl "rtCombo"
    ' cusRibbon.InvalidateControl "jbEditText"
    ' cusRibbon.InvalidateControl "lblStatus"
   
    Exit Sub
    
SetupTypeUndefined:
    MsgBox "Cannot Determine Setup Type of " & job_json("Operations")(1)("Setup Type") & vbCrLf & "Have this changed to the appriopriate type in Job Entry", vbCritical
    Exit Sub

    
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               RoutineName ComboBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub rtCombo_OnChange(ByRef control As Office.IRibbonControl, ByRef Text As Variant)

    'There doesn't seem to be a property to prevent the user from Hand-Typing into the ComboBox
    'So we have to make sure that the change is legitimate
    Dim validChange As Boolean
    validChange = False
    
    'iterate through our list of routines to see if the typed value is in there
    If Not Not runRoutineList Then
        For i = 0 To UBound(runRoutineList, 2)
            If Text = runRoutineList(0, i) Then
            
                'Erase old feature data
                validChange = True
                Exit For
            End If
        Next i
    End If
    
    'Erase the feature data but, not our Job Number or Job Routine List,
    'This means the user can still select from the drop-down and try again
    Call ClearFeatureVariables(preserveRoutines:=True)
    
    On Error GoTo 10
    If validChange = True Then
         'Set new active routine
        lblStatus_Text = runRoutineList(1, i)
        rtCombo_TextField = Text

        'Get new feature data with new active routine
        Call SetFeatureVariables
    End If
    
    'If there was new data we populate, if not then we end up clearing everything
    Call SetWorkbookInformation

10
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
    
End Sub

Public Sub rtCombo_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = rtCombo_Enabled
End Sub

Public Sub rtCombo_OnGetItemCount(ByRef control As Office.IRibbonControl, ByRef Count As Variant)
    If Not IsEmpty(runRoutineList) Then
        Count = UBound(runRoutineList, 2) + 1
    End If
End Sub

Public Sub rtCombo_OnGetItemLabel(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemLabel As Variant)
    ItemLabel = runRoutineList(0, index)
End Sub

Public Sub rtCombo_OnGetItemID(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemID As Variant)
    'Need to reference by ID? I guess not
End Sub

Public Sub rtCombo_OnGetText(ByRef control As Office.IRibbonControl, ByRef Text As Variant)
    'when we don't have any routines, use a placeholder string
    If Not Not runRoutineList Then
        Text = rtCombo_TextField
    Else
        Text = "[SELECT ROUTINE]"
    End If

End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               LoadForm Button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Callback(ByRef control As Office.IRibbonControl)
    If jobNumUcase = vbNullString Then
        MsgBox "No Job Currently Loaded", vbInformation
    ElseIf ProdQty = 0 Then
        MsgBox "There is no Production Qty for this Job" & vbCrLf & "Nothing to Verify", vbInformation
    Else
        VettingForm.Show
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Auto Load Form Toggle Button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub toggAutoForm_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    toggAutoForm_Pressed = isPressed
End Sub
Public Sub toggAutoForm_OnGetPressed(ByRef control As Office.IRibbonControl, ByRef ReturnedValue As Variant)
    
    ReturnedValue = True
    toggAutoForm_Pressed = True
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'              Show All Observations Toggle Buttom
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub allObs_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    toggShowAllObs_Pressed = isPressed
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               ML7 Test Database Toggle Button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub testDB_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    toggML7TestDB_Pressed = isPressed
    Call DatabaseModule.Close_Connections 'If we had a connection already open, need to invalidate it so we can connect to the TestDB
End Sub

Public Sub testDB_OnGetEnabled(ByRef control As Office.IRibbonControl, ByRef ReturnedValue As Variant)
    ReturnedValue = False
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               ML7 Test Database Toggle Button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub cleanSheets_Pressed(ByRef control As Office.IRibbonControl)
    ThisWorkbook.Cleanup
End Sub





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               RunStatus Label
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub lblStatus_OnGetLabel(ByRef control As Office.IRibbonControl, ByRef Label As Variant)
    If lblStatus_Text = vbNullString Then
        Label = ""
    Else
        Label = lblStatus_Text
    End If
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               JobType Check Boxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub chkFull_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
End Sub

Public Sub chkFull_OnGetPressed(ByRef control As IRibbonControl, ByRef pressed As Variant)
    pressed = chkFull_Pressed
End Sub

Public Sub chkMini_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
End Sub
Public Sub chkMini_OnGetPressed(ByRef control As IRibbonControl, ByRef pressed As Variant)
    pressed = chkMini_Pressed
End Sub
Public Sub chkNone_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
End Sub
Public Sub chkNone_OnGetPressed(ByRef control As IRibbonControl, ByRef pressed As Variant)
    pressed = chkNone_Pressed
End Sub





'****************************************************************************************
'               Extra Functions
'****************************************************************************************

'Called by Vetting form when user clicks on Print
Public Sub IterPrintRoutines()

    'Iterate through the Run's Routines, set the results and ask the workbook to print
    For i = 0 To UBound(runRoutineList, 2)
        rtCombo_TextField = runRoutineList(0, i)
        lblStatus_Text = runRoutineList(1, i)
        
        On Error GoTo 10
        Call SetFeatureVariables
        Call SetWorkbookInformation
        Call ThisWorkbook.PrintResults
    Next i

10
    'The Ribbon information will be updated to the last routine that was printed / activated
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
    
End Sub


Function JoinPivotFeatures(featureHeaderInfo() As Variant) As String

    'SQL Pivot tables will require us to specify what the columnns (part features) are, so that list needs to be dynamically generated
    Dim paramFeatures() As String
    ReDim Preserve paramFeatures(UBound(featureHeaderInfo, 2))
    For i = 0 To UBound(featureHeaderInfo, 2)
        paramFeatures(i) = "[" & featureHeaderInfo(0, i) & "]"
    Next i
    
    JoinPivotFeatures = Join(paramFeatures, ",")

End Function

Public Function GetMachiningOpInfo(routineName As Variant) As Variant()
    'Using the RunRoutine Name, find the it in our list of required PartRoutines and return the set Level
    
    Dim out_info(3) As Variant
                
    If routineName Like "*_FI_*" Then  'Max Level Routine
        Dim operation As String
        If routineName Like "*FI_DIM*" Then
            operation = "FDIM"
        ElseIf routineName Like "*FI_VIS*" Then
            operation = "FVIS"
        Else
            operation = "FI_OP"
        End If
      
        out_info(0) = operation 'opcode
        out_info(1) = 0 'opseq
        out_info(2) = "QC" 'cell
        out_info(2) = "QC" 'machin

    End If
    
    If part_json("part_routines").Count Then Err.Raise Number:=vbObjectError + 2500
    For Each job_rt In job_json("Runs")
        For Each part_rt In part_json("part_routines")
            If job_rt("Name") = part_rt("Name") Then
                
                out_info(0) = job_json("Operations")(part_rt("Level") - 1)("OpCode")
                out_info(1) = job_json("Operations")(part_rt("Level") - 1)("OpSeq")
                out_info(2) = job_json("Operations")(part_rt("Level") - 1)("Cell")
                out_info(2) = job_json("Operations")(part_rt("Level") - 1)("Machine")
            
                GetMachiningOpInfo = out_info
                
                Exit Function
            End If
        Next part_rt
    Next job_rt

RoutineNotFound:
    Err.Raise Number:=vbObjectError + 4000, Description:="Created Run of " & routineName & vbCrLf & "But couldn't find a matching required routine in MeasurLink." & _
        vbCrLf & "Can't determine what Operation the routine belongs to"
    
RoutineCountErr:
    Err.Raise Number:=vbObjectError + 2500, Description:="Couldn't figure out, what machining operation " & routineName & _
        vbCrLf & "should belong to. There were no required Part Routines" & vbCrLf & vbCrLf & Err.Description
End Function

Private Sub SetFeatureVariables()

    On Error GoTo Err1
    
    Dim isFI_DIM As Boolean
    Dim allAttr As Boolean
    If rtCombo_TextField Like "*FI_DIM*" Then
        isFI_DIM = True
        If DatabaseModule.IsAllAttribrute(routine:=rtCombo_TextField) Then
            allAttr = True
        End If
    End If

    featureHeaderInfo = DatabaseModule.GetFeatureHeaderInfo(jobNum:=jobNumUcase, routine:=rtCombo_TextField)
    
    'Should we filter or not filter observations shown based on Pass/Fail data
    'Having ShowAllObs pressed DOES NOT change the ObsFound value for the userform, that value is set in jbEditText
    If toggShowAllObs_Pressed Then
        featureTraceabilityInfo = DatabaseModule.GetAllFeatureTraceabilityData(jobNum:=jobNumUcase, routine:=rtCombo_TextField)
        featureMeasuredValues = DatabaseModule.GetAllFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=rtCombo_TextField, _
                                                delimFeatures:=JoinPivotFeatures(featureHeaderInfo))

    Else
            'If we have a FI_DIM with all Attr features then, we should leave the arrays unintialized
        If Not allAttr Then
            featureMeasuredValues = DatabaseModule.GetFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=rtCombo_TextField, _
                                                    delimFeatures:=JoinPivotFeatures(featureHeaderInfo), featureInfo:=featureHeaderInfo, IS_FI_DIM:=isFI_DIM)
            
            featureTraceabilityInfo = DatabaseModule.GetFeatureTraceabilityData(jobNum:=jobNumUcase, routine:=rtCombo_TextField, FI_DIM_ROUTINE:=isFI_DIM)
        End If
    End If
    
    Exit Sub
    
Err1:
    result = MsgBox("Could not set Job/Run information. Issue found at: " & vbCrLf & Err.Description, vbCritical)
    Err.Raise Number:=vbObjectError + 1000

End Sub

Private Sub ClearFeatureVariables(Optional preserveRoutines As Boolean)
    Set job_json = Nothing
    Set part_json = Nothing

End Sub

Private Sub SetJobVariables(jobNum As String)
    On Error GoTo jbInfoErr
    Dim jobInfo() As Variant
    
    jobInfo = DatabaseModule.GetJobInformation(JobID:=jobNum)
    
    'Add the components of the array to our variables
    partNum = jobInfo(2, 0)
    rev = jobInfo(3, 0)
    partDesc = jobInfo(5, 0)
    drawNum = jobInfo(6, 0)
    
    'If the prod Qty is null, its because we dont have a single complete Operation
    If VarType(jobInfo(7, 0)) = vbNull Then
        Dim result As Integer
        result = MsgBox("No Production Qty found, Likely because there isn't a completed Operation yet" & vbCrLf _
            & "Would you like to View the results for this job anyway?", vbYesNo)
        
        If result = vbNo Then
            Err.Raise Number:=vbObjectError + 2100, Description:="No Operations have been completed for this Job." & vbCrLf & "Cant Verify Inspections"
        Else
            GoTo skipQty
        End If
    End If
    
    ProdQty = jobInfo(7, 0)
skipQty:

    If IsNull(jobInfo(8, 0)) Then Err.Raise vbObjectError + 1000, Description:="Traveler has not been printed yet" & vbCrLf & "Can't Run Report"
    dateTravelerPrinted = jobInfo(8, 0)
    

    Dim shortRunInfo() As Variant
    shortRunInfo = GetFlaggedShortRunIR(drawNum:=drawNum, rev:=rev, datePrinted:=dateTravelerPrinted)
    If Not Not shortRunInfo Then
        If shortRunInfo(2, 0) >= 0 Then  'If the job was printed after the IR was flagged for short run Inspections
            isShortRunEnabled = True
        End If
    End If
        
        'Check if a job is a Parent Job
    IsParentJob = DatabaseModule.IsParentJob(JobNumber:=jobNum)
    If IsParentJob Then Exit Sub 'Prod Qty should exclude negative transaction adjustments
    
    
'        ProdQty = DatabaseModule.GetParentProdQty(JobNumber:=jobnum)
'        Exit Sub
'    End If


        'Check if the Job is a Child Job Instance, only check if not already a parent job
    If Not IsNumeric(Left(jobNum, 1)) And InStr(jobNum, "-") > 0 Then
        Dim jobArr() As String
        jobArr = Split(jobNum, "-")
        If UBound(jobArr) = 1 Then
            If (Len(jobArr(1)) = 1 Or Len(jobArr(1)) = 2) And IsNumeric(jobArr(1)) Then IsChildJob = True
        End If
    End If
    
    Exit Sub

jbInfoErr:
    'If the recordSet is empty
    If Err.Number = vbObjectError + 2000 Then
        MsgBox ("Not A Valid Job Number")
    Else
    'Otherwise we encountered a different problem
        result = MsgBox(Err.Description, vbExclamation)
    End If
    
    'Either way, reset the job number and invalidate the controls
    jobNumUcase = ""
    Err.Raise Number:=Err.Number, Description:="SetJobVariables" & vbCrLf & Err.Description


End Sub

Private Sub SetWorkbookInformation()
    Dim index As Integer
    Dim machine As String
    Dim attFeatHeaders() As Variant
    Dim attFeatResults() As Variant
    Dim attFeatTraceability() As Variant
    Dim resultsFailed As Boolean
    Dim noTraceability As Boolean
    Dim noVariables As Boolean
    
    If (Not Not runRoutineList) Then
        index = GetRoutineIndex(rtCombo_TextField)
        machine = runRoutineList(4, index)
        
'Conditionally handle information for FI_DIM, FI_VIS, FI_RECINSP. Breakoff attributes into their own sheet
        If InStr(rtCombo_TextField, "FI_DIM") > 0 Or InStr(rtCombo_TextField, "FI_VIS") > 0 Or InStr(rtCombo_TextField, "RECINSP") > 0 Then
            Dim ogLen As Integer
            ogLen = UBound(featureHeaderInfo, 2)
            Call SliceVariableInformation(noVariables)  'Remove attribute features from our array of information
                
                'if the array if not initialized, then we either have no variable features or there's no features at all
            If noVariables Then GoTo SetFIattr
            ' If (Not featureHeaderInfo) = -1 Then GoTo SetFIattr
            
                'If the array size is unchanged after slicing then there are only variable features, dont bother querying for attr below
            If ogLen = UBound(featureHeaderInfo, 2) - LBound(featureHeaderInfo, 2) Then GoTo SetWBinfo
SetFIattr:

            
            attFeatHeaders = DatabaseModule.GetFinalAttrHeaders(jobNumUcase, rtCombo_TextField)
            If (Not attFeatHeaders) = -1 Then GoTo SetWBinfo 'either Routine isnt real or was never created
            
            
            attFeatResults = DatabaseModule.GetFinalAttrResults(jobNumUcase, rtCombo_TextField)
            If VarType(attFeatResults(0, 0)) = vbNull Then
                resultsFailed = True
                MsgBox "No Attribute Inspections taken for routine" & vbCrLf & rtCombo_TextField
            
            ElseIf attFeatResults(0, 0) > 0 Then
                resultsFailed = True
                MsgBox "One of the most recent inspections for routine" & vbCrLf & rtCombo_TextField & vbCrLf _
                    & "Contains a Fail. Additional inspection is needed", vbCritical
            
            ElseIf attFeatResults(1, 0) <> UBound(attFeatHeaders, 2) + 1 Then
                resultsFailed = True
                MsgBox "Not all attribute features have been inspected for routine" & vbCrLf & rtCombo_TextField _
                    & vbCrLf & "Please have QC review the routine for this Job", vbCritical
            End If
                        
            attFeatTraceability = DatabaseModule.GetFinalAttrTraceability(jobNumUcase, rtCombo_TextField)
            
                'If theres no tracability or we dont have traceability data for every feature
                'This seems to be Caused by misalignment of StartObsId and that ObsId not existing in teh RunData
                'Perhaps caused by someone starting to take an observation and then not completing that value
            If (Not attFeatTraceability) = -1 Or UBound(attFeatTraceability, 2) <> UBound(attFeatHeaders, 2) Then
                noTraceability = True
                MsgBox "Traceability Information Missing on Attribute Features" & vbCrLf & "Try doing another row of Passes to resolve this issue", vbCritical
            End If
             
        End If
    Else
        'If our runRoutineList is empty, then ThisWorkbook will end up just running the cleanup anyway
        machine = ""
    End If
    
SetWBinfo:
    On Error GoTo wbErr
    ExcelHelpers.OpenDataValWB
    
    Call ThisWorkbook.populateJobHeaders(jobNum:=jobNumUcase, routine:=rtCombo_TextField, customer:=customer, _
                                            machine:=machine, partNum:=partNum, rev:=rev, partDesc:=partDesc)
    Call ThisWorkbook.populateReport(featureInfo:=featureHeaderInfo, featureMeasurements:=featureMeasuredValues, _
                                        featureTraceability:=featureTraceabilityInfo)
            
    Call ThisWorkbook.populateAttrSheet(attFeatHeaders:=attFeatHeaders, attFeatResults:=attFeatResults, _
            attFeatTraceability:=attFeatTraceability, noResults:=resultsFailed, noTraceability:=noTraceability, noVariables:=noVariables)
            
    ExcelHelpers.CloseDataValWB
    Exit Sub
wbErr:
    ExcelHelpers.CloseDataValWB
    result = MsgBox("Error Occurred at Sub: RibbonCommands.SetWorkbookInformation", vbCritical)
    Err.Raise Number:=vbObjectError + 1200
    
End Sub

'Called by SetWorkbookInformation
    'Take the Global values for featureHeaderInfo, featureMeasuredValues, and featureTraceabilityInfo and slice out the attribute features
    'Can test with... SS0245
    
    'Params
        'noVariables(byref Boolean) -> if true, then only attribute features exist in our feature arrays
Private Sub SliceVariableInformation(ByRef noVariables As Boolean)
    If (Not featureHeaderInfo) = -1 Then Exit Sub
    If (Not featureMeasuredValues) = -1 Then
        noVariables = True
        Exit Sub
    End If
    
    Dim varCols() As Variant
    Dim i As Integer
    
    For i = 0 To UBound(featureHeaderInfo, 2)
        If featureHeaderInfo(6, i) <> "Variable" Then GoTo continue
            
        If (Not varCols) = -1 Then
            ReDim Preserve varCols(0)
            varCols(0) = i + 1
        Else
            ReDim Preserve varCols(UBound(varCols) + 1)
            varCols(UBound(varCols)) = i + 1
        End If
continue:
    Next i
    
        'If there are only Attr features. Erase, as we will be querying for them in another format.
    If (Not varCols) = -1 Then
        noVariables = True
        Erase featureHeaderInfo
        Erase featureMeasuredValues
        Erase featureTraceabilityInfo
        Exit Sub
    End If
    
    'If this doesn't make sense, its because VBA slices arrays as though they were 1-indexed, and of course...
        'the ADODB library returns records as 0-indexed.
        'transposing the array twice returns the same array, but 1-indexed
        'but we're also taking care to create our temporary arrays as 1-indexed
    varCols = Application.Transpose(Application.Transpose(varCols))
    
    
    Dim tempCols() As Variant
    tempCols = ExcelHelpers.nRange(1, UBound(featureMeasuredValues, 2) + 1) 'This should get each measurement taken.....
    Dim tempRow() As Variant
    tempRow = Application.Transpose(Array(1, 2, 3, 4, 5, 6, 7, 8, 9))
    
        'Get all the header information for the Variable columns
        'Array -> Columns that we want to grab, in this case, everything
        'varCols -> The rows of variable features
    featureHeaderInfo = Application.index(featureHeaderInfo, Application.Transpose(Array(1, 2, 3, 4, 5, 6, 7, 8, 9)), varCols)
    
        'Get all the measurement information for the Variable columns
    varCols = ExcelHelpers.updateForPivotSlice(varCols)
    featureMeasuredValues = ExcelHelpers.fill_null(featureMeasuredValues)
        'varCols -> the columns that represent our variable Features
        'tempCols -> every single observation for those features
    featureMeasuredValues = Application.index(featureMeasuredValues, Application.Transpose(varCols), tempCols)
End Sub



Public Function GetRoutineIndex(routineName As String) As Integer
    'If we dont' have a runRoutineList or didnt find a routine of the given name, then return 99 as the index which we will
    'Later turn into a flag to in VettingForm when it is trying to figure out which partRoutines, not runRoutines, apply.
    If ((Not runRoutineList) = -1) Then GoTo 10

    For i = 0 To UBound(runRoutineList, 2)
        If routineName = runRoutineList(0, i) Then GoTo FoundRoutine
    Next i
10
    GetRoutineIndex = 99
    Exit Function
    
FoundRoutine:
    GetRoutineIndex = i
End Function





