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
Public customer As String
Public partNum As String
Public rev As String
Public partDesc As String
Public drawNum As String
Public ProdQty As Integer
Public dateTravelerPrinted As String
Public isShortRunEnabled As Boolean, lowerBoundCutoff As Integer, lowerBoundInspections As Integer
Public samplingSize As String, custAQL As String
Public parentChildSamplingSize As String, parentChildFinalAQL As String
Public IsChildJob As Boolean  'example: NV18209-2
Public IsParentJob As Boolean  'example: NV18209


'Epicor Operation-Specific JobInfo
Public multiMachinePart As Boolean
Public machineStageMissing As Boolean
Public missingLevels() As Integer     'For use in ThisWorkbook
Public partOperations() As Variant
    '(0,i) -> JobNum
    '(1,i) -> OprSeq
    '(2,i) -> OpCode
Public jobOperations() As Variant
    '(0,i) -> JobNum
    '(1,i) -> setupType
    '(2,i) -> Machine
    '(3,i) -> Cell
    '(4,i) -> OprSeq
    '(5,i) -> OpCode
'Routines for the part / Routines that we've run
Public partRoutineList() As Variant
    '(0,i) -> RoutineName
Public runRoutineList() As Variant
    '(0,i) -> RoutineName
    '(1,i) -> RunStatus
    '(2,i) -> ObsFound
    '(3,i) -> setupType
    '(4,i) -> machine
    '(5,i) -> cell

'Features and Measurement Information, applicable to the currently selected Routine
Dim featureHeaderInfo() As Variant
    '(0,i) -> FeatureName
    '(1,i) -> Description
    '(2,i) -> LTol
    '(3,i) -> Target
    '(4,i) -> UTol
    '(5,i) -> Insp Method
    '(6,i) -> Attribute / Variable
    '(7,i) -> Attribute Tolerance
    '(8,i) -> Balloon Num
Dim featureMeasuredValues() As Variant
    '(n,m) dimensional array where..
        'n -> number of features
        'm -> number of observations
Dim featureTraceabilityInfo() As Variant
    '(0,i) -> ObsTimestamp
    '(1,i) -> EmpID
    '(2,i) -> Obs#
    '(3,i) -> Pass / Fail


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
    Call SetJobVariables(jobNum:=jobNumUcase) 'If this errors out, just clears everything
    
    partOperations = DatabaseModule.GetPartOperationInfo(jobNumUcase)
    jobOperations = DatabaseModule.GetJobOperationInfo(jobNumUcase)
    
    'If there are no inside mach ops or there are less then there should be according to the MoM, flag that a stage is missing
    If ((Not jobOperations) = -1 And (Not Not partOperations)) Then
        machineStageMissing = True
    ElseIf (Not Not jobOperations) And (Not Not partOperations) Then
        machineStageMissing = (UBound(jobOperations, 2) < UBound(partOperations, 2))
    End If
    
SkipUbound:
    'If ops are missing, we need to determine which ones so we can ignore those respective routines later.
    If machineStageMissing = True Then
        'If we normally required mach ops for the part and we have some mach ops for this job..
        If (Not Not jobOperations) And (Not Not partOperations) Then
            For i = 0 To UBound(partOperations, 2)
                For j = 0 To UBound(jobOperations, 2)
                    If (partOperations(1, i) = jobOperations(4, j)) And (partOperations(2, i) = jobOperations(5, j)) Then
                        'If the Op# and Op Codes Match, then we dont need to do anything here
                        GoTo Nexti
                    End If
                Next j
                'Otherwise we couldnt find our part operation in the list of job operations
                If (Not missingLevels) = -1 Then
                    ReDim Preserve missingLevels(0)
                    missingLevels(0) = i
                Else
                    ReDim Preserve missingLevels(UBound(missingLevels) + 1)
                    missingLevels(UBound(missingLevels)) = i
                End If
Nexti:
            Next i
        'We have no machining operations for the job, determine which ops are missing
        ElseIf ((Not jobOperations) = -1) And (Not Not partOperations) Then
            For i = 0 To UBound(partOperations, 2)
                If (Not missingLevels) = -1 Then
                    ReDim Preserve missingLevels(0)
                    missingLevels(0) = i
                Else
                    ReDim Preserve missingLevels(UBound(missingLevels) + 1)
                    missingLevels(UBound(missingLevels)) = i
                End If
            Next i
        End If
    End If
    
'If there are more Machining Ops in the Job than the Part, things aren't missing
    'But we may need to insert a Machining Op into partOperations for later
    
    On Error GoTo PartOperationsErr
AddPartOps:
    If (Not partOperations) = -1 And (Not Not jobOperations) Then
        GoTo CompareMachOps
    'Otherwise if we have both part and job machining operations AND more machining ops in the Job than the Part MoM
    ElseIf (Not Not partOperations) And (Not Not jobOperations) Then
        If UBound(jobOperations, 2) > UBound(partOperations, 2) Then
CompareMachOps:
            Dim partMoM() As Variant, levelCounter As Integer
            partMoM = DatabaseModule.GetAllPartOps(partNum:=partNum, rev:=rev)
            
            For i = 0 To UBound(partMoM, 2)
                If partMoM(1, i) = "OUT" Then  'If theres an outside op changed to a machining op
                    For j = 0 To UBound(jobOperations, 2)
                        If (jobOperations(5, j) = "SWISS" Or jobOperations(5, j) = "CNC") And (jobOperations(4, j) = partMoM(0, i)) Then
                            'Insert the into our PartOps at the appropriate index
                            If (Not partOperations) = -1 Then
                                ReDim Preserve partOperations(2, 0)
                                partOperations(0, 0) = jobNumUcase
                                partOperations(1, 0) = partMoM(0, i) ' OprSeq
                                partOperations(2, 0) = partMoM(1, i) ' OpCode
                            Else
                                'TODO: Create Heloer function to Insert at the Correct Position
                                partOperations = ExcelHelpers.InsertOpRow(partOperations, jobNumUcase, partMoM(0, i), partMoM(1, i))
                            End If
                        End If
                    Next j
                End If
            Next i
        End If
    End If
        
    On Error GoTo 10
    If (Not Not partOperations) Then
        If UBound(partOperations, 2) > 0 Then multiMachinePart = True
    End If
    
QueryRoutines:
    Dim tempRoutineArray() As Variant
    On Error GoTo ML_QueryErr:
    customer = DatabaseModule.GetCustomerName(jobNum:=jobNumUcase)
    partRoutineList = DatabaseModule.GetPartRoutineList(partNum, rev)
    tempRoutineArray = DatabaseModule.GetRunRoutineList(jobNumUcase)

    If ((Not tempRoutineArray) = -1) Then GoTo 20 'We didnt find any routines for the run
    
    'Pass the results of the temp to the runRoutine List, we're going to add other dimensions where we
        'Keep track of the #ObsFound, setupType, machine and cell
    ReDim Preserve runRoutineList(5, UBound(tempRoutineArray, 2))
    For i = 0 To UBound(tempRoutineArray, 2)
        runRoutineList(0, i) = tempRoutineArray(0, i)
        runRoutineList(1, i) = tempRoutineArray(1, i)
    Next i
        
    'For each routine created for this run, find how many PASSed observations there are
    'We need to filter out the failed ones because this value will be used by VettingForm in ObsFound
    For i = 0 To UBound(runRoutineList, 2)
        Dim routine As String
        routine = runRoutineList(0, i)
        Dim features() As Variant
        features = DatabaseModule.GetFeatureHeaderInfo(jobNum:=jobNumUcase, routine:=routine)

        'Add the number of found Observations
        Dim featureCount() As Variant
        featureCount = DatabaseModule.GetFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=routine, _
                                        delimFeatures:=JoinPivotFeatures(features), featureInfo:=features)
        If ((Not featureCount) = -1) Then 'If we get returned an empty array, then the value is 0
            runRoutineList(2, i) = 0
        Else
            If routine Like "*FI_DIM*" Or routine Like "*FI_OP*" Then
                'FI_DIM routines. Above we checked there we at least a single inspection
                'Furthermore, if we have variable features we need to make sure we have enough good inspections of those
                'Since we only ever do a single observation of the attribute features, looking at all features
                'as a whole, this would normally assume that only one good observation exists
                If Not DatabaseModule.IsAllAttribrute(routine:=routine) Then
                    featureCount = DatabaseModule.GetFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=routine, _
                                    delimFeatures:=JoinPivotFeatures(features), featureInfo:=features, IS_FI_DIM:=True)
                    runRoutineList(2, i) = UBound(featureCount, 2) + 1
                Else
                    runRoutineList(2, i) = UBound(featureCount, 2) + 1
                End If
            Else
                runRoutineList(2, i) = UBound(featureCount, 2) + 1
            End If
        End If
    Next i
    
    On Error GoTo RoutineLevelErr
    For i = 0 To UBound(runRoutineList, 2)
    
        If multiMachinePart And (Not Not jobOperations) Then
            Dim level As Integer
            level = GetMachiningLevel(routineName:=runRoutineList(0, i)) 'Is this the first machining op, the second?, etc
            'Theoretically shouldnt have to check if a op of that level exists, since somebody bothered to create the routine for it
            For j = 0 To UBound(jobOperations, 2)
                'If OpNum, OpCode of the part machining stage and the job operation, associate Operation attribute with the routine
                If (partOperations(1, level) = jobOperations(4, j)) And (partOperations(2, level) = jobOperations(5, j)) Then
                    runRoutineList(3, i) = jobOperations(1, j) 'setupType
                    runRoutineList(4, i) = jobOperations(2, j) 'machine
                    runRoutineList(5, i) = jobOperations(3, j) 'cell
                End If
            Next j
            
        ElseIf (Not jobOperations) = -1 Then 'If we didnt make the part inside, then these attributes don't apply
            runRoutineList(3, i) = "None" 'setupType
            runRoutineList(4, i) = "NA" 'machine
            runRoutineList(5, i) = "NA" 'cell
            
        Else
            'The part only has a single machining operation, this is the bread and butter situation
            runRoutineList(3, i) = jobOperations(1, 0) 'setupType
            runRoutineList(4, i) = jobOperations(2, 0) 'machine
            runRoutineList(5, i) = jobOperations(3, 0) 'cell
        End If
    Next i
    
    'Set our Ribbon Information to the first Routine in our list, invalidate this control later
    rtCombo_TextField = runRoutineList(0, 0)
    lblStatus_Text = runRoutineList(1, 0)
    rtCombo_Enabled = True

    'Set our check boxes, displaying the setup information to the operator
    Select Case runRoutineList(3, 0)
        Case "Full"
            chkFull_Pressed = True
        Case "Mini"
            chkMini_Pressed = True
        Case "None"
            chkNone_Pressed = True
        Case Else
            If Not IsChildJob Then GoTo SetupTypeUndefined
    End Select
    
    On Error GoTo ML_RoutineInfo
    Call SetFeatureVariables
    
20
    If toggAutoForm_Pressed And ProdQty <> 0 Then VettingForm.Show
10
    'Still gets called if we have an invalid job, it should clean the page and exit out
    Call SetWorkbookInformation

     'Standard updates that are always applicable, refresh the ribbon controls
    cusRibbon.InvalidateControl "chkFull"
    cusRibbon.InvalidateControl "chkMini"
    cusRibbon.InvalidateControl "chkNone"
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
   

    Exit Sub
    
RoutineLevelErr:
    MsgBox Prompt:="Error when attmepting to associate the machining information with a routine: " & vbCrLf & runRoutineList(0, i) & vbCrLf & Err.Description, Buttons:=vbExclamation
    GoTo 10

ML_QueryErr:
    MsgBox Prompt:="Error when querying for information: " & vbCrLf & Err.Description, Buttons:=vbExclamation
    GoTo 10
    
SetupTypeUndefined:
    MsgBox Prompt:="Could not resolve Setup Type (Full, Mini, None)" & vbCrLf & "check this value in Epicor and/or ask a QE", Buttons:=vbExclamation
    GoTo 10
    
PartOperationsErr:
    If Err.Number = vbObjectError + 5000 Then
        MsgBox Prompt:="Error when trying to add an OUTside operation that has the same OP Number as an already Existing operation" & vbCrLf & vbCrLf & Err.Description, Buttons:=vbCritical
    End If
    
    GoTo 10
    
ML_RoutineInfo:
    MsgBox Prompt:="Error on handling routine: " & routine & vbCrLf & "information is either missing or incorrect, alert a QE" & vbCrLf & Err.Description, Buttons:=vbExclamation
    GoTo 10
VettingFormErr:
    If Err.Number = vbObjectError + 9999 Then 'Handled error, just ignore
        GoTo 10
    Else
        MsgBox "Unidentified error encountered when loading Vetting Form" & vbCrLf & vbCrLf & Err.Description
        GoTo 10
    End If
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

Public Function GetMachiningLevel(routineName As Variant) As Integer
    'Attempt to Associate a RoutineName with the an index position of our operations array
    
    'If we dont have partOps, then we cant determine a level, something is wrong
    If (Not partOperations) = -1 Then Err.Raise Number:=vbObjectError + 4000
    
    'set the maximum level
    Dim maxLevel As Integer
    maxLevel = UBound(partOperations, 2)
    Dim routineSub As String
    
    On Error GoTo RoutineParsingErr
    routineSub = Split(routineName, partNum & "_" & rev & "_")(1) 'Get the text appearing after   Part_Rev_"
    
    If (InStr(routineSub, "FA") > 0) Or (InStr(routineSub, "IP") > 0) Then
        If (InStr(routineSub, "IP_ASSY") > 0) Then GoTo 10
        If routineSub Like "*_MILL" Then
            routineSub = Replace(routineSub, "_MILL", "")
        End If
        
        If (IsNumeric(Right(routineSub, 1))) Then
            Dim foundLevel As Integer
            foundLevel = CInt(Right(routineSub, 1)) - 1
            If foundLevel <= maxLevel Then
                GetMachiningLevel = foundLevel
            Else
                'Return an error here, this should be impossible
                Err.Raise Number:=vbObjectError + 2500, Description:="Routine's Machining Level Number exceeds the Number of Machining Operations Found"
            End If
        Else
            GetMachiningLevel = 0
        End If
    ElseIf InStr(routineSub, "FI") > 0 Then
        'If it is an FI routine we give it the maximum level. It doesn't really matter as we see in Vetting Form, but it makes the most sense
        'IP_ASSY will also be redirected here
10
        GetMachiningLevel = maxLevel
    Else
        GoTo RoutineParsingErr
    End If
    
    Exit Function
    
RoutineParsingErr:
    Err.Raise Number:=vbObjectError + 2500, Description:="Couldn't figure out, what machining operation " & routineName & _
        vbCrLf & "should belong to. Does not follow correct naming conventions" & vbCrLf & vbCrLf & Err.Description
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
    
'Always
        'When the we try to set feature info w/o any info the wb runs cleanup and then stops
    rtCombo_TextField = ""
    lblStatus_Text = ""
    Erase featureHeaderInfo
    Erase featureMeasuredValues
    Erase featureTraceabilityInfo
    
    
    If preserveRoutines Then Exit Sub
    
'Sometimes
        'Want to skip this (likely because user entered nonsense into the routineName box)
    rtCombo_Enabled = False
    jobNumUcase = UCase(Text)
    chkFull_Pressed = False
    chkMini_Pressed = False
    chkNone_Pressed = False
    
    'Keep Job Info
    partNum = vbNullString
    rev = vbNullString
    customer = vbNullString
    partDesc = vbNullString
    dateTravelerPrinted = vbNullString
    isShortRunEnabled = False
    multiMachinePart = False
    machineStageMissing = False
    samplingSize = vbNullString
    custAQL = vbNullString
    IsChildJob = False
    IsParentJob = False
    parentChildSamplingSize = vbNullString
    parentChildFinalAQL = vbNullString

    'Keep routines for ComboBox
    Erase partRoutineList
    Erase runRoutineList
    Erase partOperations
    Erase jobOperations
    Erase missingLevels

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



