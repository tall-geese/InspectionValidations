Attribute VB_Name = "RibbonCommands"

'*************************************************************************************************
'
'   RibbonCommands
'       Event logic for the Custom Ribbon Controls
'       1. The JobID Field and our editText Form should be updated to be the same when chanes are applied
'       2. The RoutineSelection and our ComboBox should be updated to be the same when changes are applied
'       3. we should ask the DataBase Module to perform our check on whether a jobNumber actually exists and is valid
'*************************************************************************************************

'Epicor Universal Job Info
Public jobNumUcase As String
Public customer As String
Public partNum As String
Public rev As String
Public partDesc As String
Public drawNum As String
Public prodQty As Integer

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
    '(4,i) -> ProdQty
    '(5,i) -> OprSeq
    '(6,i) -> OpCode
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
    '(0,i) -> Balloon#
    '(1,i) -> Description
    '(2,i) -> LTol
    '(3,i) -> Target
    '(4,i) -> UTol
    '(5,i) -> Insp Method
    '(6,i) -> Attribute / Variable
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
    Call SetJobVariables(jobNum:=jobNumUcase)
    
    'TODO: Determine the number of Swiss/CNC operations normally required by the part, set a flag if is a multimachining op Part
    'Determine if any of our machining operations are missing as a result of outside operations, set a flag once again
        'Find out, based on the previous data - which of our operations are missing exactly
        'If we know a level is missing, then we should pass that data later on in thi function when collecting operation-Level
            'Information like machine, cell. They will need to either fill NA or use another level's information
    
    partOperations = DatabaseModule.GetPartOperationInfo(jobNumUcase)
    If UBound(partOperations, 2) > 0 Then multiMachinePart = True
    
    jobOperations = DatabaseModule.GetJobOperationInfo(jobNumUcase)
    'If there were not job machining ops or less machining ops then we expected
    If ((Not jobOperations) = -1) Then
        machineStageMissing = True
        GoTo SkipUbound
    End If
    If (UBound(jobOperations, 2) < UBound(partOperations, 2)) Then machineStageMissing = True
    
SkipUbound:
    'If ops are missing, we need to determine which ones so we can ignore those respective routines later.
    If machineStageMissing = True Then
        If (Not Not jobOperations) And (Not Not partOperations) Then 'we have a list of part operations and job operations
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
            
        ElseIf (Not Not jobOperations) And ((Not partOperations) = -1) Then 'This doesnt make sense, we have more machining ops then expected
            result = MsgBox("There are more machining operations for this job than expected. Cannot process", vbCritical)
            GoTo 10
        ElseIf ((Not jobOperations) = -1) And (Not Not partOperations) Then 'We have no machining operations, set the missing machining ops
            For i = 0 To UBound(partOperations, 2)
                If (Not missingLevels) = -1 Then
                    ReDim Preserve missingLevels(0)
                    missingLevels(0) = i
                Else
                    ReDim Preserve missingLevels(UBound(missingLevels) + 1)
                    missingLevels(UBound(missingLevels)) = i
                End If
            Next i
        ElseIf ((Not jobOperations) = -1) And ((Not partOperations) = -1) Then  'neither have been initialized, no one should call for
                                                                            'manufacturing routines anyway, skip ahead
            'theoretically we could just leav this alone
            'TODO: maybe we should set machineStageMissing back to False, since really nothing is missing now
        End If
    End If
    
    Dim tempRoutineArray() As Variant
    On Error GoTo ML_QueryErr:
    customer = DatabaseModule.GetCustomerName(jobNum:=jobNumUcase)
    partRoutineList = DatabaseModule.GetPartRoutineList(partNum, rev)
    tempRoutineArray = DatabaseModule.GetRunRoutineList(jobNumUcase)

    If ((Not tempRoutineArray) = -1) Then GoTo 20 'We didnt find any routines for the run
    
    'Pass the results of the temp to the runRoutine List, we're going to add another dimension where we
        'Keep track of the #ObsFound for each routine and use this later in the UserForm
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
        If ((Not featureCount) = -1) Then
            runRoutineList(2, i) = 0
        Else
            runRoutineList(2, i) = UBound(featureCount, 2) + 1
        End If
'        runRoutineList(2, i) = UBound(DatabaseModule.GetFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=routine, _
'                                        delimFeatures:=JoinPivotFeatures(features), featureInfo:=features), 2) + 1
        'TODO: Add run mahciningLevel, cell, machine, setup Type, Completed Qty
        'We should be using the 'FA', 'FI', code in the routine name to determine What opCode we should be searching in
        'Let another functin handle this and come up with the determined level
        'Use SD1168 to validate this, FVIS has about 400 parts Less
    Next i
    
    'TODO: set up an Epicor Read Error
    For i = 0 To UBound(runRoutineList, 2)
    
        If multiMachinePart And (Not Not jobOperations) Then
            'Is this the first machining op, the second?, etc
            Dim level As Integer
            level = GetMachiningLevel(routineName:=runRoutineList(0, i))
            'Theoretically shouldnt have to check if a op of that level exists, since somebody bothered to create the routine for it
            For j = 0 To UBound(jobOperations, 2)
                If (partOperations(1, level) = jobOperations(4, j)) And (partOperations(2, level) = jobOperations(5, j)) Then
'                    runRoutineList(3, i) = jobOperations(4, j) 'ProdQty
                    runRoutineList(3, i) = jobOperations(1, j) 'setupType
                    runRoutineList(4, i) = jobOperations(2, j) 'machine
                    runRoutineList(5, i) = jobOperations(3, j) 'cell
                End If
            Next j
            
        ElseIf (Not jobOperations) = -1 Then
            'The part has machining operations but we did them all outside
            'So in this situation, we don't have a great place to pull the acceptable quantity to base the AQL off of,
            'BUT we can try using the MAX() or greatest of the sum of the operations
            
'            Dim tempArray() As Variant
'            tempArray = DatabaseModule.GetGreatestOpQty(jobNumUcase)(1, 0)
'            runRoutineList(3, i) = DatabaseModule.GetGreatestOpQty(jobNumUcase)(1, 0)
            runRoutineList(3, i) = "None" 'setupType
            runRoutineList(4, i) = "NA" 'machine
            runRoutineList(5, i) = "NA" 'cell
            
        Else
            'The part only has a single machining operation, this is the bread and butter situation
'            runRoutineList(3, i) = jobOperations(4, 0) 'ProdQty
            runRoutineList(3, i) = jobOperations(1, 0) 'setupType
            runRoutineList(4, i) = jobOperations(2, 0) 'machine
            runRoutineList(5, i) = jobOperations(3, 0) 'cell
        End If
    Next i
    
    
    
    
    
    'Set our Ribbon Information to the first Routine in our list, invalidate this control later
    rtCombo_TextField = runRoutineList(0, 0)
    lblStatus_Text = runRoutineList(1, 0)
    rtCombo_Enabled = True

    'TODO: we dont have this variable anymore, need to switch on runRoutineList(4,0)
    Select Case runRoutineList(3, 0)
        Case "Full"
            chkFull_Pressed = True
        Case "Mini"
            chkMini_Pressed = True
        Case "None"
            chkNone_Pressed = True
        Case Else
            GoTo SetupTypeUndefined
    End Select
    
    On Error GoTo ML_RoutineInfo
    Call SetFeatureVariables
    
    On Error GoTo 10
    Call SetWorkbookInformation
20
    If toggAutoForm_Pressed Then VettingForm.Show
10
    'TODO: once again need to resolve what happens in the event of an invalid job num


     'Standard updates that are always applicable
    cusRibbon.InvalidateControl "chkFull"
    cusRibbon.InvalidateControl "chkMini"
    cusRibbon.InvalidateControl "chkNone"
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
   

    Exit Sub

ML_QueryErr:
    MsgBox Prompt:="Error when querying for information: " & vbCrLf & Err.description, Buttons:=vbExclamation
    GoTo 10
    
SetupTypeUndefined:
    MsgBox Prompt:="Could not resolve Setup Type (Full, Mini, None)" & vbCrLf & "check this value in Epicor and/or ask a QE", Buttons:=vbExclamation
    GoTo 10
    
ML_RoutineInfo:
    MsgBox Prompt:="Error on handling routine: " & routine & vbCrLf & "information is either missing or incorrect, alert a QE" & vbCrLf & Err.description, Buttons:=vbExclamation
    GoTo 10
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
    'Believe it or not, this is the proper way to check if a Variant Array has been initialized
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
    VettingForm.Show
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
     'TODO: what do we do if a part never has partOperations becuase its is always manufactured outside
    'set the maximum level
    Dim maxLevel As Integer
    maxLevel = UBound(partOperations, 2)
    Dim routineSub As String
    
    'TODO: set error handling here in case the routineName does not make sense
    routineSub = Split(routineName, partNum & "_" & rev & "_")(1)
    
    If (InStr(routineSub, "FA") > 0) Or (InStr(routineSub, "IP") > 0) Then
        If (Len(routineSub) - Len(Replace(routineSub, "_", "")) >= 2) Or (IsNumeric(Right(routineSub, 1))) Then
            If maxLevel = 1 Then
                GetMachiningLevel = 1
            ElseIf IsNumeric(Right(routineSub, 1)) Then
                Dim foundLevel As Integer
                foundLevel = CInt(Right(routineSub, 1))
                If foundLevel <= maxLevel Then
                    GetMachiningLevel = (foundLevel - 1)
                Else
                    'Return an error here, this should be impossible
                End If
            Else
                'Return an error, this naming convention is not allowed
            End If
        Else
            GetMachiningLevel = 0
        End If
    ElseIf InStr(routineSub, "FI") > 0 Then
        GetMachiningLevel = maxLevel
    Else
        'Return an error, we can't parse the abbreviation for this routineName
    End If

End Function

Private Sub SetFeatureVariables()

    On Error GoTo Err1

    featureHeaderInfo = DatabaseModule.GetFeatureHeaderInfo(jobNum:=jobNumUcase, routine:=rtCombo_TextField)
    
    'Should we filter or not filter observations shown based on Pass/Fail data
    'Having ShowAllObs pressed DOES NOT change the ObsFound value for the userform, that value is set in jbEditText
    If toggShowAllObs_Pressed Then
        featureMeasuredValues = DatabaseModule.GetAllFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=rtCombo_TextField, _
                                                delimFeatures:=JoinPivotFeatures(featureHeaderInfo))
        featureTraceabilityInfo = DatabaseModule.GetAllFeatureTraceabilityData(jobNum:=jobNumUcase, routine:=rtCombo_TextField)
    Else
        featureMeasuredValues = DatabaseModule.GetFeatureMeasuredValues(jobNum:=jobNumUcase, routine:=rtCombo_TextField, _
                                                delimFeatures:=JoinPivotFeatures(featureHeaderInfo), featureInfo:=featureHeaderInfo)
        featureTraceabilityInfo = DatabaseModule.GetFeatureTraceabilityData(jobNum:=jobNumUcase, routine:=rtCombo_TextField)
    End If
    
    Exit Sub
    
Err1:
    result = MsgBox("Could not set Job/Run information. Issue found at: " & vbCrLf & Err.description, vbCritical)
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
    machine = vbNullString
    cell = vbNullString
    partDesc = vbNullString
    
    'Keep routines for ComboBox
    Erase partRoutineList
    Erase runRoutineList


End Sub

Private Sub SetJobVariables(jobNum As String)
    On Error GoTo jbInfoErr
    Dim jobInfo() As Variant
    
    'TODO: somwhere here we need to check the size of the array (number of SWiss and/or CNC ops)
    'and possible need to check the operation numbers and maybe even op codes (need to add all this into the SQL query)
    
    jobInfo = DatabaseModule.GetJobInformation(JobID:=jobNum)
    
    
    'Add the components of the array to our variables
    
    partNum = jobInfo(2, 0)
    rev = jobInfo(3, 0)
    partDesc = jobInfo(5, 0)
    drawNum = jobInfo(6, 0)
    prodQty = jobInfo(7, 0)
    
    Exit Sub
    
    '    If Not sqlRecordSet.EOF Then
'        'Set values to pass to the Header Fields
'        If Not IsMissing(partNum) Then partNum = sqlRecordSet.Fields(2).Value
'        If Not IsMissing(rev) Then rev = sqlRecordSet.Fields(3).Value
'        If Not IsMissing(setupType) Then setupType = sqlRecordSet.Fields(4).Value
'
'        'This one is usually only called/set by the GetCustomerName()
'        If Not IsMissing(custName) Then custName = sqlRecordSet.Fields(5).Value
'
'        If Not IsMissing(machine) Then machine = sqlRecordSet.Fields(6).Value
'        If Not IsMissing(cell) Then cell = sqlRecordSet.Fields(7).Value
'        If Not IsMissing(partDescription) Then partDescription = sqlRecordSet.Fields(8).Value
'        If Not IsMissing(prodQty) Then prodQty = sqlRecordSet.Fields(9).Value
'        If Not IsMissing(drawNum) Then drawNum = sqlRecordSet.Fields(10).Value
'        GetJobInformation = True
'        Exit Function
'    End If
    

jbInfoErr:
    'If the recordSet is empty
    If Err.Number = vbObjectError + 2000 Then
        MsgBox ("Not A Valid Job Number")
    Else
    'Otherwise we encountered a different problem
        result = MsgBox(Err.description, vbExclamation)
    End If
    
    'Either way, reset the job number and invalidate the controls
    jobNumUcase = ""
    Err.Raise Number:=Err.Number, description:="SetJobVariables" & vbCrLf & Err.description


End Sub

Private Sub SetWorkbookInformation()
    Dim index As Integer
    index = GetRoutineIndex(rtCombo_TextField)
    
    On Error GoTo wbErr:
    Call ThisWorkbook.populateJobHeaders(jobNum:=jobNumUcase, routine:=rtCombo_TextField, customer:=customer, _
                                            machine:=runRoutineList(4, index), partNum:=partNum, rev:=rev, partDesc:=partDesc)
    Call ThisWorkbook.populateReport(featureInfo:=featureHeaderInfo, featureMeasurements:=featureMeasuredValues, _
                                        featureTraceability:=featureTraceabilityInfo)
    Exit Sub
wbErr:
    result = MsgBox("Could not set information to the workbook" & vbCrLf & "issue found at " & vbCrLf & Err.description, vbCritical)
    Err.Raise Number:=vbObjectError + 1200
    
End Sub

Public Function GetRoutineIndex(routineName As String) As Integer
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

