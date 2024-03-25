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

'Main Job Information
Public jobNumUcase As String
Public job_json As Dictionary
Public part_json As Dictionary
Public created_Routines As Collection

Public run_feature_json As Collection
Public run_data_json As Collection
Public run_traceability_json As Collection

    'Optional Data for _FI_ routines that will be set in PAGE_Attr
Public fi_feature_json As Collection
Public fi_run_data_json As Collection
Public fi_run_traceability_json As Collection


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

    If Text = vbNullString Then
        Call cleanSheets
        Exit Sub
    End If
    
    On Error GoTo 10
    'Call the HTTP method and set the Variables
    Dim dhr_json As Dictionary
    Set dhr_json = HTTPconnections.ValidateDHR(job_num:=jobNumUcase)
    Set job_json = dhr_json("job_info")
    Set part_json = dhr_json("part_info")

    'Set our Ribbon Information to the first created Routine in our list
    'Look for other created routines as well and add them to a collection
    'This will be our selectable routiens in the drop-down
    Set created_Routines = New Collection
    For Each run_rt In job_json("Runs")
        If run_rt("Created") = True Then
            created_Routines.Add run_rt("Name")
            If created_Routines.Count = 1 Then
                rtCombo_TextField = run_rt("Name")
            End If
        End If
    Next run_rt
    
    If created_Routines.Count = 0 Then GoTo NoRunsCreated
    
    
    Dim runStatusCode As Integer
    runStatusCode = job_json("Runs")(1)("Status")
    lblStatus_Text = TranslateRunStatus(runStatusCode)
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
            'If Not job_json("IsChildJob") Then GoTo SetupTypeUndefined
    End Select
    
    
20
    If toggAutoForm_Pressed And job_json("Qty Complete") <> 0 Then VettingForm.Show
    
'TODO: this will all eventually move to RibbonCommands.SetWorkbookInformation()
'TODO: need to be able to handle us having the attribute FI routine being the first one, this should all be done in a seperate function here
'TODO: should handle not having any operations here
'TODO: determine machining level, use THAT for Operation index and machine name

    Call SetFeatureVariables
10
    Call SetWorkbookInformation
  
     cusRibbon.InvalidateControl "chkFull"
     cusRibbon.InvalidateControl "chkMini"
     cusRibbon.InvalidateControl "chkNone"
     cusRibbon.InvalidateControl "rtCombo"
     cusRibbon.InvalidateControl "jbEditText"
     cusRibbon.InvalidateControl "lblStatus"
   
    Exit Sub
    
SetupTypeUndefined:
    MsgBox "Cannot Determine Setup Type of " & job_json("Operations")(1)("Setup Type") & vbCrLf & "Have this changed to the appriopriate type in Job Entry", vbCritical
    Exit Sub
NoRunsCreated:
    MsgBox "No Runs have been created for this Job", vbInformation
    Exit Sub
    
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               RoutineName ComboBox
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub rtCombo_OnChange(ByRef control As Office.IRibbonControl, ByRef Text As Variant)

    'There doesn't seem to be a property to prevent the user from Hand-Typing into the ComboBox
    'So we have to make sure that the change is legitimate
    Dim validChange As Boolean, run As Dictionary, status As Integer
    validChange = False
    
    For Each run In job_json("Runs")
        If run("Name") = Text And run("Created") = True Then
            validChange = True
            status = run("Status")
            Exit For
        End If
    Next run
    
    
    'Erase the feature data but, not our Job Number or Job Routine List,
    'This means the user can still select from the drop-down and try again
    Call ClearFeatureVariables(preserveRoutines:=True)
    
    On Error GoTo 10
    If validChange = True Then
         'Set new active routine
        lblStatus_Text = TranslateRunStatus(status)
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
    If Not job_json Is Nothing Then
        Count = created_Routines.Count
    End If
End Sub

Public Sub rtCombo_OnGetItemLabel(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemLabel As Variant)
    ItemLabel = created_Routines(index + 1)
End Sub

Public Sub rtCombo_OnGetItemID(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemID As Variant)
    'Need to reference by ID? I guess not
End Sub

Public Sub rtCombo_OnGetText(ByRef control As Office.IRibbonControl, ByRef Text As Variant)
    'when we don't have any routines, use a placeholder string
    If Not job_json Is Nothing Then
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
'TODO: disbaled for now, will eventually call the 8001 port number version of the same routes

'Public Sub testDB_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
'    toggML7TestDB_Pressed = isPressed
'    Call DatabaseModule.Close_Connections 'If we had a connection already open, need to invalidate it so we can connect to the TestDB
'End Sub
'
'Public Sub testDB_OnGetEnabled(ByRef control As Office.IRibbonControl, ByRef ReturnedValue As Variant)
'    ReturnedValue = False
'End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Clean Sheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub cleanSheets_Pressed(ByRef control As Office.IRibbonControl)
    Call cleanSheets
End Sub

Public Sub cleanSheets()
    ThisWorkbook.Cleanup
    ClearFeatureVariables
    rtCombo_Enabled = False
    jobNumUcase = ""
    cusRibbon.InvalidateControl "chkFull"
    cusRibbon.InvalidateControl "chkMini"
    cusRibbon.InvalidateControl "chkNone"
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
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

    Dim run_routine As Dictionary
    For Each run_routine In job_json("Runs")
        If run_routine("Created") = False Then GoTo skip_rt
        rtCombo_TextField = run_routine("Name")
        lblStatus_Text = TranslateRunStatus(run_routine("Status"))
        
        On Error GoTo 10
        Call ClearFeatureVariables(preserveRoutines:=True, preserveRoutineName:=True)
        Call SetFeatureVariables
        Call SetWorkbookInformation
        Call ThisWorkbook.PrintResults
skip_rt:
    Next run_routine

10
    'The Ribbon information will be updated to the last routine that was printed / activated
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
    
    
End Sub


Public Function GetMachiningOpInfo(routineName As Variant) As Variant()
    'Using the RunRoutine Name, find the it in our list of required PartRoutines and return the set Level
    
    Dim out_info(3) As Variant
'    out_info(0) = SWISS 'opcode
'    out_info(1) = 10 'opseq
'    out_info(2) = "QC" 'cell
'    out_info(3) = "QC" 'machine
                
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
        out_info(3) = "QC" 'machin
        GetMachiningOpInfo = out_info
        Exit Function

    End If
    
    If part_json("part_routines").Count = 0 Then Err.Raise Number:=vbObjectError + 2500
    For Each part_rt In part_json("part_routines")
        If routineName = part_rt("Name") Then
            
            out_info(0) = job_json("Operations")(part_rt("Level") + 1)("OpCode")
            out_info(1) = job_json("Operations")(part_rt("Level") + 1)("OprSeq")
            out_info(2) = job_json("Operations")(part_rt("Level") + 1)("Cell")
            out_info(3) = job_json("Operations")(part_rt("Level") + 1)("Machine")
        
            GetMachiningOpInfo = out_info
            
            Exit Function
        End If
    Next part_rt

RoutineNotFound:
    Err.Raise Number:=vbObjectError + 4000, Description:="Created Run of " & routineName & vbCrLf & "But couldn't find a matching required routine in MeasurLink." & _
        vbCrLf & "Can't determine what Operation the routine belongs to"
    
RoutineCountErr:
    Err.Raise Number:=vbObjectError + 2500, Description:="Couldn't figure out, what machining operation " & routineName & _
        vbCrLf & "should belong to. There were no required Part Routines" & vbCrLf & vbCrLf & Err.Description
End Function



Private Sub ClearFeatureVariables(Optional preserveRoutines As Boolean, Optional preserveRoutineName As Boolean = False)
    'Clear the Inspection Data for a Run, Optionally clear the Job Info itself if we are moving onto another job

    If Not IsMissing(preserveRoutines) Then
        If preserveRoutines = False Then
            Set job_json = Nothing
            Set part_json = Nothing
            Set created_Routines = Nothing
        End If
    End If

    Set run_feature_json = Nothing
    Set run_data_json = Nothing
    Set run_traceability_json = Nothing
    Set fi_feature_json = Nothing
    Set fi_run_data = Nothing
    Set fi_run_traceability = Nothing
    
    If Not preserveRoutineName Then
        rtCombo_TextField = ""
    End If
    
End Sub



Private Sub SetFeatureVariables()
    'Called by jbEditText_OnChange()
        ' and rtCombo_OnChange()

    On Error GoTo Err1
    
    If toggShowAllObs_Pressed And Not rtCombo_TextField Like "*_FI_*" Then
        'TODO:  dont have the routes for these yet
    
    
    ElseIf rtCombo_TextField Like "*_FI_*" Then  'FI routines, seperate out the Variable and the Attribute
        'Variable
        Set result = HTTPconnections.GetPassedInspData(jobNumUcase, rtCombo_TextField, feature_type_only:=DataSources.TYPE_VARIABLE)
        Set run_feature_json = result("feature_info")
        Set run_data_json = result("insp_data")
        Set run_traceability_json = result("traceability")
            
        'Attribute
        Set result = HTTPconnections.GetPassedInspData(jobNumUcase, rtCombo_TextField, feature_type_only:=DataSources.TYPE_ATTRIBUTE)
        Set fi_feature_json = result("feature_info")
        Set fi_run_data_json = result("insp_data")
        Set fi_run_traceability_json = result("traceability")
        
        
            'If we returned an empty list, just unset this
        If run_data_json.Count = 0 Then Set run_data_json = Nothing
        If fi_run_data_json.Count = 0 Then Set fi_run_data_json = Nothing

    Else  'The Bread and Butter
        Set result = HTTPconnections.GetPassedInspData(jobNumUcase, rtCombo_TextField)
        Set run_feature_json = result("feature_info")
        Set run_data_json = result("insp_data")
        Set run_traceability_json = result("traceability")
    
    End If
    
    Exit Sub
    
Err1:
    result = MsgBox("Could not set Job/Run information. Issue found at: " & vbCrLf & Err.Description, vbCritical)
    Err.Raise Number:=vbObjectError + 1000

End Sub


Private Sub SetWorkbookInformation()
    'Should be called After SetFeatureVariables
    'Populate difference aspects of the Report given Routine conditions
    
    Dim machine As String
    Dim op_info() As Variant
    
    'Find Out what Machine the Routine had run on
    If job_json Is Nothing Then Exit Sub
    If job_json("Operations").Count = 0 Then
        machine = "N/A"
    Else
        op_info = GetMachiningOpInfo(rtCombo_TextField)
        machine = op_info(3)
    End If
    
    
    'We call Cleanup at the top of Populate Job Headers
    ThisWorkbook.populateJobHeaders jobNum:=jobNumUcase, routine:=rtCombo_TextField, _
        customer:=job_json("Customer"), machine:=machine, partNum:=job_json("PartNum"), _
        rev:=job_json("RevisionNum"), partDesc:=job_json("PartDescription")
    
    
    On Error GoTo wbErr
    ExcelHelpers.OpenDataValWB
    Call ThisWorkbook.populateReport(header_info:=run_feature_json, insp_data:=run_data_json, traceability:=run_traceability_json)
    Call ThisWorkbook.populateAttrSheet(job_json:=job_json, feature_json:=fi_feature_json, insp_data:=fi_run_data_json, traceability_json:=fi_run_traceability_json)
    
    
            
    ExcelHelpers.CloseDataValWB
    Exit Sub
wbErr:
    ExcelHelpers.CloseDataValWB
    result = MsgBox("Error Occurred at Sub: RibbonCommands.SetWorkbookInformation", vbCritical)
    Err.Raise Number:=vbObjectError + 1200
    
End Sub


Public Function TranslateRunStatus(runStatusCode As Integer) As String
    Select Case runStatusCode
        Case DataSources.RUN_STATUS_NOT_CREATED  '0
            TranslateRunStatus = "Never Created"
        Case DataSources.RUN_STATUS_NEW  '1
            TranslateRunStatus = "New"
        Case DataSources.RUN_STATUS_SUSPENDED  '2
            TranslateRunStatus = "Suspended"
        Case DataSources.RUN_STATUS_ACTIVE  '3
            TranslateRunStatus = "Active"
        Case DataSources.RUN_STATUS_CLOSED  '4
            TranslateRunStatus = "Closed"
        Case DataSources.RUN_STATUS_ARCHIVED  '12
            TranslateRunStatus = "Archived"
        Case DataSources.RUN_STATUS_SIGNED  '260
            TranslateRunStatus = "Signed"
        Case Else
            TranslateRunStatus = "???"
    End Select
End Function

