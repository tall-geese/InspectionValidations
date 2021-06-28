Attribute VB_Name = "RibbonCommands"

'*************************************************************************************************
'
'   RibbonCommands
'       Event logic for the Custom Ribbon Controls
'       1. The JobID Field and our editText Form should be updated to be the same when chanes are applied
'       2. The RoutineSelection and our ComboBox should be updated to be the same when changes are applied
'       3. we should ask the DataBase Module to perform our check on whether a jobNumber actually exists and is valid
'*************************************************************************************************


Dim cusRibbon As IRibbonUI

Dim toggAutoForm_Pressed As Boolean

Dim editTextUcase As String

Dim partRoutineList() As Variant

Dim lblStatus_Text As String

Dim runRoutineList() As Variant
Dim rtCombo_TextField As String
Dim rtCombo_Enabled As Boolean

Dim chkFull_Pressed As Boolean
Dim chkMini_Pressed As Boolean
Dim chkNone_Pressed As Boolean




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               UI Ribbon
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Ribbon_OnLoad(uiRibbon As IRibbonUI)
    Set cusRibbon = uiRibbon
    cusRibbon.ActivateTab "mlTab"
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               LoadForm Button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Callback(ByRef control As Office.IRibbonControl)
    sampleText = "try This out"
    cusRibbon.InvalidateControl "jbEditText"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Auto Load Form Toggle Button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub toggAutoForm_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    toggAutoForm_Pressed = isPressed
End Sub

Public Sub toggAutoForm_OnGetPressed(ByRef control As Office.IRibbonControl, ByRef ReturnedValue As Variant)
    toggAutoForm_Pressed = True
    ReturnedValue = toggAutoForm_Pressed
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Job Number EditTextField
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub jbEditText_onGetText(ByRef control As IRibbonControl, ByRef Text)
    Text = editTextUcase
    'Ask the workbook to Add the Information to Header Fields
    ThisWorkbook.populateHeaders jobNum:=editTextUcase, routine:=rtCombo_TextField

End Sub

Public Sub jbEditText_OnChange(ByRef control As Office.IRibbonControl, ByRef Text As String)
    'Reset the Variables
    editTextUcase = UCase(Text)
    chkFull_Pressed = False
    chkMini_Pressed = False
    chkNone_Pressed = False
    lblStatus_Text = vbNullString
    
    rtCombo_Enabled = False
    rtCombo_TextField = vbNullString
    Erase partRoutineList
    Erase runRoutineList
    
    If Text = vbNullString Then GoTo 10
    
    Dim PartNum As String
    Dim rev As String
    Dim setupType As String
    
    If DatabaseModule.VerifyJobExists(Text, PartNum, rev, setupType) Then
    
        On Error GoTo ML_NotApplicable:
        'TODO create two respective routine retrievals for both run and Part
        runRoutineList = DatabaseModule.GetRunRoutineList(editTextUcase).GetRows()
        partRoutineList = DatabaseModule.GetPartRoutineList(PartNum, rev).GetRows()
        
        'TODO: reset the error handling here, test with a 1/0 math
        
        rtCombo_TextField = runRoutineList(0, 0)
        lblStatus_Text = runRoutineList(1, 0)
        rtCombo_Enabled = True

        Select Case setupType
            Case "Full"
                chkFull_Pressed = True
            Case "Mini"
                chkMini_Pressed = True
            Case "None"
                chkNone_Pressed = True
            Case Else
                'Todo: Handle we don't know what the setupType is.
        End Select
        

    Else
        MsgBox ("Not A Valid Job Number")
        editTextUcase = ""
    End If
10
    'Standard updates that are always applicable
    cusRibbon.InvalidateControl "chkFull"
    cusRibbon.InvalidateControl "chkMini"
    cusRibbon.InvalidateControl "chkNone"
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"
    cusRibbon.InvalidateControl "lblStatus"
    


    Exit Sub

ML_NotApplicable:
    MsgBox Prompt:="Not Routines Found for this Job or Part Number " & vbCrLf & "If this is a MeasurLink Job, bring to QE's attention ", Buttons:=vbExclamation
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
    
    If Not Not runRoutineList Then
        For i = 0 To UBound(runRoutineList, 2)
            If Text = runRoutineList(0, i) Then
                validChange = True
                lblStatus_Text = runRoutineList(1, i)
                
                'We have to update the routine header here becuase selecting from the list won't call the OnGetText() event
                ThisWorkbook.populateHeaders jobNum:=editTextUcase, routine:=Text
            End If
        Next i
    End If
    
    If validChange = False Then
        rtCombo_TextField = ""
        lblStatus_Text = ""
        cusRibbon.InvalidateControl "rtCombo"
    End If
    
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
    'TODO
    'Debug.Print ("hit the get item ID")
End Sub

Public Sub rtCombo_OnGetText(ByRef control As Office.IRibbonControl, ByRef Text As Variant)
    'Believe it or not, this is the proper way to check if a Variant Array has been initialized
    'TODO: do we even need to check if this array is initialized? Maybe we can just check rtCombo_TextField here
    If Not Not runRoutineList Then
        Text = rtCombo_TextField
    Else
        Text = "[SELECT ROUTINE]"
    End If
        'Ask the workbook to Add the Information to Header Fields
    ThisWorkbook.populateHeaders jobNum:=editTextUcase, routine:=rtCombo_TextField

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
'              Show All Observations Toggle Buttom
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub allObs_Toggle(ByRef control As Office.IRibbonControl, ByRef isPressed As Boolean)
    'TODO:
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               JobType Check Boxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub chkFull_OnAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    'TODO:
End Sub

Public Sub chkFull_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
End Sub

Public Sub chkFull_OnGetPressed(ByRef control As IRibbonControl, ByRef pressed As Variant)
    pressed = chkFull_Pressed
End Sub

Public Sub chkMini_OnAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    'TODO:
End Sub
Public Sub chkMini_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
End Sub
Public Sub chkMini_OnGetPressed(ByRef control As IRibbonControl, ByRef pressed As Variant)
    pressed = chkMini_Pressed
End Sub

Public Sub chkNone_OnAction(ByRef control As IRibbonControl, ByRef pressed As Boolean)
    'TODO:
End Sub
Public Sub chkNone_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False
End Sub
Public Sub chkNone_OnGetPressed(ByRef control As IRibbonControl, ByRef pressed As Variant)
    pressed = chkNone_Pressed
End Sub





