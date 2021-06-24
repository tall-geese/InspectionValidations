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


Dim editTextUcase As String

Dim routineList() As Variant
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
'               Job Number EditTextField
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub jbEditText_onGetText(ByRef control As IRibbonControl, ByRef Text)
    Text = editTextUcase

End Sub

Public Sub jbEditText_OnChange(ByRef control As Office.IRibbonControl, ByRef Text As String)
    'Reset the Variables
    editTextUcase = UCase(Text)
    chkFull_Pressed = False
    chkMini_Pressed = False
    chkNone_Pressed = False
    
    rtCombo_Enabled = False
    rtCombo_TextField = vbNullString
    Erase routineList
    
    If Text = vbNullString Then GoTo 10
    
    Dim PartNum As String
    Dim rev As String
    Dim setupType As String
    
    If DatabaseModule.VerifyJobExists(Text, PartNum, rev, setupType) Then
    
        On Error GoTo ML_NotApplicable:
        routineList = DatabaseModule.GetRoutineList(PartNum, rev).GetRows()
        
        'TODO: reset the error handling here, test with a 1/0 math
        
        rtCombo_TextField = routineList(0, 0)
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
    End If
10
    'Standard updates that are always applicable
    cusRibbon.InvalidateControl "chkFull"
    cusRibbon.InvalidateControl "chkMini"
    cusRibbon.InvalidateControl "chkNone"
    cusRibbon.InvalidateControl "rtCombo"
    cusRibbon.InvalidateControl "jbEditText"

    Exit Sub

ML_NotApplicable:
    MsgBox Prompt:="Not Routines Found for this Job Number " & vbCrLf & "If this is a MeasurLink Job, bring to QE's attention ", Buttons:=vbExclamation
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
    If Not Not routineList Then
        For i = 0 To UBound(routineList, 2)
            If Text = routineList(0, i) Then validChange = True
            Debug.Print (routineList(0, i))
        Next i
    End If
    
    If validChange = False Then
        rtCombo_TextField = ""
        cusRibbon.InvalidateControl "rtCombo"
    
    End If
    
End Sub

Public Sub rtCombo_OnGetEnabled(ByRef control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = rtCombo_Enabled
End Sub

Public Sub rtCombo_OnGetItemCount(ByRef control As Office.IRibbonControl, ByRef Count As Variant)
    If Not IsEmpty(routineList) Then
        Count = UBound(routineList, 2) + 1
    End If
End Sub

Public Sub rtCombo_OnGetItemLabel(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemLabel As Variant)
    ItemLabel = routineList(0, index)
End Sub

Public Sub rtCombo_OnGetItemID(ByRef control As Office.IRibbonControl, ByRef index As Integer, ByRef ItemID As Variant)
    'TODO
    'Debug.Print ("hit the get item ID")
End Sub

Public Sub rtCombo_OnGetText(ByRef control As Office.IRibbonControl, ByRef Text As Variant)
    'Believe it or not, this is the proper way to check if a Variant Array has been initialized
    If Not Not routineList Then
        Text = rtCombo_TextField
    Else
        Text = "[SELECT ROUTINE]"
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




