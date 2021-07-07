VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VettingForm 
   Caption         =   "UserForm1"
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


Private Sub ObsFound_Click()

End Sub

Private Sub ResultFrame_Click()

End Sub

Private Sub Routine1_Click()

End Sub

Private Sub UserForm_Initialize()

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
            Dim routineType As String
            routineType = Split(RibbonCommands.partRoutineList(0, i), RibbonCommands.partNum & "_" & RibbonCommands.rev & "_")(1)
            'TODO: put error handling before the switching in case we get a routine of an unexpected Name
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
                Case "FA_SYLVAC", "FA_CMM" ' TODO: something for RAM as well here
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
                Case "IP_1XSHIFT"  'TODO: sometimes we seem to have a I XSHIFT? not 1. Needs to be corrected
                    .Caption = "98"  'TODO: we need to query for the total number of shifts worked here
                    .Visible = True
                Case "FI_DIM", "FI_VIS", "IP_LAST"
                    .Caption = "1"
                    .Visible = True
                Case Else ''TODO: This is the placeholder for the AQL, but there might end up being other routines we missed
                    .Caption = ExcelHelpers.GetAQL(customer:=RibbonCommands.customer, drawNum:=RibbonCommands.drawNum, _
                                            prodQty:=RibbonCommands.prodQty)
                    'We  should error handle differently here in case we cannot find the excel workbook
                    .Visible = True
            End Select
        End With
    Next i
    
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
                Debug.Print (RibbonCommands.runRoutineList(0, i) & ": " & RibbonCommands.runRoutineList(2, i))
                Me.RoutineFrame.Controls(j).ForeColor = RGB(0, 0, 0)
                Me.ObsFound.Controls(j).Caption = runRoutineList(2, i)
                Me.ObsFound.Controls(j).Visible = True
                GoTo NextControl
            End If
        Next j
        MsgBox ("Something went wrong here, couldn't find this routine")
        'TODO: if we found a routine that doesn't belong in the list of part number applicable routines then we should error out here
NextControl:
    Next i
    
    
    'Call Verifiction

End Sub

Private Sub UserForm_Activate()
'    MsgBox (Me.Controls("RoutineFrame").Routine1.Caption)
End Sub

Private Sub UserForm_Click()

End Sub


