Attribute VB_Name = "ExcelHelpers"
Public Function GetAQL(customer As String, drawNum As String, prodQty As Integer) As String
    Dim partWb As Workbook
    Dim aqlWb As Workbook
    Dim aqlVal As String
    Dim reqQty As String
    Dim row As String
    Dim col As Integer

    prefixPath = "J:\Inspection Reports\" & customer & "\" & drawNum & "\" & "Current Revision\"
    
    'TODO: if the reuslt of dir is "", then that means that we didnt find it in Current Revision, we should switch to draft
    Filename = Dir(prefixPath & drawNum & "*.xlsm")
    
    
    If Filename = "" Then
        prefixPath = "J:\Inspection Reports\" & customer & "\" & drawNum & "\" & "Draft\"
        Filename = Dir(prefixPath & drawNum & "*.xlsm")
        
        If Filename = "" Then GoTo FileDirErr
        
    End If
    
    Application.ScreenUpdating = False
    Set partWb = Workbooks.Open(Filename:=prefixPath & Filename, UpdateLinks:=0, ReadOnly:=True)
        
    On Error GoTo WbReadErr
    
    aqlVal = partWb.Worksheets("ML Frequency Chart").Range("B7").Value
    If aqlVal = "" Then GoTo WbReadErr
    
    'TODO: need to add an AQL worksheet to this workbook so we can implement the rules and lookup of the a AQL value
    If aqlVal = "100%" Then
        'Set the value equal to the found prodQty and Exit sub
    End If
    
    'TODO: somehwere her we need to find the IR Tables workbook so we can switch on its value
    
    Set aqlWb = Workbooks.Open(Filename:="\\JADE76\IQS Documents\Current\IR Tables.xlsx", UpdateLinks:=0, ReadOnly:=True)
    
    
    Select Case prodQty
        Case 2 To 8
            row = "2"
        Case 9 To 15
            row = "3"
        Case 16 To 25
            row = "4"
        Case 26 To 50
            row = "5"
        Case 51 To 90
            row = "6"
        Case 91 To 150
            row = "7"
        Case 151 To 280
            row = "8"
        Case 281 To 500
            row = "9"
        Case 501 To 1200
            row = "10"
        Case 1201 To 3200
            row = "11"
        Case 3201 To 99999
            row = "12"
        Case Else
            'TODO: Error here, the value doesnt make sense
    End Select
    
    With aqlWb.Worksheets("AQL")
        col = .Range("B1:J1").Find(aqlVal).column
        reqQty = .Range(GetAddress(col) & row).Value
    End With
    
    'sometimes The qty required by an AQL is greater than the amount of parts we've made for some reason
    'Like for 10 parts with an AQL of 1.00
    If reqQty > prodQty Then
        GetAQL = prodQty
    Else
        GetAQL = reqQty
    End If
    
    GoTo 10
    
FileDirErr:
    Result = MsgBox("There was a problem opening an Inspection Report for " & vbCrLf & "Customer: " & customer & vbCrLf _
                & "Drawing: " & vbTab & drawNum & vbCrLf & vbCrLf & "The customer name may be incorrect or the " _
                    & "Inspection Report may be named incorrectly, contact a QE", vbExclamation)
                    
WbReadErr:
    Result = MsgBox("There was a problem when trying to read the AQL Level defined on the ML Frequency Chart Worksheet" & _
                    vbCrLf & "Please let a QE know to fill this value in", vbExclamation)
10
    partWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
End Function


Public Function GetAddress(column As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, column).Address(True, False), "$")
    GetAddress = vArr(0)

End Function

