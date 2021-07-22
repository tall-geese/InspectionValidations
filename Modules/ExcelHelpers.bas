Attribute VB_Name = "ExcelHelpers"
'*************************************************************************************************
'
'   ExcelHelpers
'       For Interacting with other Microsoft Office objects outside of ThisWorkbook
'       1. GetAQL  - from the inspection report workbook.
'       2. CreateEmail - to the cell lead / QC manager as applicable. Generate table of failed routines and why they failed
'*************************************************************************************************


Public Function GetAQL(customer As String, drawNum As String, prodQty As Integer) As String
    Dim partWb As Workbook
    Dim aqlWb As Workbook
    Dim aqlVal As String
    Dim reqQty As String
    Dim row As String
    Dim col As Integer

    prefixPath = "J:\Inspection Reports\" & customer & "\" & drawNum & "\" & "Current Revision\"
    
    Filename = Dir(prefixPath & drawNum & "*.xlsm")
    
    
    If Filename = "" Then
        'If there isn't an xl file in the directory, it may be in the draft
        prefixPath = "J:\Inspection Reports\" & customer & "\" & drawNum & "\" & "Draft\"
        Filename = Dir(prefixPath & drawNum & "*.xlsm")
        
        'If still nothing then something wrong with the inspection report
        If Filename = "" Then GoTo FileDirErr
        
    End If
    
    Application.ScreenUpdating = False
    Set partWb = Workbooks.Open(Filename:=prefixPath & Filename, UpdateLinks:=0, ReadOnly:=True)
        
    On Error GoTo WbReadErr
    
    aqlVal = partWb.Worksheets("ML Frequency Chart").Range("B7").Value
    If aqlVal = "" Then GoTo WbReadErr
    
    If aqlVal = "100%" Then
        GetAQL = prodQty
        Exit Function
    End If
    
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
            GoTo ProdQtyErr
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
    
ProdQtyErr:
    result = MsgBox("There was a problem attempting to interpret this job's production quantity of " & prodQty & vbCrLf & _
                     "Verify that this qty is correct in Epicor and contact a QE for assistance.", vbExclamation)
    GoTo 10
    
FileDirErr:
    result = MsgBox("There was a problem opening an Inspection Report for " & vbCrLf & "Customer: " & customer & vbCrLf _
                & "Drawing: " & vbTab & drawNum & vbCrLf & vbCrLf & "The customer name may be incorrect or the " _
                    & "Inspection Report may be named incorrectly, contact a QE", vbExclamation)
    GoTo 10
                    
WbReadErr:
    result = MsgBox("There was a problem when trying to read the AQL Level defined on the ML Frequency Chart Worksheet" & _
                    vbCrLf & "Please let a QE know to fill this value in", vbExclamation)
10
    partWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
End Function


Public Sub CreateEmail(qcManager As Boolean, cellLead As Boolean, cellLeadEmail As String, jobNum As String, machine As String, failInfo() As Variant)
    Dim oApp As Outlook.Application
    Dim myMail As Outlook.MailItem
    Dim HTMLContent As String
    
    Set oApp = New Outlook.Application
    Set oMail = oApp.CreateItem(olMailItem)
    
    With oMail
        .To = DataSources.PQCI_TO
        If cellLead Then
            .To = .To & ";" & cellLeadEmail
        End If
        If qcManager Then
            .To = .To & ";" & DataSources.QCMAN_TO
        End If
        
        .Subject = Replace(DataSources.EMAIL_SUBJECT, "{Job}", jobNum)
        .Subject = Replace(.Subject, "{Machine}", machine)
        
        HTMLContent = DataSources.EMAIL_BODY_HEADER
        
        HTMLContent = HTMLContent & "<table class=" & Chr(34) & "MsoTableGrid" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " cellspacing=" & Chr(34) & _
    "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse;border:none" & Chr(34) & ">"
        
        HTMLContent = HTMLContent & "<td width=" & Chr(34) & "290" & Chr(34) & ">" & "Routine Name" & "</td>"
        HTMLContent = HTMLContent & "<td width=" & Chr(34) & "100" & Chr(34) & ">" & "ObsReq" & "</td>"
        HTMLContent = HTMLContent & "<td width=" & Chr(34) & "100" & Chr(34) & ">" & "ObsFound" & "</td>"
    
        For i = 0 To UBound(failInfo, 2)
            HTMLContent = HTMLContent & "<tr>"
            For j = 0 To 2
               HTMLContent = HTMLContent & "<td>" & failInfo(j, i) & "</td>"
            Next j
            HTMLContent = HTMLContent & "</tr>"
        Next i
        
        HTMLContent = HTMLContent & "</table>"
        
        HTMLContent = HTMLContent & DataSources.EMAIL_BODY_FOOTER
        
        .HTMLBody = HTMLContent
    
    End With
    
    oMail.Display
    
End Sub

Function xTab(num As Integer) As String
    For i = 1 To num
        xTab = xTab & vbTab
    Next i
End Function


Public Function GetAddress(column As Integer) As String
    Dim vArr
    vArr = Split(cells(1, column).Address(True, False), "$")
    GetAddress = vArr(0)

End Function
