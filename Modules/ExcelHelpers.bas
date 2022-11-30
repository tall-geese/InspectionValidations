Attribute VB_Name = "ExcelHelpers"
'*************************************************************************************************
'
'   ExcelHelpers
'       For Interacting with other Microsoft Office objects outside of ThisWorkbook
'       1. GetAQL  - from the inspection report workbook.
'       2. CreateEmail - to the cell lead / QC manager as applicable. Generate table of failed routines and why they failed
'*************************************************************************************************

Private valWB As Workbook
Private valLookupRange As Range
Private valRtRange As Range


Public Function GetAQL(customer As String, drawNum As String, ProdQty As Integer, Optional isShortRunEnabled As Boolean, _
                        Optional isChildOrParentJob As Boolean) As String()
    'Returns - String()
        '(0) -> Qty to Inspect
        '(1) -> AQL
            '**If its a Parent or Child Job**
        '(2) -> Qty to Inspect at FI_DIM Routines
        '(3) -> AQL applicable to FI_DIM Routines
            '**If its a Short Run Job**
        '(4) -> Short Run Size Cutoff
        '(5) -> Minimum Inspections
                        
                        
    Dim partWb As Workbook
    Dim aqlWB As Workbook
    Dim aqlVal As String
    Dim finalAQLVal As String
    Dim reqQty As String
    Dim finalReqQty As String
    Dim row As String
    Dim col As Integer
    
    Dim returnAQL() As String
    If isChildOrParentJob Then
        ReDim Preserve returnAQL(3)
    Else
        ReDim Preserve returnAQL(1)
    End If

    prefixPath = "J:\Inspection Reports\" & customer & "\" & drawNum & "\" & "Current Revision\"
    
    Filename = Dir(prefixPath & drawNum & "*.xlsm")
    
    If ProdQty = 0 Then GoTo ProdQtyErr
    
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
    
    If isChildOrParentJob Then
        finalAQLVal = partWb.Worksheets("ML Final Chart").Range("E7").Value
        If finalAQLVal = "" Then GoTo WbFinalReadErr
    End If
    
    If aqlVal = "100%" Or ProdQty = 1 Then
        returnAQL(0) = CStr(ProdQty)
        returnAQL(1) = "100%"
        
        If isChildOrParentJob Then  'Having a 100% AQL normally will override in the FinalAQL, this shouldn't normally happen though
            returnAQL(2) = CStr(ProdQty)
            returnAQL(3) = "100%"
        End If
        
        GetAQL = returnAQL
        Exit Function
    End If
        
    Set aqlWB = Workbooks.Open(Filename:=DataSources.IR_TABLES_WB, UpdateLinks:=0, ReadOnly:=True)
    
    
    Select Case ProdQty
        Case 2 To 4
            row = "2"
        Case 5 To 10
            row = "3"
        Case 11 To 15
            row = "4"
        Case 16 To 20
            row = "5"
        Case 21 To 25
            row = "6"
        Case 26 To 30
            row = "7"
        Case 31 To 35
            row = "8"
        Case 36 To 50
            row = "9"
        Case 51 To 90
            row = "10"
        Case 91 To 150
            row = "11"
        Case 151 To 280
            row = "12"
        Case 281 To 500
            row = "13"
        Case 501 To 1200
            row = "14"
        Case 1201 To 3200
            row = "15"
        Case 3201 To 10000
            row = "16"
        Case 10001 To 99999
            row = "17"
        Case Else
            GoTo ProdQtyErr
    End Select
    
    With aqlWB.Worksheets("AQL_SmallLot")
        col = Application.WorksheetFunction.Match(CDbl(aqlVal), .Range("A1:J1"), 0)
        reqQty = .Range(GetAddress(col) & row).Value
        
        If isChildOrParentJob Then
            col = Application.WorksheetFunction.Match(CDbl(finalAQLVal), .Range("A1:J1"), 0)
            finalReqQty = .Range(GetAddress(col) & row).Value
        End If
    End With
    
    'sometimes The qty required by an AQL is greater than the amount of parts we've made for some reason
    'Like for 10 parts with an AQL of 1.00
    If reqQty > ProdQty Then
        returnAQL(0) = CStr(ProdQty)
    Else
        returnAQL(0) = CStr(reqQty)
    End If
    
    returnAQL(1) = aqlVal
    
    If isChildOrParentJob Then
        If finalReqQty > ProdQty Then
            returnAQL(2) = CStr(ProdQty)
        Else
            returnAQL(2) = CStr(finalReqQty)
        End If
        
        returnAQL(3) = finalAQLVal
    End If
    
    
    
    If isShortRunEnabled Then
        ReDim Preserve returnAQL(UBound(returnAQL) + 2)
        On Error GoTo LowerBoundErr
        
        With partWb.Worksheets("ML Frequency Chart")
            returnAQL(UBound(returnAQL) - 1) = .Range("N14").Value
            returnAQL(UBound(returnAQL)) = .Range("R14").Value
        End With
            
    End If
    
    
    GetAQL = returnAQL
    GoTo 10
    
ProdQtyErr:
    result = MsgBox("There was a problem attempting to interpret this job's production quantity of " & ProdQty & vbCrLf & _
                     "Verify that this qty is correct in Epicor and contact a QE for assistance.", vbExclamation)
    GoTo 10
    
FileDirErr:
    result = MsgBox("There was a problem opening an Inspection Report for " & vbCrLf & "Customer: " & customer & vbCrLf _
                & "Drawing: " & vbTab & drawNum & vbCrLf & vbCrLf & "The customer name may be incorrect or the " _
                    & "Inspection Report may be named incorrectly, contact a QE", vbExclamation)
    GoTo 10
                    
WbReadErr:
    result = MsgBox("There was a problem when trying to read the AQL Level defined on the ML Frequency Chart Worksheet" & _
                    vbCrLf & "Please let a QE know to fill this value in" & vbCrLf & Err.Description, vbExclamation)
    GoTo 10
    
WbFinalReadErr:
    MsgBox "There was a problem when trying to read the AQL Level defined on the ML Final Chart Worksheet" & _
                    vbCrLf & "Please let a QE know to fill this value in" & vbCrLf & Err.Description, vbExclamation
    GoTo 10
    

LowerBoundErr:
    MsgBox "This DrawingNumber was set as LowerBound Frequency Enabled" & vbCrLf & "But Couldn't access the Cutoff amount of Inspections Due" _
                & vbCrLf & "Please Have a QE fix the IR", vbCritical
                    
10
    partWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
End Function


Public Sub CreateEmail(qcManager As Boolean, cellLead As Boolean, pmodManager As Boolean, cellLeadEmail As String, _
            jobNum As String, machine As String, failInfo() As Variant, shiftDetails() As Variant)
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
        If pmodManager Then
            .To = .To & ";" & DataSources.PMODMAN_TO
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
        
        'Possibly build the next table validating the number of the 1Xshfit inspections
        If Not Not shiftDetails Then
            HTMLContent = HTMLContent & "<br><h4>1XShift Production Records</h4>"
            HTMLContent = HTMLContent & MakeDetailsTable(shiftDetails)
        End If
        
        .HTMLBody = HTMLContent
    
    End With
    
    oMail.Display
    
End Sub

Private Function MakeDetailsTable(shiftDetails() As Variant) As String
    Dim outString As String
    Dim i As Integer, j As Integer
    For i = 0 To UBound(shiftDetails, 2)
        Dim opDetails() As Variant
        Dim table As String
        
        opDetails = shiftDetails(2, i)
        table = table & "<h5>" & shiftDetails(0, i) & " - " & shiftDetails(1, i) & "</h5>"
        
        
        table = table & "<table class=" & Chr(34) & "MsoTableGrid" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " cellspacing=" & Chr(34) & _
    "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse;border:none" & Chr(34) & ">"
        
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "JobNum" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Emp Name" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Date" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Shift" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Prod Qty" & "</td>"
    
        For j = 0 To UBound(opDetails, 2)
            table = table & "<tr>"
            For k = 0 To 4
               table = table & "<td>" & opDetails(k, j) & "</td>"
            Next k
            table = table & "</tr>"
        Next j
        
        table = table & "</table>"
        outString = outString & table
        table = vbNullString
    Next i
    
    MakeDetailsTable = outString
    

End Function

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


Public Function nRange(lower As Integer, upper As Integer) As Variant()
    Dim outArr() As Variant
    Dim i As Integer
    ReDim Preserve outArr(lower To upper)
    
    For i = lower To upper
        outArr(i) = CDbl(i)
    Next i
    nRange = outArr
    
End Function

Public Function updateForPivotSlice(inputArr() As Variant) As Variant()
    Dim outArr() As Variant
    ReDim Preserve outArr(1 To UBound(inputArr) + 1)
    Dim i As Integer
    
    For i = 1 To UBound(outArr)
        If i = 1 Then
            outArr(1) = CDbl(1)
        Else
            outArr(i) = inputArr(i - 1) + 1
        End If
    Next i
    
    updateForPivotSlice = outArr
End Function

Public Function fill_null(inputArr() As Variant) As Variant()
    'For whatever reason, VBA wont let us transpose or slice an array that has null values in it.
        'So will have to replace them with some filler values. However, it doenst really matter what
        'they are mostly for Arritbute features and we are going to be slicing those off anyways

    Dim i As Integer
    Dim j As Integer
    For i = 0 To UBound(inputArr)
        For j = 0 To UBound(inputArr, 2)
            If IsNull(inputArr(i, j)) Then
                inputArr(i, j) = 0
            End If
        Next j
    Next i
    fill_null = inputArr
End Function


Public Function InsertOpRow(partOperations() As Variant, jobNum As String, opNum As Variant, opCode As Variant) As Variant()
    'Take a non-empty array insert a new operation so tha the opNumber is in order.
    'VBA has no built-in sorting for 2-d arrays
    Dim i As Integer, outArr() As Variant, insertedNew As Boolean
    ReDim Preserve outArr(2, UBound(partOperations, 2) + 1)
    For i = 0 To UBound(outArr, 2)
        If i = UBound(outArr, 2) Then
            If Not insertedNew Then
                outArr(0, i) = jobNum
                outArr(1, i) = opNum
                outArr(2, i) = opCode
            Else
                GoTo prevPartOp
            End If
        ElseIf insertedNew Then  'If we've got our row in there, just add in the partOperations
prevPartOp:
            outArr(0, i) = partOperations(0, i - 1)
            outArr(1, i) = partOperations(1, i - 1)
            outArr(2, i) = partOperations(2, i - 1)
        ElseIf partOperations(1, i) = opNum Then  'Not allowed to be equal
            Err.Raise Number:=vbObjectError + 5000, Description:="OprSeq equal to an existing operation, cant insert this new Op" _
                & vbCrLf & vbCrLf & "partOp" & vbCrLf & partOperations(1, i) & vbTab & partOperations(2, i) _
                & vbCrLf & vbCrLf & "Insert-Op" & vbCrLf & opNum & vbTab & opCode
        ElseIf opNum < partOperations(1, i) And Not insertedNew Then   'Found the position to insert our input
            outArr(0, i) = jobNum
            outArr(1, i) = opNum
            outArr(2, i) = opCode
            insertedNew = True
        Else
            outArr(0, i) = partOperations(0, i)
            outArr(1, i) = partOperations(1, i)
            outArr(2, i) = partOperations(2, i)
        End If
    Next i
    
    InsertOpRow = outArr
    
End Function






'Handle Getting Alternate names for Inspection Methods to fit within our cell format

Public Sub OpenDataValWB()
    Set valWB = Workbooks.Open(Filename:=DataSources.DATA_VAL_WB, UpdateLinks:=False, ReadOnly:=True)
    With valWB.Worksheets("InspMethods")
        Set valLookupRange = .Range("A2:A" & .Range("A2").End(xlDown).row)
        Set valRtRange = .Range("B2:B" & .Range("A2").End(xlDown).row)
    
    End With
End Sub

Public Function GetShortHandMethod(inspMeth As Variant) As String
    GetShortHandMethod = Application.WorksheetFunction.XLookup(inspMeth, valLookupRange, valRtRange, CStr(inspMeth), 0)
End Function


Public Sub CloseDataValWB()
    valWB.Close SaveChanges:=False
    Set valWB = Nothing
End Sub



