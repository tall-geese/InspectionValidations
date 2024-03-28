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




Public Sub CreateEmail(qcManager As Boolean, cellLead As Boolean, pmodManager As Boolean, cellLeadEmail As String, _
            jobNum As String, machine As String, failInfo() As Variant, shiftDetails() As Variant, shiftTraceability() As Variant)
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
        
        .CC = DataSources.EMAIL_CC
        .Subject = Replace(DataSources.EMAIL_SUBJECT, "{Job}", jobNum)
        .Subject = Replace(.Subject, "{Machine}", machine)
        
        HTMLContent = DataSources.EMAIL_BODY_HEADER
        
        HTMLContent = HTMLContent & "<table class=" & Chr(34) & "MsoTableGrid" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " cellspacing=" & Chr(34) & _
    "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse;border:none" & Chr(34) & ">"
        
        HTMLContent = HTMLContent & "<td width=" & Chr(34) & "290" & Chr(34) & ">" & "Routine Name" & "</td>"
        HTMLContent = HTMLContent & "<td width=" & Chr(34) & "100" & Chr(34) & ">" & "ObsReq" & "</td>"
        HTMLContent = HTMLContent & "<td width=" & Chr(34) & "100" & Chr(34) & ">" & "ObsFound" & "</td>"
    
        'failed Routine Name, its Obs_Req and its Obs_Found
        If Not Not failInfo Then
            
            For i = 0 To UBound(failInfo, 2)
                HTMLContent = HTMLContent & "<tr>"
                For j = 0 To 2
                   HTMLContent = HTMLContent & "<td>" & failInfo(j, i) & "</td>"
                Next j
                HTMLContent = HTMLContent & "</tr>"
            Next i
        End If
        
        HTMLContent = HTMLContent & "</table>"
        
        HTMLContent = HTMLContent & DataSources.EMAIL_BODY_FOOTER
        
        'Possibly build the next table validating the number of the 1Xshfit inspections
        If Not Not shiftDetails And Not Not shiftTraceability Then
            HTMLContent = HTMLContent & "<br><h4>1XShift Production Records</h4>"
            HTMLContent = HTMLContent & MakeProductionDetailsTable(shiftDetails)
            
            HTMLContent = HTMLContent & "<br><h4>1XShift Inspection Records</h4>"
            HTMLContent = HTMLContent & MakeInspectionDetailsTable(shiftTraceability)
        End If
        
        .HTMLBody = HTMLContent
    
    End With
    
    oMail.Display
    
End Sub

    'TODO: this should be modular to do the same operation for the MeasurLink inspections of the 1XSHIFT routine, dont have time now

Private Function MakeProductionDetailsTable(shiftDetails() As Variant) As String
    Dim outString As String
    Dim i As Integer, j As Integer
    For i = 0 To UBound(shiftDetails, 2)
        Dim opDetails As Collection
        Dim table As String
        
        Set opDetails = shiftDetails(2, i)
        table = table & "<h5>" & shiftDetails(0, i) & " - " & shiftDetails(1, i) & "</h5>"
        
        
        table = table & "<table class=" & Chr(34) & "MsoTableGrid" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " cellspacing=" & Chr(34) & _
    "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse;border:none" & Chr(34) & ">"
        
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "JobNum" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Emp Name" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Emp ID" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Date" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Shift" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Prod Qty" & "</td>"
    
        'Get each key from the Collection
        Dim keys() As Variant
        keys = Array("JobNum", "Name", "EmpID", "PayrollDate", "Shift", "LaborQty")
        For j = 1 To opDetails.Count
            
            table = table & "<tr>"
            For k = 0 To 5
               table = table & "<td>" & opDetails(j)(keys(k)) & "</td>"
            Next k
            table = table & "</tr>"
        Next j
        
        table = table & "</table>"
        outString = outString & table
        table = vbNullString
    Next i
    
    MakeProductionDetailsTable = outString
    

End Function

Private Function MakeInspectionDetailsTable(shiftTraceability() As Variant) As String
    Dim outString As String
    Dim i As Integer, j As Integer
    For i = 0 To UBound(shiftTraceability, 2)
        Dim opDetails As Collection
        Dim table As String
        
        Set opDetails = shiftTraceability(2, i)
        table = table & "<h5>" & shiftTraceability(0, i) & " - " & shiftTraceability(1, i) & "</h5>"   'JobNum - Routine
        
        
        table = table & "<table class=" & Chr(34) & "MsoTableGrid" & Chr(34) & " border=" & Chr(34) & "1" & Chr(34) & " cellspacing=" & Chr(34) & _
    "0" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " style=" & Chr(34) & "border-collapse:collapse;border:none" & Chr(34) & ">"
        
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "ObsTimeStamp" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "EmpID" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Obs#" & "</td>"
        table = table & "<td width=" & Chr(34) & "150" & Chr(34) & ">" & "Pass/Fail" & "</td>"
    
        Dim keys() As Variant
        keys = Array("TimeStamp", "EmployeeID", "ObsNo", "Result")
        For j = 1 To opDetails.Count
            table = table & "<tr>"
            For k = 0 To 3
               table = table & "<td>" & opDetails(j)(keys(k)) & "</td>"
            Next k
            table = table & "</tr>"
        Next j
        
        table = table & "</table>"
        outString = outString & table
        table = vbNullString
    Next i
    
    MakeInspectionDetailsTable = outString
    

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



