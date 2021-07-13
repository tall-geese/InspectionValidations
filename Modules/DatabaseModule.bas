Attribute VB_Name = "DatabaseModule"

'*************************************************************************************************
'
'   DataBase Module
'
'*************************************************************************************************


Dim E10DataBaseConnection As ADODB.Connection
Dim ML7DataBaseConnection As ADODB.Connection
Dim KioskDataBaseConnection As ADODB.Connection
Dim sqlCommand As ADODB.Command
Dim sqlRecordSet As ADODB.Recordset
Dim fso As New FileSystemObject
Dim query As String
Dim params() As Variant

Private Enum Connections
    E10 = 0
    ML7 = 1
    Kiosk = 2
End Enum


Sub Init_Connections()

    On Error GoTo Err_Conn
    
    If ML7DataBaseConnection Is Nothing Then
        
        Set ML7DataBaseConnection = New ADODB.Connection
        If RibbonCommands.toggML7TestDB_Pressed Then
            ML7DataBaseConnection.ConnectionString = DataSources.ML7TEST_CONN_STRING
        Else
            ML7DataBaseConnection.ConnectionString = DataSources.ML7_CONN_STRING
        End If
        
        ML7DataBaseConnection.Open
    End If
      
    If E10DataBaseConnection Is Nothing Then
        Set E10DataBaseConnection = New ADODB.Connection
        E10DataBaseConnection.ConnectionString = DataSources.E10_CONN_STRING
        E10DataBaseConnection.Open
    End If
    
    If KioskDataBaseConnection Is Nothing Then
        Set KioskDataBaseConnection = New ADODB.Connection
        KioskDataBaseConnection.ConnectionString = DataSources.KIOSK_CONN_STRING
        KioskDataBaseConnection.Open
    End If
    
    
       
        
    Exit Sub
    
Err_Conn:
    Err.Raise Number:=Err.Number, description:="There was an error connecting with the Epicor and/or MeasurLink Database " _
        & "you may not be connected to the Network or you may not have permission from the Administrator to read from the MeasurLink DataBase"

End Sub

    'Calling function should pass us a query and array of parameters to set
    'In this sub we should execute the query and set the topLevel recordset that the calling function can use GetRows() on
Private Sub ExecQuery(query As String, params() As Variant, conn_enum As Connections)

    Call Init_Connections
    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = GetConnection(conn_enum)
        .CommandType = adCmdText
        .CommandText = query
    
        'Params structure
        'params(0) = "jh.JoNum,'NV1452'"
        For i = 0 To UBound(params)
            Dim queryParam As ADODB.Parameter
            Set queryParam = .CreateParameter(Name:=Split(params(i), ",")(0), Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=Split(params(i), ",")(1))
            .Parameters.Append queryParam
        Next i
    
    End With
    
    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.Open sqlCommand

    If sqlRecordSet.EOF Then
        'Raise an error here that we returned no rows? That would mean someone created a routine but didnt do any inpsections
        'See if we can cascade the error upwards where the routine name could be accessed.
    End If

End Sub



Private Function GetConnection(conn_enum As Connections) As ADODB.Connection
    Select Case conn_enum
        Case 0
            Set GetConnection = E10DataBaseConnection
        Case 1
            Set GetConnection = ML7DataBaseConnection
        Case 2
            Set GetConnection = KioskDataBaseConnection
        Case Else
    End Select
End Function


Sub Close_Connections()
    If Not (ML7DataBaseConnection Is Nothing) Then ML7DataBaseConnection.Close
    If Not (E10DataBaseConnection Is Nothing) Then E10DataBaseConnection.Close
    If Not (KioskDataBaseConnection Is Nothing) Then KioskDataBaseConnection.Close
End Sub












''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Epicor
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetJobInformation(JobID As String, Optional ByRef partNum As Variant, Optional ByRef rev As Variant, _
                                Optional ByRef setupType As Variant, Optional ByRef custName As Variant, _
                                Optional ByRef machine As Variant, Optional ByRef cell As Variant, _
                                Optional ByRef partDescription As Variant, Optional prodQty As Variant, _
                                Optional ByRef drawNum As Variant) As Boolean
    
    Set fso = New FileSystemObject
    params = Array("jh.JobNum," & JobID)
    query = fso.OpenTextFile(DataSources.QUERIES_PATH & "EpicorJobInfo.sql").ReadAll
    
    
    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.E10)
    
    
    If Not sqlRecordSet.EOF Then
        'Set values to pass to the Header Fields
        If Not IsMissing(partNum) Then partNum = sqlRecordSet.Fields(2).Value
        If Not IsMissing(rev) Then rev = sqlRecordSet.Fields(3).Value
        If Not IsMissing(setupType) Then setupType = sqlRecordSet.Fields(4).Value
        
        'This one is usually only called/set by the GetCustomerName()
        If Not IsMissing(custName) Then custName = sqlRecordSet.Fields(5).Value
        
        If Not IsMissing(machine) Then machine = sqlRecordSet.Fields(6).Value
        If Not IsMissing(cell) Then cell = sqlRecordSet.Fields(7).Value
        If Not IsMissing(partDescription) Then partDescription = sqlRecordSet.Fields(8).Value
        If Not IsMissing(prodQty) Then prodQty = sqlRecordSet.Fields(9).Value
        If Not IsMissing(drawNum) Then drawNum = sqlRecordSet.Fields(10).Value
        GetJobInformation = True
        Exit Function
    End If
    
    'If now rows are returned, the job doesn't exist
    GetJobInformation = False
End Function


Function Get1XSHIFTInsps(JobID As String) As String

    Set fso = New FileSystemObject
    params = Array("jo.JobNum," & JobID)
    query = fso.OpenTextFile(DataSources.QUERIES_PATH & "1XSHIFT.sql").ReadAll

    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.E10)
    
    If Not sqlRecordSet.EOF Then
        Get1XSHIFTInsps = sqlRecordSet.Fields(1).Value
        Exit Function
    End If
    
    'TODO: Error here, we will need to the customer name to determine the workbook directory for AQL
    Get1XSHIFTInsps = "None"
End Function










''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               MeasurLink
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetFeatureHeaderInfo(jobNum As String, routine As String) As Variant()
    
    Set fso = New FileSystemObject
    query = fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_FeatureHeaderInfo.sql").ReadAll
    params = Array("r.RunName," & jobNum, "rt.RoutineName," & routine)
    
    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)
    
    If Not sqlRecordSet.EOF Then
        GetFeatureHeaderInfo = sqlRecordSet.GetRows()
        Exit Function
    End If

End Function

    'Get all observation values, filter out failures
Function GetFeatureMeasuredValues(jobNum As String, routine As String, delimFeatures As String, featureInfo() As Variant) As Variant()

    Set fso = New FileSystemObject
    query = Replace(Split(fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_FeatureMeasurements.sql").ReadAll, ";")(0), "{Features}", delimFeatures)
    
    Dim whereClause As String
    For i = 0 To UBound(featureInfo, 2)
        If featureInfo(6, i) = "Attribute" Then
            'Filter out Attribute failures by column
            whereClause = whereClause & "(Pvt.[" & featureInfo(0, i) & "] <> 1 OR Pvt.[" & featureInfo(0, i) & "] IS NULL)"
        Else
            'Filter out Variable failure by column
            whereClause = whereClause & "(Pvt.[" & featureInfo(0, i) & "] <> 99.998 OR Pvt.[" & featureInfo(0, i) & "] IS NULL)"
        End If
        'if its the last statement, no need for the AND
        If i <> UBound(featureInfo, 2) Then
            whereClause = whereClause & " AND "
        End If
    Next i

    query = query & whereClause
    params = Array("r.RunName," & jobNum, "rt.RoutineName," & routine, "r.RunName," & jobNum, "rt.RoutineName," & routine)
    
    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)

    If Not sqlRecordSet.EOF Then
        GetFeatureMeasuredValues = sqlRecordSet.GetRows()
        Exit Function
    End If

End Function

    'Dont filter out any values
Function GetAllFeatureMeasuredValues(jobNum As String, routine As String, delimFeatures As String) As Variant()

    Set fso = New FileSystemObject
    query = Replace(Split(fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_FeatureMeasurements.sql").ReadAll, ";")(1), "{Features}", delimFeatures)
    params = Array("r.RunName," & jobNum, "rt.RoutineName," & routine, "r.RunName," & jobNum, "rt.RoutineName," & routine)
    
    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)

    If Not sqlRecordSet.EOF Then
        GetAllFeatureMeasuredValues = sqlRecordSet.GetRows()
        Exit Function
    End If

End Function

    'Date, Employee ID - Filter out failed observations
Function GetFeatureTraceabilityData(jobNum As String, routine As String) As Variant()

    Set fso = New FileSystemObject
    query = Split(fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_ObsTraceability.sql").ReadAll, ";")(0)
    params = Array("r.RunName," & jobNum, "rt.RoutineName," & routine, "r.RunName," & jobNum, "rt.RoutineName," & routine)
    
    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)

    If Not sqlRecordSet.EOF Then
        GetFeatureTraceabilityData = sqlRecordSet.GetRows()
        Exit Function
    End If

End Function

    'Date, Employee ID - Dont Filter out failed observations
Function GetAllFeatureTraceabilityData(jobNum As String, routine As String) As Variant()

    Set fso = New FileSystemObject
    query = Split(fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_ObsTraceability.sql").ReadAll, ";")(1)
    params = Array("r.RunName," & jobNum, "rt.RoutineName," & routine, "r.RunName," & jobNum, "rt.RoutineName," & routine)
    
    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)

    If Not sqlRecordSet.EOF Then
        GetAllFeatureTraceabilityData = sqlRecordSet.GetRows()
        Exit Function
    End If
End Function

    'Called by userform to determine how many Inspections it should require for FI_DIM
Function IsAllAttribrute(routine As Variant) As Boolean

    Set fso = New FileSystemObject
    query = fso.OpenTextFile(DataSources.QUERIES_PATH & "AnyVariables.sql").ReadAll
    params = Array("r.RoutineName," & routine)

    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)

    If sqlRecordSet.EOF Then
        IsAllAttribrute = True
    Else
        IsAllAttribrute = False
    End If

End Function

    'All of the routines set up for this Part Number
Function GetPartRoutineList(partNum As String, Revision As String) As Variant()
    Set fso = New FileSystemObject
    Dim mlPartNum As String
    mlPartNum = partNum & "_" & Revision
    query = fso.OpenTextFile(DataSources.QUERIES_PATH & "PartRoutineList.sql").ReadAll
    params = Array("p.PartName," & mlPartNum)

    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)
    
    If Not sqlRecordSet.EOF Then
        GetPartRoutineList = sqlRecordSet.GetRows()
        Exit Function
    End If
End Function

    'All the routines created for this run
Function GetRunRoutineList(jobNum As String) As Variant()
    Set fso = New FileSystemObject
    query = fso.OpenTextFile(DataSources.QUERIES_PATH & "RunRoutineList.sql").ReadAll
    params = Array("r.RunName," & jobNum)

    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.ML7)
    
    If Not sqlRecordSet.EOF Then
        GetRunRoutineList = sqlRecordSet.GetRows()
        Exit Function
    End If
End Function









''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               InspectionKiosk
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetCustomerName(jobNum As String) As String
    Set fso = New FileSystemObject
    query = "SELECT CustomerName FROM InspectionKiosk.dbo.CustomerTranslation WHERE Abbreviation=?"
    Dim searchParam As String

    'If our job is an inventory job like 'NVxxx' then, we can just search by the first two characters
    If Len(jobNum) > 2 And Not IsNumeric(Left(jobNum, 1)) And Not IsNumeric(Mid(jobNum, 2, 1)) Then
        searchParam = Left(jobNum, 2)
        GoTo 20
    End If

    GetJobInformation JobID:=jobNum, custName:=searchParam

20

    params = Array("Abbreviation," & searchParam)

    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.Kiosk)


    If Not sqlRecordSet.EOF Then
        GetCustomerName = sqlRecordSet.Fields(0).Value
        Exit Function
    End If

End Function

Function GetCellLeadEmail(cell As String) As String
    Set fso = New FileSystemObject
    query = "SELECT Email FROM InspectionKiosk.dbo.VettingEmails WHERE Cell=?"
    params = Array("Cell," & cell)

    Call ExecQuery(query:=query, params:=params, conn_enum:=Connections.Kiosk)

    If Not sqlRecordSet.EOF Then
        GetCellLeadEmail = sqlRecordSet.Fields(0).Value
        Exit Function
    End If

'Error here

End Function


