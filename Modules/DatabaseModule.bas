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

Sub Init_Connections()

    On Error GoTo Err_Conn
    
    If ML7DataBaseConnection Is Nothing Then
        Set ML7DataBaseConnection = New ADODB.Connection
        ML7DataBaseConnection.ConnectionString = ML7_CONN_STRING
        ML7DataBaseConnection.Open
    End If
      
    If E10DataBaseConnection Is Nothing Then
        Set E10DataBaseConnection = New ADODB.Connection
        E10DataBaseConnection.ConnectionString = E10_CONN_STRING
        E10DataBaseConnection.Open
    End If
    
    If KioskDataBaseConnection Is Nothing Then
        Set KioskDataBaseConnection = New ADODB.Connection
        KioskDataBaseConnection.ConnectionString = KIOSK_CONN_STRING
        KioskDataBaseConnection.Open
    End If
       
        
    Exit Sub
    
Err_Conn:
    Err.Raise Number:=Err.Number, description:="There was an error connecting with the Epicor and/or MeasurLink Database " _
        & "you may not be connected to the Network or you may not have permission from the Administrator to read from the MeasurLink DataBase"

End Sub





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               Epicor
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetJobInformation(JobID As String, Optional ByRef partNum As Variant, Optional ByRef rev As Variant, _
                                Optional ByRef setupType As Variant, Optional ByRef custName As Variant, _
                                Optional ByRef machine As Variant, Optional ByRef cell As Variant, _
                                Optional ByRef partDescription As Variant) As Boolean
    Call Init_Connections
    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = E10DataBaseConnection
        .CommandType = adCmdText
        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "EpicorJobInfo.sql").ReadAll
                        
        
        Dim jobParam As ADODB.Parameter
        Set jobParam = .CreateParameter(Name:="jh.JobNum", Type:=adVarChar, Size:=14, Direction:=adParamInput)
        jobParam.Value = JobID
        .Parameters.Append jobParam
    End With
        
    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.Open sqlCommand
    
    'If any rows at all were returned, we know that the job exists
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
        GetJobInformation = True
        Exit Function
    End If

    GetJobInformation = False
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               MeasurLink
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Retrieve the Header Information for the Features applicable to our Run and Routine
Function GetFeatureHeaderInfo(jobNum As String, routine As String) As Variant()
    Call Init_Connections

    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = ML7DataBaseConnection
        .CommandType = adCmdText
        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_FeatureHeaderInfo.sql").ReadAll
        
        
        Dim partParam As ADODB.Parameter
        Set partParam = .CreateParameter(Name:="r.RunName", Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=jobNum)
        Dim partParam2 As ADODB.Parameter
        Set partParam2 = .CreateParameter(Name:="rt.RoutineName", Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=routine)
        
        .Parameters.Append partParam
        .Parameters.Append partParam2

    End With

    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
    

    If Not sqlRecordSet.EOF Then
        GetFeatureHeaderInfo = sqlRecordSet.GetRows()
        Exit Function
    End If

End Function


Function GetFeatureMeasuredValues(jobNum As String, routine As String, features As String) As Variant()
    Call Init_Connections

    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = ML7DataBaseConnection
        .CommandType = adCmdText
        .CommandText = Replace(fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_FeatureMeasurements.sql").ReadAll, "{Features}", features)
        
        Dim params() As Variant
        params = Array("r.RunName", "rt.RoutineName")
        Dim values() As Variant
        values = Array(jobNum, routine)

        For i = 0 To 3
            Dim partParam As ADODB.Parameter
            Set partParam = .CreateParameter(Name:=params(i Mod 2), Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=values(i Mod 2))
            .Parameters.Append partParam
        Next i

    End With

    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
    

    If Not sqlRecordSet.EOF Then
        GetFeatureMeasuredValues = sqlRecordSet.GetRows()
        Exit Function
    End If

End Function


Function GetFeatureTraceabilityData(jobNum As String, routine As String) As Variant()
    Call Init_Connections

    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = ML7DataBaseConnection
        .CommandType = adCmdText
        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "ML_ObsTraceability.sql").ReadAll
        
        Dim params() As Variant
        params = Array("r.RunName", "rt.RoutineName")
        Dim values() As Variant
        values = Array(jobNum, routine)

        For i = 0 To 3
            Dim partParam As ADODB.Parameter
            Set partParam = .CreateParameter(Name:=params(i Mod 2), Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=values(i Mod 2))
            .Parameters.Append partParam
        Next i

    End With

    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
    

    If Not sqlRecordSet.EOF Then
        GetFeatureTraceabilityData = sqlRecordSet.GetRows()
        Debug.Print ("stopping here")
        Exit Function
    End If

End Function



'see how this can be called recursively??
'Function SanityTest()
'    Call Init_Connections
'
'    Set fso = New FileSystemObject
'
'    Set sqlCommand = New ADODB.Command
'    With sqlCommand
'        .ActiveConnection = ML7DataBaseConnection
'        .CommandType = adCmdText
'        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "ParameterSanityTest.sql").ReadAll
'
'        Dim params() As Variant
'        params = Array("r.RunName", "rt.RoutineName")
'        Dim values() As Variant
'        values = Array("SD1284", "DRW-00717-01_RAJ_IP_IXSHIFT")
'
'        For i = 0 To 3
'            Dim partParam As ADODB.Parameter
'            Set partParam = .CreateParameter(Name:=params(i Mod 2), Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=values(i Mod 2))
'            .Parameters.Append partParam
'        Next i
'
'    End With
'
'    Set sqlRecordSet = New ADODB.Recordset
'    sqlRecordSet.CursorLocation = adUseClient
'    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
'
'
'    If Not sqlRecordSet.EOF Then
'        While Not sqlRecordSet.EOF
'            With sqlRecordSet
'                Debug.Print (.Fields(0).Value & vbTab & .Fields(1).Value & vbTab & .Fields(2).Value)
'            End With
'
'            sqlRecordSet.MoveNext
'
'        Wend
'
'
'    End If
'
'End Function

' This and the function below it need to be compressed
Function GetPartRoutineList(partNum As String, Revision As String) As ADODB.Recordset
    Call Init_Connections

    Set fso = New FileSystemObject
    Dim mlPartNum As String
    mlPartNum = partNum & "_" & Revision

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = ML7DataBaseConnection
        .CommandType = adCmdText
        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "PartRoutineList.sql").ReadAll
        
        Dim partParam As ADODB.Parameter
        Set partParam = .CreateParameter(Name:="p.PartName", Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=mlPartNum)
        .Parameters.Append partParam
    End With

    Set sqlRecordSet = New ADODB.Recordset
    'Location and Static type allow us to access the total number of records, will need this for callback function later
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
    

    If Not sqlRecordSet.EOF Then
        Set GetPartRoutineList = sqlRecordSet.Clone
        Exit Function
    End If

    'TODO: Error here on the available Routines, None should be handled differently than an actual error
    Set GetPartRoutineList = Nothing
End Function

Function GetRunRoutineList(jobNum As String) As ADODB.Recordset
    Call Init_Connections

    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = ML7DataBaseConnection
        .CommandType = adCmdText
        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "RunRoutineList.sql").ReadAll
        
        Dim partParam As ADODB.Parameter
        Set partParam = .CreateParameter(Name:="r.RunName", Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=jobNum)
        .Parameters.Append partParam
    End With

    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
    

    If Not sqlRecordSet.EOF Then
        Set GetRunRoutineList = sqlRecordSet.Clone
        Exit Function
    End If

    'TODO: Error here on the available Routines, None should be handled differently than an actual error
    Set GetRunRoutineList = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'               InspectionKiosk
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'TODO: this is a WIP, needs to be tested still

Function GetCustomerName(jobNum As String) As String
    Call Init_Connections

    Dim searchParam As String

    'If our job is an inventory job like 'NVxxx' then, we can just search by the first two characters
    If Len(jobNum) > 2 And Not IsNumeric(Left(jobNum, 1)) And Not IsNumeric(Mid(jobNum, 2, 1)) Then
        searchParam = Left(jobNum, 2)
        GoTo 20
    End If

    GetJobInformation JobID:=jobNum, custName:=searchParam

20
    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = KioskDataBaseConnection
        .CommandType = adCmdText
        .CommandText = "SELECT CustomerName FROM InspectionKiosk.dbo.CustomerTranslation WHERE Abbreviation=?"

        Dim partParam As ADODB.Parameter
        Set partParam = .CreateParameter(Name:="Abbreviation", Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=searchParam)
        .Parameters.Append partParam
    End With

    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic


    If Not sqlRecordSet.EOF Then
        GetCustomerName = sqlRecordSet.Fields(0).Value
        Exit Function
    End If


    'TODO: Error here, we don't can't find the customer name in our table, the QE should update the Database
    GetCustomerName = vbNullString

End Function






