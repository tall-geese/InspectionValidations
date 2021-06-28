Attribute VB_Name = "DatabaseModule"

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
    Err.Raise Number:=Err.Number, Description:="There was an error connecting with the Epicor and/or MeasurLink Database " _
        & "you may not be connected to the Network or you may not have permission from the Administrator to read from the MeasurLink DataBase"

End Sub



Function VerifyJobExists(JobID As String, ByRef PartNum As String, ByRef rev As String, ByRef setupType As String) As Boolean
    Call Init_Connections

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = E10DataBaseConnection
        .CommandType = adCmdText
        .CommandText = "SELECT jh.JobNum, jh.Company, jh.PartNum, jh.RevisionNum, jo.Character01 " _
                        & "FROM EpicorLive10.dbo.JobHead jh " _
                        & "LEFT OUTER JOIN EpicorLive10.dbo.JobOper jo ON jo.JobNum = jh.JobNum " _
                        & "WHERE jh.JobNum = ? AND jh.Company = 'JPMC' AND jo.OpCode IN ('SWISS','CNC')"
        
        Dim jobParam As ADODB.Parameter
        Set jobParam = .CreateParameter(Name:="jh.JobNum", Type:=adVarChar, Size:=14, Direction:=adParamInput)
        jobParam.Value = JobID
        .Parameters.Append jobParam
    End With
        
    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.Open sqlCommand
    
    'If any rows at all were returned, we know that the job exists
    If Not sqlRecordSet.EOF Then
        PartNum = sqlRecordSet.Fields(2).Value
        rev = sqlRecordSet.Fields(3).Value
        setupType = sqlRecordSet.Fields(4).Value
        VerifyJobExists = True
        Exit Function
    End If

    VerifyJobExists = False
End Function

'TODO:
Function GetCustomerFromProject(jobNum As String) As String
    Call Init_Connections

    Set fso = New FileSystemObject

    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = E10DataBaseConnection
        .CommandType = adCmdText
        .CommandText = fso.OpenTextFile(DataSources.QUERIES_PATH & "ProjectCusName.sql").ReadAll
        
        Dim partParam As ADODB.Parameter
        Set partParam = .CreateParameter(Name:="jh.JobNum", Type:=adVarChar, Size:=14, Direction:=adParamInput, Value:=jobNum)
        .Parameters.Append partParam
    End With

    Set sqlRecordSet = New ADODB.Recordset
    sqlRecordSet.CursorLocation = adUseClient
    sqlRecordSet.Open Source:=sqlCommand, CursorType:=adOpenStatic
    

    If Not sqlRecordSet.EOF And sqlRecordSet.RecordCount = 1 Then
        GetCustomerFromProject = sqlRecordSet.Fields(0).Value
        Exit Function
    End If

    'TODO: Error here, couldn't find the customer name
    GetCustomerFromProject = vbNullString
End Function

'TODO: there is certainly a way we can consolidate the below two functions into one, so much repeated code here

Function GetPartRoutineList(PartNum As String, Revision As String) As ADODB.Recordset
    Call Init_Connections

    Set fso = New FileSystemObject
    Dim mlPartNum As String
    mlPartNum = PartNum & "_" & Revision

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


Function GetCustomerName(jobNum As String) As String
    Call Init_Connections
    
    Dim searchParam As String
    
    'If our job is an inventory job like 'NVxxx' then, we can just search by the first two characters
    If Len(test) > 2 And Not IsNumeric(Left(test, 1)) And Not IsNumeric(Mid(test, 2, 1)) Then
        searchParam = Left(jobNum, 2)
        GoTo 20
    End If
    
    searchParam = GetCustomerFromProject(jobNum)
    
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

