Attribute VB_Name = "HTTPconnections"

'*************************************************************
'*************************************************************
'*                  HTTPconnections
'*
'*  Connect to the API at Jade76 and
'*    SELECT, UPDATE or INSERT Custom Field Information
'*    for each Part's Feautres
'*
'*
'*
'*     'Dictionaries in python are translated as VBA Dictionaries, Lists like Collections, so.....
'*         '{'hello': 'world'}    ->   parsed("hello") ->> "world"
'*         '{'hello': {'goodbye': 'world'}}   -> parsed("hello")("goodbye") ->> "world"
'*         '{'hello': [{'goodbye': 'world'}]}   -> parsed("hello")(1)("goobye") ->> "world"
'*
'*         'Just keep in mind that Collections are 1 based.
'*         'When trying to flatten results, use TypeName() -> 'Collection' | 'Dictionary' | [some scalar]
'*             'To figure out how to iterate through the items
'*
'*************************************************************
'*************************************************************



'****************************************************
'**************   Main Routine   ********************
'****************************************************

Private Function send_http(url As String, method As String, Optional payload As String, Optional q_params As Variant, Optional api_key As Variant) As String
    On Error GoTo HTTP_Err:

    Dim req As ServerXMLHTTP60
    Dim parsed As Object

    Set req = New ServerXMLHTTP60
    
    If Not IsMissing(q_params) Then
        'Set up the Url to add query parameters to the end
        'q_params(i)(0) -> key
        'q_params(i)(1) -> val
        
        url = url & "?"
        Dim i As Integer
        For i = 0 To UBound(q_params)
            If i > 0 Then
                url = url & "&"
            End If
            url = url & Replace(q_params(i)(0), " ", "%20") & "=" & Replace(q_params(i)(1), " ", "%20")
        Next i
    End If
    
    With req
        'Set request headers here...
        .Open method, url, False   'We can do this asyncronously??
        If method <> DataSources.HTTP_GET Then
            .setRequestHeader "Content-Type", "application/json;charset=utf-8"
        End If
        .setRequestHeader "Accept", "application/json;charset=utf-8"
        
        If Not IsMissing(api_key) Then
            .setRequestHeader "X-Request-ID", api_key
            .setRequestHeader "Authorization", Environ("Username")
        End If
        
        .Send payload
    End With

    Dim resp As String, header As String, headers As String
    headers = req.getAllResponseHeaders()
    
    Debug.Print (headers)
    Debug.Print (req.Status & vbTab & req.statusText)
    
    If req.Status <> 200 Then GoTo HTTP_Err
    'Should read the response type here and possible raise and error based on the different response types we can get
    
    send_http = req.responseText
    
    Exit Function
    
HTTP_Err:
    If req.readyState < 4 Then
        Err.Raise Number:=vbObjectError + 6010, Description:="send_http Error" & vbCrLf & vbCrLf & "No response from the server. The server may be down or the API service may not be running"

    ElseIf req.Status = 406 Or req.Status = 400 Or req.Status = 404 Then
        'Adding a user: Either not in QA department or they have already been reigstered
        Err.Raise Number:=vbObjectError + 6000 + req.Status, Description:=req.responseText
    Else
        'Unhandled HTTP Errors, Likely for Internal Server 500
        Err.Raise Number:=vbObjectError + 6000, Description:="send_http Error" & vbCrLf & headers & vbCrLf & "Status:" & req.Status & vbTab & req.statusText _
            & vbCrLf & "RequestBody: " & vbCrLf & req.responseText & vbCrLf & vbclrf
    End If
End Function


'****************************************************
'************   Public Callables   ******************
'****************************************************

Public Function ValidateDHR(job_num As String) As Object


    On Error GoTo DHR_Err:
    
    Dim resp As String, url As String
    url = DataSources.API_DHR & job_num & "/"
    resp = send_http(url:=url, method:=DataSources.HTTP_GET)
    Set ValidateDHR = JsonConverter.ParseJson(resp)
    
    Exit Function
    
DHR_Err:
    If Err.Number = vbObjectError + 6000 Then  'Unhandled Exceptions Like Internal Server Error
        MsgBox Err.Description
    ElseIf Err.Number = vbObjectError + 6010 Then  'Server Not Responding
        MsgBox Err.Description, vbExclamation
    Else
        MsgBox "Unexpected Exception Occured Func: HTTPConnections.ValidateDHR() when parsing JSON" & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If
End Function

Public Function GetPassedInspData(job_name As String, routine_name As String) As Object
'    {
'    "feature_info": [
'        {
'            "type": 2,
'            "name": "(i)",
'            "properties": null,  // or {
'                "nominal": 0.1861,
'                "utol": 0.1891,
'                "ltol": 0.1831
'            },
'            "custom_fields": [
'                {
'                    "valueType": 1,
'                    "customFieldId": 3,
'                    "value": "NA"
'                }]
'        },]
'    "insp_data": [
'        {
'            "ObsNo": 1,
'            "(i)": 0.0,
'            "(ii)": 0.0,
'            "(iii)": 0.0,
'            "0_006_00": 0.0,
'            "0_010_00": 0.0,
'            "0_012_00": 0.1852,
'            "0_024_00": 0.0
'        },]
    
    

    On Error GoTo PassedData_Err:
    
    Dim resp As String, url As String, q_params(1) As Variant
    q_params(0) = Array("job_name", job_name)
    q_params(1) = Array("routine_name", routine_name)
    
    
    resp = send_http(url:=DataSources.API_RUN_DATA_PASSED, method:=DataSources.HTTP_GET, q_params:=q_params)
    Set GetPassedInspData = JsonConverter.ParseJson(resp)
    
    Exit Function
    
PassedData_Err:
    If Err.Number = vbObjectError + 6000 Then  'Unhandled Exceptions Like Internal Server Error
        MsgBox Err.Description
    ElseIf Err.Number = vbObjectError + 6010 Then  'Server Not Responding
        MsgBox Err.Description, vbExclamation
    Else
        MsgBox "Unexpected Exception Occured Func: HTTPConnections.GetPassedInspData() when parsing JSON" & vbCrLf & vbCrLf & Err.Description, vbCritical
    End If
End Function






