Sub HttpRequest_POST()
    Dim objXMLHttpRequest As Object: Set objXMLHttpRequest = CreateObject("MSXML2.XMLHTTP")
    Dim strXHMHttpResponseText As String
    
    Dim strUrl As String: strUrl = "https://gssnplus-int.i.mercedes-benz.com/int/gssnplus-api/api/v1/outlets/search?page=0&pageSize=10"
    Dim strContentType As String: strContentType = "application/json"
    Dim strAuth As String: strAuth = "a853a413-5f45-461c-a3d6-3aa5a8b8c6a3"
    Dim strPayload As String: strPayload = "{  ""names"": [    ""*Burger Schloz*""  ]}"
    
    With objXMLHttpRequest
        .Open "POST", strUrl, False
        .setRequestHeader "Content-Type", strContentType
        .setRequestHeader "x-api-key", strAuth
        .send strPayload
        strXHMHttpResponseText = .responseText
        'Debug.Print .responseText
        'Debug.Print .getAllResponseHeaders
        'Debug.Print .Status
    End With
    
    Set objXMLHttpRequest = Nothing
End sub
