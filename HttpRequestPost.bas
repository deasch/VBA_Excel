Sub HttpRequestPost()
    Dim objXMLHttpRequest As Object: Set objXMLHttpRequest = CreateObject("MSXML2.XMLHTTP")
    Dim strXHMHttpResponseText As String
    
    Dim strUrl As String: strUrl = "https://"
    Dim strContentType As String: strContentType = "application/json"
        Dim strAuth As String: strAuth = "" 'Bsp: x-api-key/bearer/authorization
        Dim strPayload As String: strPayload = "{  ""key"": [    ""*value*""  ]}"
    
    With objXMLHttpRequest
        .Open "POST", strUrl, False
        .setRequestHeader "Content-Type", strContentType
        .setRequestHeader "x-api-key", strAuth 'Bsp: x-api-key/bearer/authorization
        .send strPayload
        strXHMHttpResponseText = .responseText
        'Debug.Print .responseText
        'Debug.Print .getAllResponseHeaders
        'Debug.Print .Status
    End With
    
    objXMLHttpRequest = Nothing
End sub
