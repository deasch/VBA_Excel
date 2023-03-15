Function RegEx_CapturingGroup(strText As String, strRegExPattern As String,Optional, blnRegExMultiline as Boolean, Optional blnRegExIgnoreCase As Boolean) As String()
'RegExp is a Component Object Model(COM), which we need to reference in the VBA editor. To enable RegExp set the reference for Microsoft VBScript Regular Expression 5.5
    
    Dim objRegEx As Object
    Set objRegEx = New RegExp

    With objRegEx
        .Pattern = strRegExPattern
        .Global = True
        .Multiline = blnRegExMultiline
        .IgnoreCase = blnRegExIgnoreCase
    End With
    
    If objRegEx.Test(strText) Then
        'Debug.Print "pattern is matched (found)"
        
        Dim objRegExMatches As Object
        Set objRegExMatches = objRegEx.Execute(strText)
        
        Dim lngRegExMatchesCount As Long
        lngRegExMatchesCount = objRegExMatches.Count
        'Debug.Print "Matches: " & objRegExMatches.Count
        
        Dim a As Long
        Dim arrCapturingGroup() As String
        ReDim arrCapturingGroup(Int(lngRegExMatchesCount - 1))
        For a = 0 To lngRegExMatchesCount - 1
            'Debug.Print a & " | " & objRegExMatches(a - 1).Value
            arrCapturingGroup(a) = objRegExMatches(a).Value
        Next a
       
    Else
        'Debug.Print "pattern is not matched (found)"
    End If
    
    Set objRegExMatches = Nothing
    Set objRegEx = Nothing
    
    RegEx_CapturingGroup = arrCapturingGroup()
End Function
