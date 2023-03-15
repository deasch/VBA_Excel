Function FileNameExists(strFullFileName As String) As Boolean
   If Dir(strFullFileName) = "" Then
        MsgBox "The selected file doesn't exist"
        FileNameExists = False
    Else
        MsgBox "The selected file exists"
        FileNameExists = True
    End If
End Function
