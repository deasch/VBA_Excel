Function FolderNameExists(strFolderName As String) As Boolean
    If Dir(strFolderName, vbDirectory) = "" Then
        MsgBox "The selected folder doesn't exist"
        FolderNameExists = False
    Else
        MsgBox "The selected folder exists"
        FolderNameExists = True
    End If
End Function
