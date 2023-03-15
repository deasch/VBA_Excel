Function GetFolderName(strFullFileName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderName = fso.GetParentFolderName(strFullFileName) & "\"
    Set fso = Nothing
End Function
