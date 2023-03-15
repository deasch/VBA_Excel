Function GetParentFolderName(strFullFileName As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName = fso.GetParentFolderName(strFullFileName)
    Set fso = Nothing
End Function
