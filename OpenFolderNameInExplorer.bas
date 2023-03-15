Sub OpenFolderNameInExplorer(strFullFileName As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Shell "C:\WINDOWS\explorer.exe """ & fso.GetParentFolderName(strFullFileName) & "\" & """", vbNormalFocus
    Set fso = Nothing
End Sub
