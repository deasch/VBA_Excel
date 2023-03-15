Sub OpenParentFolderNameInExplorer(strFullFileName As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Shell "C:\WINDOWS\explorer.exe """ & fso.GetParentFolderName(strFullFileName) & """", vbNormalFocus
End Sub
