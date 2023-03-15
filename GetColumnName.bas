Function GetColumnName(ByVal rngCell As Range) As String
    GetColumnName = Split(rngCell.Address, "$")(1)
End Function
