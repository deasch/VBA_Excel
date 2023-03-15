Function ValueIsInArray(varValueToBeFound As Variant, varArray As Variant) As Boolean
Dim varElement As Variant
On Error GoTo ValueIsInArray: 'array is empty
    For Each varElement In varArray
        If varElement = varValueToBeFound Then
            ValueIsInArray = True
            Exit Function
        End If
    Next varElement
Exit Function
ValueIsInArray:
On Error GoTo 0
ValueIsInArray = False
End Function
