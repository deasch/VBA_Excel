Sub ChangeColumnWidth(xlSheet As Worksheet, strColumnName As String, dblWidth As Double)
    xlSheet.Columns(strColumnName & ":" & strColumnName).ColumnWidth = dblWidth 'xx.xx Pixel
End Sub
