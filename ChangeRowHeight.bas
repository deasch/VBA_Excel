Sub ChangeRowHeight(xlSheet As Worksheet, lngRowNumber As Long, dblHeight As Double)
  xlSheet.Rows(lngRowNumber & ":" & lngRowNumber).RowHeight = dblHeight ' xx.xx Pixel
End Sub
