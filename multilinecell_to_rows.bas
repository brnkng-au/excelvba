# Modified from https://stackoverflow.com/a/39652521
# CC-by-SA

Sub Split()

For Each Cell In Range("A1", Range("A2").End(xlDown))
    If InStr(1, Cell, Chr(10)) <> 0 Then
        tmpArr = Split(Cell, Chr(10))

        Cell.EntireRow.Copy
        Cell.Offset(1, 0).Resize(UBound(tmpArr), 1). _
            EntireRow.Insert xlShiftDown

        Cell.Resize(UBound(tmpArr) + 1, 1) = Application.Transpose(tmpArr)
    End If
Next

Application.CutCopyMode = False

End Sub
