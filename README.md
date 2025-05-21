Function ExtractTextWithChar(rng As Range, charToFind As String) As Variant
    Dim cell As Range
    Dim result() As String
    Dim count As Long
    count = 0

    For Each cell In rng
        If VarType(cell.Value) = vbString Then
            If InStr(1, cell.Value, charToFind, vbTextCompare) > 0 Then
                ReDim Preserve result(count)
                result(count) = cell.Value
                count = count + 1
            End If
        End If
    Next cell

    If count = 0 Then
        ExtractTextWithChar = Application.Transpose(Array())
    Else
        ExtractTextWithChar = Application.Transpose(result)
    End If
End Function
