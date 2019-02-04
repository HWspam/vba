Sub homework():

Dim equals As Double
get_count = Cells(Rows.Count, "A").End(xlUp).Row
Range("H1").Value = "tick ID"
Range("I1").Value = "stock volume"

For h = 2 To get_count
    If Cells(h + 1, 1).Value <> Cells(h, 1).Value Then
        equals = equals + Cells(h, 7).Value
        Range("H" & 2 + i).Value = Cells(h, 1).Value
        Range("I" & 2 + i).Value = equals
        equals = 0
        i = i + 1
    Else
        equals = equals + Cells(h, 7).Value
    End If
Next h

End Sub