

Sub ApplyAllAdjustments()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim i As Long
    i = 2 ' Start at row 2 (assuming headers)

    Do While i < ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim nextRowText As String
        nextRowText = LCase(Trim(ws.Cells(i + 1, 1).Value))

        If InStr(nextRowText, "rabatt:") > 0 Or _
           InStr(nextRowText, "klubbpris:") > 0 Or _
           Instr(nextRowText, "Prisnedsättning") > 0 Or _
           InStr(nextRowText, "+pant") > 0 Then

            ' Add the price (column C) from next row to current row
            ws.Cells(i, 3).Value = ws.Cells(i, 3).Value + ws.Cells(i + 1, 3).Value

            ' Delete the next row
            ws.Rows(i + 1).Delete
        Else
            i = i + 1
        End If
    Loop

    MsgBox "Rabatt, Klubbpris and Pant processed successfully!", vbInformation
End Sub
