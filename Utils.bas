Function ultimaLinha(planilha As String, col As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(planilha)
    ultimaLinha = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function
