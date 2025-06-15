Function ultimaLinha(planilha As String, col As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(planilha)
    ultimaLinha = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

'busca linha por codigo
Function getClienteByCodigo(codigo As Variant) As cliente
    Dim ws As Worksheet
    Dim cl As cliente
    
    Set ws = ThisWorkbook.Sheets("CLIENTES_DADOS")
    ul = ultimaLinha("CLIENTES_DADOS", 1)
    
    For I = 1 To ul
        If ws.Cells(I, 1).Value = codigo Then
            Set cl = New cliente
            cl.codigo = ws.Cells(I, 1).Value
            cl.nome = ws.Cells(I, 2).Value
            cl.cpf = ws.Cells(I, 3).Value
            cl.telefone = ws.Cells(I, 4).Value
            cl.email = ws.Cells(I, 5).Value
            cl.endereco = ws.Cells(I, 6).Value
            
            Set getClienteByCodigo = cl
            Exit Function
        End If
    Next I
    Set getClienteByCodigo = Nothing
End Function

'atualiza linha
Function updateCliente(cliente As cliente) As Boolean
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("CLIENTES_DADOS")
    ul = ultimaLinha("CLIENTES_DADOS", 1)
    
    For I = 2 To ul
        If CStr(ws.Cells(I, 1).Value) = CStr(cliente.codigo) Then
            ws.Cells(I, 2).Value = cliente.nome
            ws.Cells(I, 3).Value = cliente.cpf
            ws.Cells(I, 4).Value = cliente.telefone
            ws.Cells(I, 5).Value = cliente.email
            ws.Cells(I, 6).Value = cliente.endereco
            
            updateCliente = True
            Exit Function
        End If
    Next I
    updateCliente = False
End Function
