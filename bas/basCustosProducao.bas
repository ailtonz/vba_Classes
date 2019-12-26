Attribute VB_Name = "basCustosProducao"
Public Sub cadastroCustoProducao()

Dim wsCP As Worksheet
Dim CP As clsCustoProducao

Dim lRow As Long, x As Long

Set wsCP = Worksheets("CUSTOS_PRODUCAO")
Set CP = New clsCustoProducao

carregarBanco
    
''find  first empty row in database
lRow = wsCP.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Row
        
    For x = 2 To lRow - 1

        With CP
            .ID = wsCP.Range("A" & x).Value
            .Paginas = wsCP.Range("B" & x).Value
            .Valor = wsCP.Range("C" & x).Value
            .Tipo = wsCP.Range("D" & x).Value
            .Estilo = wsCP.Range("E" & x).Value
            .SubTipo = wsCP.Range("F" & x).Value
            .add CP
        End With

        If CP.ID = "" Then
            CP.Insert Bnc, CP
        ElseIf CP.ID <> "" And CP.Paginas <> "" Then
            CP.Update Bnc, CP
        Else
            CP.Delete Bnc, CP
        End If

    Next x

Set CP = Nothing
Set Bnc = Nothing

End Sub

Sub ListarCustosProducao()

carregarBanco

Dim prf As clsCustoProducao
Set prf = New clsCustoProducao

Dim col As clsCustoProducao
Set col = prf.getCustosProducao(Bnc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("CUSTOS_PRODUCAO")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each prf In col.Itens
        
        ws.Range("A" & lRow).Value = prf.ID
        ws.Range("B" & lRow).Value = prf.Paginas
        ws.Range("C" & lRow).Value = prf.Valor
        ws.Range("D" & lRow).Value = prf.Tipo
        ws.Range("E" & lRow).Value = prf.Estilo
        ws.Range("F" & lRow).Value = prf.SubTipo
        
        lRow = lRow + 1
    Next prf

End Sub

