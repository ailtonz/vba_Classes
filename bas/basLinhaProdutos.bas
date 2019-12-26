Attribute VB_Name = "basLinhaProdutos"
Public Sub cadastroLinhaProduto()

Dim ws As Worksheet
Dim obj As clsLinhaProdutos
Dim lRow As Long, x As Long

Set ws = Worksheets("LINHAS")
Set obj = New clsLinhaProdutos
    
carregarBanco

''find  first empty row in database
lRow = ws.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Row
            
    For x = 2 To lRow - 1
            
        With obj
            .ID = ws.Range("A" & x).Value
            .Linha = ws.Range("B" & x).Value
            .Maximo = ws.Range("C" & x).Value
            .Minimo = ws.Range("D" & x).Value
            .Estilo = ws.Range("E" & x).Value
            .add obj
        End With
        
        If obj.ID = "" Then
            obj.Insert Bnc, obj
        ElseIf obj.ID <> "" And obj.Linha <> "" Then
            obj.Update Bnc, obj
        Else
            obj.Delete Bnc, obj
        End If
        
    Next x
                      
Set IR = Nothing
Set Bnc = Nothing

End Sub

Sub ListarLinhas()

carregarBanco

Dim prf As clsLinhaProdutos
Set prf = New clsLinhaProdutos

Dim col As clsLinhaProdutos
Set col = prf.getLinhas(Bnc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("LINHAS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each prf In col.Itens
        
        ws.Range("A" & lRow).Value = prf.ID
        ws.Range("B" & lRow).Value = prf.Linha
        ws.Range("C" & lRow).Value = prf.Maximo
        ws.Range("D" & lRow).Value = prf.Minimo
        ws.Range("E" & lRow).Value = prf.Estilo
        
        lRow = lRow + 1
    Next prf

End Sub
