Attribute VB_Name = "basGrand"
Public Sub cadastroGrand()

Dim wsGrd As Worksheet
Dim Grd As clsGrands
Dim Orc As clsOrcamentos
Dim lRow As Long, x As Long
        
Set wsGrd = Worksheets("GRANDS")
Set Orc = New clsOrcamentos
Set Grd = New clsGrands

carregarBanco
    
''find  first empty row in database
lRow = wsGrd.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
    
        With Orc
            .Controle = wsGrd.Range("B" & x).Value
            .Vendedor = wsGrd.Range("C" & x).Value
            .add Orc
        End With
        
        With Grd
            .ID = wsGrd.Range("A" & x).Value
            .Profissao = wsGrd.Range("D" & x).Value
            .Nome = wsGrd.Range("E" & x).Value
            .ValorLiquido = wsGrd.Range("F" & x).Value
            .add Grd
        End With
                          
        If Grd.ID = "" Then
            Grd.Insert Bnc, Orc, Grd
        ElseIf Grd.ID <> "" And Grd.Nome <> "" Then
            Grd.Update Bnc, Orc, Grd
        Else
            Grd.Delete Bnc, Orc, Grd
        End If
    
    Next x

Set Orc = Nothing
Set Grd = Nothing
Set Bnc = Nothing


End Sub


Sub ListarGrands()

carregarBanco

Dim prf As clsGrands
Set prf = New clsGrands

Dim col As clsGrands
Set col = prf.getGrands(Bnc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("GRANDS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For Each prf In col.Itens
        ws.Range("A" & lRow).Value = prf.ID
        ws.Range("D" & lRow).Value = prf.Profissao
        ws.Range("E" & lRow).Value = prf.Nome
        ws.Range("F" & lRow).Value = prf.ValorLiquido
        lRow = lRow + 1
    Next prf


End Sub
