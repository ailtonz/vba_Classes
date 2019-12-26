Attribute VB_Name = "basProfissoes"
Public Sub cadastroProfissao()

Dim ws As Worksheet
Dim obj As clsProfissoes
Dim lRow As Long, x As Long
        
Set ws = Worksheets("PROFISSOES")
Set obj = New clsProfissoes
        
carregarBanco
            
''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1

        With obj
            .ID = ws.Range("A" & x).Value
            .Profissao = ws.Range("B" & x).Value
            .add obj
        End With
        
        If obj.ID = "" Then
            obj.Insert Bnc, obj
        ElseIf obj.ID <> "" And obj.Profissao <> "" Then
            obj.Update Bnc, obj
        Else
            obj.Delete Bnc, obj
        End If
        
    Next x
                      
Set obj = Nothing
Set Bnc = Nothing

End Sub

Sub ListarProfissoes()

carregarBanco

Dim prf As clsProfissoes
Set prf = New clsProfissoes

Dim col As clsProfissoes
Set col = prf.getProfissoes(Bnc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("PROFISSOES")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each prf In col.Itens
        ws.Range("A" & lRow).Value = prf.ID
        ws.Range("B" & lRow).Value = prf.Profissao
        lRow = lRow + 1
    Next prf

End Sub
