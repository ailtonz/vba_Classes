Attribute VB_Name = "basIR"
Public Sub cadastroIR()

Dim wsIR As Worksheet
Dim IR As clsIR
Dim lRow As Long, x As Long
        
Set wsIR = Worksheets("IR")
Set IR = New clsIR

carregarBanco
        
''find  first empty row in database
lRow = wsIR.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With IR
            .ID = wsIR.Range("A" & x).Value
            .Ano = CStr(wsIR.Range("B" & x).Value)
            .Descricao = CStr(wsIR.Range("C" & x).Value)
            .FaixaInicial = CStr(wsIR.Range("D" & x).Value)
            .FaixaFinal = CStr(wsIR.Range("E" & x).Value)
            .Aliquota = CStr(wsIR.Range("F" & x).Value)
            .ParcelaDeduzir = CStr(wsIR.Range("G" & x).Value)
            .add IR
        End With
        
        If IR.ID = "" Then
            IR.Insert Bnc, IR
        ElseIf IR.ID <> "" And IR.Descricao <> "" Then
            IR.Update Bnc, IR
        Else
            IR.Delete Bnc, IR
        End If
        
    Next x
                      
Set IR = Nothing
Set Bnc = Nothing

End Sub

Sub ListarIR()

carregarBanco

Dim prf As clsIR
Set prf = New clsIR

Dim col As clsIR
Set col = prf.getIR(Bnc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("IR")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each prf In col.Itens
        
        ws.Range("A" & lRow).Value = prf.ID
        
        ws.Range("B" & lRow).Value = prf.Ano
        ws.Range("C" & lRow).Value = prf.Descricao
        ws.Range("D" & lRow).Value = prf.FaixaInicial
        ws.Range("E" & lRow).Value = prf.FaixaFinal
        ws.Range("F" & lRow).Value = prf.Aliquota
        ws.Range("G" & lRow).Value = prf.ParcelaDeduzir
        
        lRow = lRow + 1
    Next prf

End Sub

