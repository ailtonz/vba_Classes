Attribute VB_Name = "basEstilo"
Dim Bnc As New clsBancos

Public Sub cadastroEstilo()

Dim ws As Worksheet
Dim obj As clsEstilos
Dim lRow As Long, x As Long
        
Set ws = Worksheets("ESTILOS")
Set obj = New clsEstilos

carregarBanco

''find  first empty row in database
lRow = ws.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
            
        With obj
            .ID = ws.Range("A" & x).Value
            .Estilo = ws.Range("B" & x).Value
            .add obj
        End With
        
        If obj.ID = "" Then
            obj.Insert Bnc, obj
        ElseIf obj.ID <> "" And obj.Estilo <> "" Then
            obj.Update Bnc, obj
        Else
            obj.Delete Bnc, obj
        End If
        
    Next x
                      
Set IR = Nothing
Set Bnc = Nothing

End Sub

Sub ListarEstilos()

carregarBanco

Dim prf As clsEstilos
Set prf = New clsEstilos

Dim col As clsEstilos
Set col = prf.getEstilos(Bnc)

Dim lRow As Long, x As Long
Dim ws As Worksheet
Set ws = Worksheets("ESTILOS")

''find  first empty row in database
lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row

    For Each prf In col.Itens
        ws.Range("A" & lRow).Value = prf.ID
        ws.Range("B" & lRow).Value = prf.Estilo
        lRow = lRow + 1
    Next prf

End Sub
