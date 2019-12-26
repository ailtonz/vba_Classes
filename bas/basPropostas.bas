Attribute VB_Name = "basPropostas"
Public Sub impressaoPropostas()

Dim ws As Worksheet, wsPrj As Worksheet
Dim obj As clsProposta, objPrj As clsProjeto, col As New clsProjeto
Dim lRow As Long, x As Long
        
Set ws = Worksheets("PROPOSTAS")
Set wsPrj = Worksheets("PROJETOS")

Set obj = New clsProposta

            
    ''find  first empty row in database
    lRow = ws.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1

        With obj
            .ArqCaminho = ws.Range("A" & x).Value
            .ArqNome = ws.Range("B" & x).Value
            .Controle = ws.Range("C" & x).Value
            .Cliente = ws.Range("D" & x).Value
            .Responsavel = ws.Range("E" & x).Value
            .Projeto = ws.Range("F" & x).Value
            .Journal = ws.Range("G" & x).Value
            .Autor = ws.Range("H" & x).Value
            .Publisher = ws.Range("I" & x).Value
            .GerenteNome = ws.Range("J" & x).Value
            .GerenteTelefone = ws.Range("K" & x).Value
            .GerenteCelular01 = ws.Range("L" & x).Value
            .GerenteCelular02 = ws.Range("M" & x).Value
            .GerenteIDnextel = ws.Range("N" & x).Value
            .add obj
        End With
        
    Next x
    
    ''find  first empty row in database
    lRow = wsPrj.Cells(Rows.count, 2).End(xlUp).Offset(1, 0).Row
    
    For x = 2 To lRow - 1
        Set objPrj = New clsProjeto
        With objPrj
            .ID = x
            .Opcao = wsPrj.Range("A" & x).Value
            .Idioma = wsPrj.Range("B" & x).Value
            .Volume = wsPrj.Range("C" & x).Value
            .PrcVendas = wsPrj.Range("D" & x).Value
            .PrcTotal = wsPrj.Range("E" & x).Value
            col.add objPrj
        End With
        
    Next x
    
    obj.GerarProposta obj, col
                      
Set obj = Nothing
Set objPrj = Nothing
Set Bnc = Nothing

End Sub
