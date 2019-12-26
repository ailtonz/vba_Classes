Attribute VB_Name = "tmp"
'Sub listagemProfissoes()
'carregarBanco
'
'Dim Connection As New ADODB.Connection
'Set Connection = OpenConnection(Bnc)
'
'Dim rst As ADODB.Recordset
'Dim cd As ADODB.Command
'Dim obj As clsProfissoes
'
'Dim lRow As Long, x As Long
'
'Dim ws As Worksheet
'
'Set ws = Worksheets("PROFISSOES")
'Set obj = New clsProfissoes
'Set cd = New ADODB.Command
'
'With cd
'    .ActiveConnection = Connection
'    .CommandText = "select * from qryProfissoes"
'    .CommandType = adCmdText
'    Set rst = .Execute
'End With
'
'x = 2
'Do While Not rst.EOF
'
'    ws.Range("A" & x).Value = rst.Fields("codCategoria")
'    ws.Range("B" & x).Value = rst.Fields("Descricao")
'
'    rst.MoveNext
'    x = x + 1
'Loop
'
'Connection.Close
'
'Set obj = Nothing
'Set cd = Nothing
'
'End Sub
