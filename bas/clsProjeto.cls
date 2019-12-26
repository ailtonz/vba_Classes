VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pOpcao As String
Private pIdioma As String
Private pVolume As String
Private pPrcVendas As String
Private pPrcTotal As String

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Opcao() As String
    Opcao = pOpcao
End Property

Public Property Let Opcao(ByVal vOpcao As String)
    pOpcao = vOpcao
End Property

Public Property Get Idioma() As String
    Idioma = pIdioma
End Property

Public Property Let Idioma(ByVal vIdioma As String)
    pIdioma = vIdioma
End Property

Public Property Get Volume() As String
    Volume = pVolume
End Property

Public Property Let Volume(ByVal vVolume As String)
    pVolume = vVolume
End Property

Public Property Get PrcVendas() As String
    PrcVendas = pPrcVendas
End Property

Public Property Let PrcVendas(ByVal vPrcVendas As String)
    pPrcVendas = vPrcVendas
End Property

Public Property Get PrcTotal() As String
    PrcTotal = pPrcTotal
End Property

Public Property Let PrcTotal(ByVal vPrcTotal As String)
    pPrcTotal = vPrcTotal
End Property

''---------------
'' administração
''---------------

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCol = Nothing
End Sub

Public Sub add(ByVal rec As clsProjeto, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsProjeto)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsProjeto(ByVal i As Variant) As clsProjeto
    Set clsProjeto = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsProjeto
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

'Public Function getProjetos(banco As clsBancos) As clsProjeto
'Dim colObj As New clsProjeto
'Dim recObj As New clsProjeto
'
'Dim cnn As New ADODB.Connection
'Set cnn = OpenConnection(banco)
'
'Dim cd As ADODB.Command
'Set cd = New ADODB.Command
'
'Dim rst As ADODB.Recordset
'
'With cd
'    .ActiveConnection = cnn
'    .CommandText = "select * from qryProfissoes"
'    .CommandType = adCmdText
'    Set rst = .Execute
'End With
'
'Do While Not rst.EOF
'    Set recObj = New clsProjeto
'    With recObj
'        .ID = rst.Fields("codCategoria")
'        .Profissao = rst.Fields("Descricao")
'        colObj.all recObj
'    End With
'    rst.MoveNext
'Loop
'
'Set getProfissoes = colObj
'
'Set recObj = Nothing
'Set colObj = Nothing
'
'End Function
