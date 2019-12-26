VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProposta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String

Private pArqNome As String
Private pArqCaminho As String

Private pControle As String
Private pCliente As String
Private pResponsavel As String
Private pProjeto As String
Private pJournal As String
Private pAutor As String
Private pPublisher As String

Private pGerenteNome As String
Private pGerenteTelefone As String
Private pGerenteCelular01 As String
Private pGerenteCelular02 As String
Private pGerenteIDnextel As String


''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get ArqCaminho() As String
    ArqCaminho = pArqCaminho
End Property

Public Property Let ArqCaminho(ByVal vArqCaminho As String)
    pArqCaminho = vArqCaminho
End Property

Public Property Get ArqNome() As String
    ArqNome = pArqNome
End Property

Public Property Let ArqNome(ByVal vArqNome As String)
    pArqNome = vArqNome
End Property

Public Property Get Controle() As String
    Controle = pControle
End Property

Public Property Let Controle(ByVal vControle As String)
    pControle = vControle
End Property

Public Property Get Cliente() As String
    Cliente = pCliente
End Property

Public Property Let Cliente(ByVal vCliente As String)
    pCliente = vCliente
End Property

Public Property Get Responsavel() As String
    Responsavel = pResponsavel
End Property

Public Property Let Responsavel(ByVal vResponsavel As String)
    pResponsavel = vResponsavel
End Property

Public Property Get Projeto() As String
    Projeto = pProjeto
End Property

Public Property Let Projeto(ByVal vProjeto As String)
    pProjeto = vProjeto
End Property

Public Property Get Journal() As String
    Journal = pJournal
End Property

Public Property Let Journal(ByVal vJournal As String)
    pJournal = vJournal
End Property

Public Property Get Autor() As String
    Autor = pAutor
End Property

Public Property Let Autor(ByVal vAutor As String)
    pAutor = vAutor
End Property

Public Property Get Publisher() As String
    Publisher = pPublisher
End Property

Public Property Let Publisher(ByVal vPublisher As String)
    pPublisher = vPublisher
End Property


Public Property Get GerenteNome() As String
    GerenteNome = pGerenteNome
End Property

Public Property Let GerenteNome(ByVal vGerenteNome As String)
    pGerenteNome = vGerenteNome
End Property


Public Property Get GerenteTelefone() As String
    GerenteTelefone = pGerenteTelefone
End Property

Public Property Let GerenteTelefone(ByVal vGerenteTelefone As String)
    pGerenteTelefone = vGerenteTelefone
End Property


Public Property Get GerenteCelular01() As String
    GerenteCelular01 = pGerenteCelular01
End Property

Public Property Let GerenteCelular01(ByVal vGerenteCelular01 As String)
    pGerenteCelular01 = vGerenteCelular01
End Property


Public Property Get GerenteCelular02() As String
    GerenteCelular02 = pGerenteCelular02
End Property

Public Property Let GerenteCelular02(ByVal vGerenteCelular02 As String)
    pGerenteCelular02 = vGerenteCelular02
End Property


Public Property Get GerenteIDnextel() As String
    GerenteIDnextel = pGerenteIDnextel
End Property

Public Property Let GerenteIDnextel(ByVal vGerenteIDnextel As String)
    pGerenteIDnextel = vGerenteIDnextel
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

Public Sub add(ByVal rec As clsProposta, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsProposta)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsProposta(ByVal i As Variant) As clsProposta
    Set clsProposta = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsProposta
    Set Item = pCol(i)
End Property


''---------------
'' FUNÇÕES
''---------------

Public Sub GerarProposta(ByVal obj As clsProposta, ByVal prj As clsProjeto)
Dim wdApp As Word.Application, wdDoc As Word.Document

Dim col As New clsProjeto

On Error Resume Next

Set wdApp = GetObject(, "Word.Application")

'Word isn't already running
If Err.Number <> 0 Then
    Set wdApp = CreateObject("Word.Application")
End If

On Error GoTo 0

Set wdDoc = wdApp.Documents.Open(obj.ArqCaminho & "\" & obj.ArqNome, ReadOnly:=True)

wdApp.Visible = True

With wdDoc
    '' PROPOSTA
    .Bookmarks("N_CONTROLE").Range.Text = obj.Controle
    .Bookmarks("CLIENTE").Range.Text = obj.Cliente
    .Bookmarks("RESPONSAVEL").Range.Text = obj.Responsavel
    .Bookmarks("PROJETO").Range.Text = obj.Projeto
    .Bookmarks("JOURNAL").Range.Text = obj.Journal
    .Bookmarks("AUTOR").Range.Text = obj.Autor
    .Bookmarks("PUBLISHER").Range.Text = obj.Publisher
    
    '' PROJETOS
    For Each prj In col.Itens

        .Bookmarks("FORMATO").Range.Text = prj.Opcao
        .Bookmarks("IDIOMA").Range.Text = prj.Idioma
        .Bookmarks("VOLUME").Range.Text = prj.Volume
        .Bookmarks("PRC_VENDA").Range.Text = prj.PrcVendas
        .Bookmarks("PRC_TOTAL").Range.Text = prj.PrcTotal

    Next prj
    
    '' GERENTE DE CONTAS
    .Bookmarks("G_CONTAS").Range.Text = obj.GerenteNome
    .Bookmarks("TELEFONE").Range.Text = obj.GerenteTelefone
    .Bookmarks("CELULAR_01").Range.Text = obj.GerenteCelular01
    .Bookmarks("CELULAR_02").Range.Text = obj.GerenteCelular02
    .Bookmarks("ID_NEXTEL").Range.Text = obj.GerenteIDnextel
End With

wdDoc.SaveAs pathDesktopAddress & "\" & Format(CDate(Now()), "dd-mm-yy_hh-mm") & "_" & obj.ArqNome
wdDoc.Close
wdApp.Application.Quit

End Sub

Private Function pathDesktopAddress() As String
    pathDesktopAddress = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
End Function

Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.add BookmarkToUpdate, BMRange
End Sub


'Public Function getProfissoes(banco As clsBancos) As clsProposta
'Dim colObj As New clsProposta
'Dim recObj As New clsProposta
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
'    Set recObj = New clsProposta
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



