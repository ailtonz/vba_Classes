VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGrands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCol As New Collection

Private pID As String
Private pNome As String
Private pProfissao As String
Private pValorLiquido As Currency

''---------------
'' propriedades
''---------------

Public Property Get ID() As String
    ID = pID
End Property

Public Property Let ID(ByVal vID As String)
    pID = vID
End Property

Public Property Get Nome() As String
    Nome = pNome
End Property

Public Property Let Nome(ByVal vNome As String)
    pNome = vNome
End Property

Public Property Get Profissao() As String
    Profissao = pProfissao
End Property

Public Property Let Profissao(ByVal vProfissao As String)
    pProfissao = vProfissao
End Property

Public Property Get ValorLiquido() As Currency
    ValorLiquido = pValorLiquido
End Property

Public Property Let ValorLiquido(ByVal vValorLiquido As Currency)
    pValorLiquido = vValorLiquido
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

Public Sub add(ByVal rec As clsGrands, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsGrands)
    pCol.add rec, rec.ID
End Sub

Public Function count() As Integer
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsGrand(ByVal index As Variant) As clsGrands
    Set clsGrand = pCol.Item(index)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsGrands
    Set Item = pCol(i)
End Property

''---------------
'' FUNÇÕES
''---------------

Public Function Insert(banco As clsBancos, Orcamento As clsOrcamentos, Grand As clsGrands) As Boolean
On Error GoTo Err_Insert
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandNew"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
        .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)
    
        .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, Grand.Profissao)
        .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, Grand.Nome)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, Grand.ValorLiquido)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Insert = False
    Else
      'Return a good return code
      Insert = True
    End If
  
Exit_Insert:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Insert:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Insert()" & vbNewLine & Err.Description)
  Insert = False
  Resume Exit_Insert

End Function

Public Function Update(banco As clsBancos, Orcamento As clsOrcamentos, Grand As clsGrands) As Boolean
On Error GoTo Err_Update
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandUpd"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("@NM_CONTROLE", adVarChar, adParamInput, 50, Orcamento.Controle)
        .Parameters.Append .CreateParameter("@NM_VENDEDOR", adVarChar, adParamInput, 50, Orcamento.Vendedor)
    
        .Parameters.Append .CreateParameter("@NM_PROFISSAO", adVarChar, adParamInput, 50, Grand.Profissao)
        .Parameters.Append .CreateParameter("@NM_NOME", adVarChar, adParamInput, 50, Grand.Nome)
        .Parameters.Append .CreateParameter("@NM_VALOR", adVarChar, adParamInput, 50, Grand.ValorLiquido)
        
        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, Grand.ID)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Update = False
    Else
      'Return a good return code
      Update = True
    End If
  
Exit_Update:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Update:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Update()")
  Update = False
  Resume Exit_Update

End Function

Public Function Delete(banco As clsBancos, Orcamento As clsOrcamentos, Grand As clsGrands) As Boolean
On Error GoTo Err_Delete
Dim Connection As New ADODB.Connection
Set Connection = OpenConnection(banco)
Dim rst As ADODB.Recordset
Dim cd As ADODB.Command
Dim lRecordsAffected As Long

    Set cd = New ADODB.Command
    With cd
        .ActiveConnection = Connection
        .CommandText = "admGrandDel"
        .CommandType = adCmdStoredProc
    
        .Parameters.Append .CreateParameter("@NM_ID", adVarChar, adParamInput, 50, Grand.ID)
    
        Set rst = .Execute(lRecordsAffected)
    End With
    Connection.Close
    
    If lRecordsAffected = 0 Then
      Delete = False
    Else
      'Return a good return code
      Delete = True
    End If
  
Exit_Delete:
  'Clean up the connection to the database
  Set cd = Nothing
  Set rst = Nothing

  Exit Function

Err_Delete:
  Call MsgBox("Class: " & TypeName(Me) & " , Function: Delete()")
  Delete = False
  Resume Exit_Delete

End Function

Public Function getGrands(banco As clsBancos) As clsGrands
Dim colObj As New clsGrands
Dim recObj As New clsGrands

Dim cnn As New ADODB.Connection
Set cnn = OpenConnection(banco)

Dim cd As ADODB.Command
Set cd = New ADODB.Command

Dim rst As ADODB.Recordset

With cd
    .ActiveConnection = cnn
    .CommandText = "select * from qryGrand"
    .CommandType = adCmdText
    Set rst = .Execute
End With

Do While Not rst.EOF
    Set recObj = New clsGrands
    With recObj
        .ID = rst.Fields("codigo")
        .Nome = rst.Fields("nome")
        .Profissao = rst.Fields("Profissao")
        .ValorLiquido = rst.Fields("Valor_Liquido")
        colObj.all recObj
    End With
    rst.MoveNext
Loop

Set getGrands = colObj

Set recObj = Nothing
Set colObj = Nothing

End Function
