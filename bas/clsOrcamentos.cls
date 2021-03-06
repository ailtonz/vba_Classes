VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrcamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pOrcamentos As New Collection

Private pControle As String
Private pVendedor As String
Private pStatus As String
Private pOperator As String

Public Property Get Controle() As String
    Controle = pControle
End Property

Public Property Let Controle(ByVal vControle As String)
    pControle = vControle
End Property

Public Property Get Vendedor() As String
    Vendedor = pVendedor
End Property

Public Property Let Vendedor(ByVal vVendedor As String)
    pVendedor = vVendedor
End Property

Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(ByVal vStatus As String)
    pStatus = vStatus
End Property

Public Property Get Operator() As String
    Operator = pOperator
End Property

Public Property Let Operator(ByVal vOperator As String)
    pOperator = vOperator
End Property


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pOrcamentos.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pOrcamentos = New Collection
End Sub

Private Sub Class_Terminate()
    Set pOrcamentos = Nothing
End Sub

Public Sub add(ByVal rec As clsOrcamentos, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pOrcamentos.add rec, key, before, after
End Sub

Public Function count() As Integer
    count = pOrcamentos.count
End Function

Public Function remove(ByVal index As Integer)
    pOrcamentos.remove index
End Function

Public Function clsOrcamentos(ByVal index As Variant) As clsOrcamentos
    Set clsOrcamentos = pOrcamentos.Item(index)
End Function
