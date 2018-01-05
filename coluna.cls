VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "coluna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variáveis de Propriedades
Private mindice As Integer
Private mtipoDados As String
Private mvalor As String
Private mparent As colunas
Private blnAlterada As Boolean
Enum clTipoDados
    Texto = 1
    real = 2
    inteiro = 3
    Data = 4
    moeda = 5
    porcentagem = 6
End Enum

'Proriedades

Public Property Let parent(value As colunas)
    Set mparent = value
End Property

Public Property Get parent() As colunas
    Set parent = mparent
End Property

'início proriedade indice====================================================================
Public Property Let indice(value As Integer)
    mindice = value
End Property

Public Property Get indice() As Integer
    indice = mindice
End Property
'fim propriedade indice====================================================================

'início proriedade tipoDados====================================================================
Public Property Let tipoDados(value As clTipoDados)
    mtipoDados = value
End Property

Public Property Get tipoDados() As clTipoDados
    tipoDados = mtipoDados
End Property
'fim propriedade tipoDados====================================================================

'início proriedade valor====================================================================
Public Property Let valor(value As String)
    mvalor = value
    blnAlterada = True
End Property

Public Property Get valor() As String
    valor = mvalor
End Property
'fim propriedade valor====================================================================

Property Get letraColuna() As String
    Dim mo As Integer
    Dim div As Integer
    Dim result As String

    If mindice > 0 Then
        mo = mindice Mod 26
        div = mindice \ 26
        If mo = 0 Then
            mo = 26
            div = div - 1
        End If
        
        If mindice < 27 Then
            result = Chr(64 + mo)
        Else
            result = Chr(64 + div) & Chr(64 + mo)
        End If
    End If
    
    letraColuna = result
End Property

Public Sub resetAlterada()
    blnAlterada = False
End Sub

Property Get colAlterada() As Boolean
    colAlterada = blnAlterada
End Property
