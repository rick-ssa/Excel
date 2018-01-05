VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variáveis de Propriedades
Private mlinha As Integer
Private mindice As Integer
Private mcolunas As colunas
Attribute mcolunas.VB_VarHelpID = -1
Private mexcluido As Boolean

'Proriedades
'início proriedade linha====================================================================
Public Property Let linha(value As Integer)
    mlinha = value
End Property

Public Property Get linha() As Integer
    linha = mlinha
End Property
'fim propriedade linha====================================================================

'início proriedade indice====================================================================
Public Property Let indice(value As Integer)
    mindice = value
End Property

Public Property Get indice() As Integer
    indice = mindice
End Property
'fim propriedade indice====================================================================

'início proriedade colunas====================================================================
Public Property Get colunas() As colunas
    Set colunas = mcolunas
End Property
'fim propriedade colunas====================================================================

'início proriedade excluido====================================================================
Public Property Let excluido(value As Boolean)
    mexcluido = value
End Property

Public Property Get excluido() As Boolean
    excluido = mexcluido
End Property
'fim propriedade excluido====================================================================

Private Sub Class_Initialize()
    Set mcolunas = New colunas
    Set mcolunas.parent = Me
End Sub

Public Function getColunasAlteradas() As colunas
    Dim cols As New colunas
    Dim hex As String
    Dim indiceColuna As Integer
    Dim x As Integer
    Dim y As Integer
    Dim char As String
    
    y = 0
    
    indiceColuna = mcolunas.Count
    hex = mcolunas.alteracoesHex
    If hex <> "0" Then
        x = VBA.Len(hex)
        Do While x > 0
            char = VBA.Mid(hex, x, 1)
            If char <> "0" Then
                For y = 0 To 3
                    If CInt(CInt("&h" & char) And (2 ^ y)) <> 0 Then
                        cols.addColuna mcolunas.GetColuna(indiceColuna)
                    End If
                    indiceColuna = indiceColuna - 1
                Next y
                y = 0
            End If
            x = x - 1
        Loop
        Set getColunasAlteradas = cols
        Exit Function
    End If
    
    
    Set getColunasAlteradas = Nothing
End Function
