VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Colunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mycollection As New Collection
Private myVal As coluna
Private mparent As Registro

'Adicionar objetos à minha coleção
Public Function Add(indice As Integer, tipoDados As clTipoDados) As coluna
    Set myVal = New coluna
    With myVal
        .indice = indice
        .tipoDados = tipoDados
        .parent = Me
    End With
    mycollection.Add myVal
    Set Add = myVal
    Set myVal = Nothing
    
End Function

Public Function addColuna(col As coluna)
    mycollection.Add col
End Function

Public Function GetColuna(indice As Integer) As coluna
    Set GetColuna = mycollection.Item(indice)
End Function

Public Function Count() As Integer
    Count = mycollection.Count
End Function

Public Sub Remove(indice As Integer)
    mycollection.Remove indice
End Sub

Public Sub Clear()
    Dim Total As Integer
    Dim x As Integer
    Total = mycollection.Count
    For x = 1 To Total
        mycollection.Remove 1
    Next
End Sub

Public Function alteracoesHex() As String
  
    Dim y As Integer
    Dim n As Integer
    Dim e As Integer
    Dim x As Integer
    Dim result As String
    
    result = ""
    
    y = mycollection.Count
    
    Do While y >= 4
        e = 0
        x = y
        n = 0
        Do While x >= (y - 3)
            n = n + mycollection.Item(x).colAlterada * 2 ^ e
            e = e + 1
            x = x - 1
        Loop
        y = y - 4
        result = getHex(VBA.Abs(n)) + result
    Loop
    
    e = 0
    n = 0
    Do While y >= 1
        n = n + mycollection.Item(y).colAlterada * 2 ^ e
        e = e + 1
        y = y - 1
    Loop
    
    If n <> 0 Then
        result = getHex(VBA.Abs(n)) + result
    End If
    
    alteracoesHex = retireZerosDireita(result)
End Function

Public Function getHex(n As Integer) As String
    Dim r As String
    Select Case n
        
        Case 10
            r = "A"
        Case 11
            r = "B"
        Case 12
            r = "C"
        Case 13
            r = "D"
        Case 14
            r = "E"
        Case 15
            r = "F"
        Case Else
            r = CStr(n)
    End Select
    getHex = r
End Function

Private Function retireZerosDireita(hex As String)
    Dim x As Integer
    x = 0
    Do While x < VBA.Len(hex)
        x = x + 1
        If VBA.Mid(hex, x, 1) <> "0" Then
            retireZerosDireita = VBA.Mid(hex, x)
            Exit Function
        End If
    Loop
    
    retireZerosDireita = "0"
End Function

Public Property Set parent(value As Registro)
    Set mparent = value
End Property

Public Property Get parent() As Registro
    Set parent = mparent
End Property
