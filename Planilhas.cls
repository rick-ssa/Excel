VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mycollection As New Collection
Private myVal As Planilha

'Adicionar objetos à minha coleção
Public Function Add(sheet As Worksheet, registros As registros) As Planilha
    Set myVal = New Planilha
    With myVal
        .sheet = sheet
        .registros = registros
    End With
    mycollection.Add myVal
    Set Add = myVal
    Set myVal = Nothing
End Function

Public Function GetPlanilha(indice As Integer) As Planilha
    Set GetPlanilha = mycollection.Item(indice)
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

