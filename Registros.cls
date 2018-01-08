VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Registros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mycollection As New Collection
Private myVal As Registro
Private sortedByIndex As Integer 'store the index that the registers is sorted by

'Adicionar objetos à minha coleção
Public Function Add(linha As Integer, indice As Integer) As Registro
    Set myVal = New Registro
    With myVal
        .linha = linha
        .indice = indice
    End With
    mycollection.Add myVal
    Set Add = myVal
    Set myVal = Nothing
End Function

Public Function Getregistro(indice As Integer) As Registro
    Set Getregistro = mycollection.Item(indice)
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

'function merge sort
Private Function sortAndMerge(indiceColuna As Integer, tipoDados As clTipoDados, inicio As Integer, fim As Integer) As Collection
    Dim meio As Integer 'store the midle index of the registers
    Dim iRight As Integer 'to interate at right half
    Dim iLeft As Integer 'to interate at left half
    Dim r1 As Registro 'to store the actual register of right half
    Dim r2 As Registro 'to store the actual register of left half
    Dim nColl As New Collection 'to store the merging of the two halves
    Dim rightHalf As New Collection 'the right half collection
    Dim leftHalf As New Collection 'the left half collection
    Dim v As Integer 'interate to the final merge
    
    meio = inicio + (fim - inicio) \ 2
    
    If inicio <> meio And inicio <> fim Then 'be sure the right half has more than one item
        Set rightHalf = sortAndMerge(indiceColuna, tipoDados, inicio, meio) 'recursive call for right half
        Set leftHalf = sortAndMerge(indiceColuna, tipoDados, meio + 1, fim) 'recursive call for left half
        
    ElseIf inicio = fim Then 'if the half has only one item
        nColl.Add mycollection.Item(fim) 'add the item to new collection
        Set sortAndMerge = nColl 'return the new collection
        Exit Function
        
    Else 'if the half has two itens, spread one for each collection
        rightHalf.Add mycollection.Item(inicio)
        leftHalf.Add mycollection.Item(fim)
    End If
    
    'set the interators
    iRight = 1
    iLeft = 1
    
    'merge part =========================================================
    Do While iRight <= rightHalf.Count And iLeft <= leftHalf.Count
        'store the registers
        Set r1 = rightHalf.Item(iRight)
        Set r2 = leftHalf.Item(iLeft)
        
        'if the interators are the last from each halves add the new collection and exit the while loop
        If iRight = rightHalf.Count And iLeft = leftHalf.Count Then
            If r1.colunas.GetColuna(indiceColuna).valor > r2.colunas.GetColuna(indiceColuna).valor Then
                nColl.Add r2
                nColl.Add r1
            Else
                nColl.Add r1
                nColl.Add r2
            End If
            Exit Do
        End If
        
        'compare the two columns of the two registers
        If r1.colunas.GetColuna(indiceColuna).valor > r2.colunas.GetColuna(indiceColuna).valor Then
            nColl.Add r2 'add the registers that has the lower value
            If iLeft = leftHalf.Count Then 'if it is the last register of that half
                'add the rest of the other half registrs to the new collection
                For v = iRight To rightHalf.Count
                    Set r1 = rightHalf.Item(iRight)
                    nColl.Add r1
                    iRight = iRight + 1
                Next v
                Exit Do 'exit the while loop
            End If
            iLeft = iLeft + 1 'interator
        Else
            nColl.Add r1 'add the registers that has the lower value
            If iRight = rightHalf.Count Then 'if it is the last register of that half
                'add the rest of the other half registrs to the new collection
                For v = iLeft To leftHalf.Count
                    Set r2 = leftHalf.Item(iLeft)
                    nColl.Add r2
                    iLeft = iLeft + 1
                Next v
                Exit Do 'exit the while loop
            End If
            iRight = iRight + 1 'interator
        End If
    Loop
    Set sortAndMerge = nColl 'return the new collection
End Function

'call the merge sort and assingned to mycollection
Public Sub sort(indiceColuna As Integer)
    Dim r As Registro
    Dim c As coluna
    
    If mycollection.Count > 0 And sortedByIndex <> indiceColuna Then
        Set r = mycollection.Item(1)
        Set c = r.colunas.GetColuna(indiceColuna)
        Set mycollection = sortAndMerge(indiceColuna, c.tipoDados, 1, mycollection.Count)
        sortedByIndex = indiceColuna 'update the sorted index
    End If
End Sub
Public Function cloneRegistro(r As Registro) As Registro
    Dim nR As New Registro
    Dim x As Integer
    Dim nc As coluna
    Dim c As coluna
    For x = 1 To r.colunas.Count
        Set c = r.colunas.GetColuna(x)
        Set nc = nR.colunas.Add(c.indice, c.tipoDados)
        nc.colAlterada = c.colAlterada
        Set nc.parent = c.parent
        nc.valor = c.valor
    Next x
    nR.excluido = r.excluido
    nR.indice = r.indice
    nR.linha = r.linha
End Function
