VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Vari�veis de Propriedades
Private msheet As Worksheet
Private mregistros As registros

'Proriedades
'in�cio proriedade sheet====================================================================
Public Property Let sheet(value As Worksheet)
    msheet = value
End Property

Public Property Get sheet() As Worksheet
    sheet = msheet
End Property
'fim propriedade sheet====================================================================

'in�cio proriedade registros====================================================================
Public Property Let registros(value As registros)
    mregistros = value
End Property

Public Property Get registros() As registros
    registros = mregistros
End Property
'fim propriedade registros====================================================================



