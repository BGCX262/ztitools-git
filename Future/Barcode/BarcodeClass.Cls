VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BarcodeClass"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit

Private ConversionTable(0 To 9) As Integer

Public Function barcodeITF2from5(value As String) As String
    Dim Target As String, _
        i As Integer, _
        pair As Integer, _
        bitset(2) As Integer
        
    Target = "("
    
    For i = 1 To Len(value) Step 2
        pair = _
            ConversionTable(Val(Mid$(value, i, 1))) Or _
            (ConversionTable(Val(Mid$(value, i + 1, 1))) \ 2)
        bitset(0) = pair \ 32: bitset(1) = pair And &H1F
        Target = Target + _
                Chr$(AdjustValue(bitset(0)) + Asc("A")) + _
                Chr$(AdjustValue(bitset(1)) + Asc("a"))
    Next
    
    Target = Target + ")"
    
    barcodeITF2from5 = Target
End Function

Private Function AdjustValue(valor As Integer) As Integer
    Dim NovoValor As Integer
    
    Select Case valor
        Case 0 To 20
            NovoValor = valor
        Case 21
            NovoValor = 31 'Erro!
        Case 22
            NovoValor = 21
        Case 23
            NovoValor = 31 'Erro!
        Case 24 To 28
            NovoValor = valor - 2
        Case Else
            NovoValor = 31 'Erro!
    End Select
    
    AdjustValue = NovoValor
End Function



Private Sub Class_Initialize()
    ConversionTable(0) = &H28
    ConversionTable(1) = &H202
    ConversionTable(2) = &H82
    ConversionTable(3) = &H280
    ConversionTable(4) = &H22
    ConversionTable(5) = &H220
    ConversionTable(6) = &HA0
    ConversionTable(7) = &HA
    ConversionTable(8) = &H208
    ConversionTable(9) = &H88
End Sub


