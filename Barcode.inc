<%
REM -------------------------------------------------------------------------
REM  /ZTITools/BarCode.INC
REM -------------------------------------------------------------------------
REM  Descricao   : Monta Codigo de Barra
REM  Cria‡ao     : 08/03/1998 6:39:00PM
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1997 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

Const conBB0   = "0.gif"
Const conBB1   = "1.gif"
Const conBW0   = "2.gif"
Const conBW1   = "3.gif"

Const conStart = "8.gif"
Const conStop  = "9.gif"

REM -------------------------------------------------------------------------
REM Mostra o Codigo 2 of 5 
REM -------------------------------------------------------------------------
Public Sub ShowI2o5(ByVal strValue)
  Dim intValueLen, intCount
  Dim strResult
  Dim strFirst, strSecond
  
  strResult = ""
  
  strValue = NormalizeString(strValue)
  
  intValueLen = Len(strValue) 
  
  intCount = 1
  
  
  Do
    strFirst  = Mid(strValue, intCount, 1)

    intCount = intCount + 1

    If intCount <= intValueLen Then
      strSecond  = Mid(strValue, intCount, 1)
      
      intCount = intCount + 1 
      
    Else
      strSecond = "0"
      
    End If

    strResult = strResult & MountBarCodeI2o5(strFirst, strSecond)
        
  Loop While intCount < intValueLen
  
  ShowHTMLCR MountImageI5o5("Start", "") & strResult & MountImageI5o5("Stop", "")
  

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowI2o5


REM -------------------------------------------------------------------------
REM Mostra o Codigo 2 of 5 
REM -------------------------------------------------------------------------
Private Function MountBarCodeI2o5(ByVal strFirst, ByVal strSecond)
  Dim strResult
  Dim i
  
  strFirst  = GetPacketI5o5(strFirst)
  strSecond = GetPacketI5o5(strSecond)
  
  For i = 1 To 5
    strResult = strResult & MountImageI5o5(Mid(strFirst, i, 1), "B")
    strResult = strResult & MountImageI5o5(Mid(strSecond, i, 1), "W")
    
  Next

  MountBarCodeI2o5 = strResult

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ShowI2o5

REM -------------------------------------------------------------------------
REM Monta a imagem
REM -------------------------------------------------------------------------
Private Function MountImageI5o5(ByVal strBit, ByVal strType)
  Dim strResult
  Dim strTest 
  
  strTest = strType & strBit
  
  Select Case strTest
    Case "B0"
      strResult = conBB0
      
    Case "B1"
      strResult = conBB1

    Case "W0"
      strResult = conBW0
      
    Case "W1"
      strResult = conBW1
      
    Case "Start"
      strResult = conStart

    Case "Stop"
      strResult = conStop
      
    Case Else
      
   End Select
    

  MountImageI5o5 = "<IMG SRC=" & strResult & ">"
  
  REM " ALT=" & strBit &

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ShowI2o5

REM -------------------------------------------------------------------------
REM Monta a imagem
REM -------------------------------------------------------------------------
Private Function GetPacketI5o5(ByVal strValue)
  
  
  Select Case strValue
    Case "1"       
      GetPacketI5o5 = "10001"

    Case "2"       
      GetPacketI5o5 = "01001"
     

    Case "3"       
      GetPacketI5o5 = "11000"
     

    Case "4"       
      GetPacketI5o5 = "00101"
     

    Case "5"       
      GetPacketI5o5 = "10100"
     

    Case "6"       
      GetPacketI5o5 = "01100"
     
    Case "7"       
      GetPacketI5o5 = "00011"
     
    Case "8"       
      GetPacketI5o5 = "10010"
     
    Case "9"       
      GetPacketI5o5 = "01010"
     
    Case "0"       
      GetPacketI5o5 = "00110"

    Case "Start"       
      GetPacketI5o5 = "0000"

    Case "Stop"       
      GetPacketI5o5 = "100"
     
     
    Case Else
      
   End Select
    

End Function
REM -------------------------------------------------------------------------
REM Final da Sub GetPacketI5o5

REM -------------------------------------------------------------------------
REM Monta a imagem
REM -------------------------------------------------------------------------
Public Function ShowBoletoBC(ByVal strBanco, ByVal lngValor, ByVal strAgencia, ByVal strConta, ByVal strNossoNumero,  ByVal strGuess)
  Dim strResult
  
  strResult = strBanco & "99"
  
  lngValor = lngValor * 100
  
  strResult = CStr(lngValor)
  

End Function
REM -------------------------------------------------------------------------
REM Final da Sub GetPacketI5o5

REM -------------------------------------------------------------------------
REM Fim do Import.asp
%>
