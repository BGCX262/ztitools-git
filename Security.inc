<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Security.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de seguranca
REM  Cria‡ao     : 10/29/98 1:16AM
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Calcula o digito padrao modulo 11 dos bancos
REM -------------------------------------------------------------------------
Public Function DVModulo11(ByVal lngValor)
  Dim intPosicao 
  Dim intMultiplicador 
  Dim intSoma 
  Dim intDigito
  
  intSoma = 0
  intMultiplicador = 2
  
  For intPosicao = Len(lngValor) To 1 Step -1
    intSoma = intSoma + Int(Mid(lngValor, intPosicao, 1)) * intMultiplicador
    intMultiplicador = intMultiplicador + 1
    
    If intMultiplicador = 10 Then
      intMultiplicador = 2
      
    End If
  Next
  
  intDigito = 11 - (intSoma Mod 11)
  
  If intDigito < 1 Or intDigito > 9 Then
    intDigito = 1
    
  End If

  DVModulo11 = intDigito
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function DVModulo11

REM -------------------------------------------------------------------------
REM Calcula o Modulo 11 com o X padrao Banco do Brasil
REM -------------------------------------------------------------------------
Public Function DVModulo11X(ByVal lngValor)
  Dim intPosicao 
  Dim intMultiplicador 
  Dim intSoma 
  Dim intDigito
  
  intSoma = 0
  intMultiplicador = 2
  
  For intPosicao = Len(lngValor) To 1 Step -1
    intSoma = intSoma + Int(Mid(lngValor, intPosicao, 1)) * intMultiplicador
    intMultiplicador = intMultiplicador + 1
    
    If intMultiplicador = 10 Then
      intMultiplicador = 2
      
    End If
  Next
  
  intDigito = 11 - (intSoma Mod 11)
  
  If intDigito < 1 Or intDigito > 9 Then
    intDigito = "X"
    
  End If
  
  DVModulo11X = intDigito
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function DVModulo11X

REM -------------------------------------------------------------------------
REM Calculo do Digito Verificador Modulo 10 padrao dos Bancos
REM -------------------------------------------------------------------------
Public Function DVModulo10(ByVal lngValor)
  Dim intPosicao
  Dim intMultiplicador 
  Dim intProduto
  Dim intSoma 
  Dim intDigito
  
  intSoma = 0
  intMultiplicador = 2
  
  For intPosicao = Len(lngValor) To 1 Step -1
    intProduto = Int(Mid(lngValor, intPosicao, 1)) * intMultiplicador

    intSoma    = intSoma + (intProduto \ 10) + (intProduto Mod 10)

    intMultiplicador = intMultiplicador - 1

    If intMultiplicador = 0 Then
      intMultiplicador = 2

    End If
  Next
  
  intDigito = 10 - (intSoma Mod 10)

  If intDigito = 10 Then
    intDigito = "0"
      
  End If
  
  DVModulo10 = intDigito

End Function
REM -------------------------------------------------------------------------
REM Final da Function DVModulo10

Dim sblnIsUncriptNumber

sblnIsUncriptNumber = False

Const conEncript   = "A6EhHJbKwNqOzRrS3Tl5ZacUdDefgIi8jk7mXnBoGpYstLuQv0xFyV1W2P4M9C"
Const conMaxLenght = 8
REM -------------------------------------------------------------------------
REM Criptografa um conjunto de caracteres com o modulo 256 ciclico
REM -------------------------------------------------------------------------
Public Function EncriptPassword(ByVal strText)
  Dim i
  Dim intNext
  
  strText = UCase(strText)
  
  intNext = Asc(Right(strText, 1))
  
  For i = 1 To Len(strText)
    intNext = ((Asc(Mid(strText, i, 1)) * intNext) Mod 61) + 1
    
    EncriptPassword = EncriptPassword & Mid(conEncript, intNext, 1)
    
  Next 
  
End Function
REM -------------------------------------------------------------------------
REM End Function EncriptPassword

REM -------------------------------------------------------------------------
REM Criptografa um conjunto de caracteres com o modulo 256 ciclico
REM -------------------------------------------------------------------------
Public Function EncriptNumber(ByVal strNumber)
  Dim intLen
  Dim dblNumber
  Dim strPacket

  Do While strNumber > ""  
    dblNumber = Left(strNumber, conMaxLenght)
    
    dblNumber = CDbl(DVModulo11(dblNumber) & dblNumber)
    
    strPacket = ""
  
    Do While dblNumber > 61
      strPacket = strPacket & Mid(conEncript, (dblNumber Mod 61) + 1, 1)
      
      dblNumber = Fix(dblNumber / 61)

    Loop
  
    EncriptNumber = strPacket & Mid(conEncript, dblNumber + 1, 1) & EncriptNumber
    
    intLen = Len(strNumber)
    
    If intLen > conMaxLenght Then
      strNumber = Mid(strNumber, conMaxLenght + 1)

    Else
      strNumber = ""
      
    End If
    
  Loop

End Function
REM -------------------------------------------------------------------------
REM End Function EncriptNumber

REM -------------------------------------------------------------------------
REM Criptografa um conjunto de caracteres com o modulo 256 ciclico
REM -------------------------------------------------------------------------
Public Function UncriptNumber(ByVal strNumber)
  Dim i
  Dim lngPacket
  Dim intError
  
  intError = 0
  
  lngPacket = Instr(conEncript, Right(strNumber, 1)) - 1

  For i = Len(strNumber) - 1 To 1 Step -1
    If Len(lngPacket) >= conMaxLenght + 1 Then
      If Not CInt(Left(lngPacket, 1)) = CInt(DVModulo11(Mid(lngPacket, 2))) Then
        intError = intError + 1
        
      End If

      UncriptNumber = UncriptNumber & Mid(lngPacket, 2)
      
      lngPacket = Instr(conEncript, Mid(strNumber, i, 1)) - 1

    Else
      lngPacket = lngPacket * 61 + Instr(conEncript, Mid(strNumber, i, 1)) - 1

    End If
    
  Next
  
  If Not CInt(Left(lngPacket, 1)) = CInt(DVModulo11(Mid(lngPacket, 2))) Then
    intError = intError + 1
        
  End If

  UncriptNumber = UncriptNumber & Mid(lngPacket, 2)
  
  sblnIsUncriptNumber = True
  
  If intError > 0 Then  
    sblnIsUncriptNumber = False
    
  End If
  
End Function
REM -------------------------------------------------------------------------
REM End Function UncriptNumber

REM -------------------------------------------------------------------------
REM Retorna se houve erro no processamento da decriptacao do numero
REM -------------------------------------------------------------------------
Function IsUncriptNumberOK

  IsUncriptNumberOK = sblnIsUncriptNumber
  
End Function
REM -------------------------------------------------------------------------
REM End Function UncriptNumber

REM -------------------------------------------------------------------------
REM Fim do Security.inc
%>