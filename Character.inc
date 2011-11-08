<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Charcter.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Caracteres
REM  Cria‡ao     : 11:23h 7/2/1998
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

Const conFontDirectory = "/ZTITools/Img/GraphChar/"

REM -------------------------------------------------------------------------
REM E uma Tag de final de linha
REM -------------------------------------------------------------------------
Public Function IsValidChar(ByVal strChar)
  Dim Result

  Result = False

  If InStr("abcdefghijklmnopqrstuvwxyz", strChar) Then
    Result = True

  ElseIf InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", strChar) Then
    Result = True

  ElseIf InStr("0123456789", strChar) Then
    Result = True

  ElseIf InStr("áéíóúÁÉÍÓÚýÝ", strChar) Then
    Result = True

  ElseIf InStr("àèìòùÀÈÌÒÙ", strChar) Then
    Result = True

  ElseIf InStr("äëïöüÄËÏÖÜ", strChar) Then
    Result = True

  ElseIf InStr("âêîôûÂÊÎÔÛ", strChar) Then
    Result = True

  ElseIf InStr("ãõÃÕ", strChar) Then
    Result = True

  ElseIf InStr("Çç", strChar) Then
    Result = True

  End If

  IsValidChar = Result

End Function
REM -------------------------------------------------------------------------
REM Final da Function isValidChar

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Function isCapital(ByVal strChar)
  Dim blnResult
  
  blnResult = False
  
  If strChar > "" Then
    strChar = Left(strChar, 1)
    
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÝÀÈÌÒÙÄËÏÖÜÂÊÎÔÛÃÕÇ", strChar) Then
      blnResult = True
      
    End If
  End If
  
  isCapital = blnResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function isCapital

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Function isSmall(ByVal strChar)
  Dim blnResult
  
  blnResult = False
  
  If strChar > "" Then
    strChar = Left(strChar, 1)
    
    If InStr("abcdefghijklmnopqrstuvwxyzáéíóúýàèìòùäëïöüâêîôûãõç", strChar) Then
      blnResult = True
      
    End If
  End If
  
  isSmall = blnResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function isSmall

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Function isNumber(ByVal strChar)
  Dim blnResult
  
  blnResult = False
  
  If strChar > "" Then
    strChar = Left(strChar, 1)
    
    If InStr("0123456789", strChar) Then
      blnResult = True
      
    End If
  End If
  
  isSmall = blnResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function isNumber

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Sub ShowGraphString(ByVal strString, ByVal strFont)
  Dim strChar
  Dim i

 	If strString > "" Then
    For i = 1 to Len(strString)
    	strChar = Mid(strString, i, 1)
    	
      ShowGraphChar strChar, strFont
    Next
  End If

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowGraphFont

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Sub ShowGraphChar(ByVal strChar, ByVal strFont)
  Dim strFontDir, strFontFull, strExt, strCaps, strCharResult
  
  If strChar > "" Then        
    strExt = GetGraphCharExt(strFont)
    strCaps = ""
  
    If Len(strChar) > 1 Then
      Select Case LCase(strChar)
        Case "left"
        Case "right"
      Case Else
        strChar = ""
        
      End Select
    Else      	
    	If isCapital(strChar) Then
      	strCaps = "C"
      	
      ElseIf isSmall(strChar) Then
      	strCaps = "S"
      	
      End If

      Select Case strChar
        Case "?" ' question mark
          strCharResult = "quest"

        Case "*" ' asterix sign
          strCharResult = "ast"

        Case "/" ' forward slash
          strCharResult = "fslash"

        Case "\" ' backward slash
          strCharResult = "bslash"

        Case "#" ' Number sign
          strCharResult = "Number"

        Case ":" ' Number sign
          strCharResult = "colon"

        Case Else
          If strChar = " " Then
            strCharResult = "Space"
          Else
            strCharResult = ConvAscii2Latin1(strChar)
      
          End If
        End Select
  
      If strCharResult > "" Then
        strChar = strCharResult
        
      End If
  
    End If
    
    strFontDir = conFontDirectory & strFont & "/"
    strFontFull = strFontDir & strFont & strCaps & strChar & strExt
  
    ShowHTML "<IMG SRC=""" & strFontFull & """ ALT=""" & strChar & """>"
    
  End If
End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowGraphChar

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Function GetGraphCharExt(ByVal strFont)
  Dim strExt
  
    If IsFileExist(Server.MapPath(conFontDirectory) & "\" & strFont & "\" & strFont & "Space.gif") Then
      strExt = ".gif"
      
    ElseIf IsFileExist(Server.MapPath(conFontDirectory) & "\" & strFont & "\" & strFont & "Space.jpg") Then
      strExt = ".jpg"
      
    Else
      strExt = ""
      
    End If

  GetGraphCharExt = strExt
  
End Function 
REM -------------------------------------------------------------------------
REM Fim da Function GetGraphCharExt

REM -------------------------------------------------------------------------
REM  Mostra os caracteres graficos
REM -------------------------------------------------------------------------
Public Sub ShowButton(ByVal strString, ByVal strFont)

  ShowGraphChar "Left", strFont
  ShowGraphString strString, strFont
  ShowGraphChar "Right", strFont
  
End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowButton

REM -------------------------------------------------------------------------
REM  Mostra o botao do tipo cilindrico
REM -------------------------------------------------------------------------
Public Sub ShowButtonCil(ByVal strString)
  ShowButton strString, "Cil"
  
End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowButtonCil

REM -------------------------------------------------------------------------
REM  Mostra o botao do tipo cilindrico
REM -------------------------------------------------------------------------
Public Sub ShowButtonWin(ByVal strString)
  ShowButton strString, "Win"
  
End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowButtonWin

REM -------------------------------------------------------------------------
REM  ShowNumbers(nNumbers, cDigit)
REM -------------------------------------------------------------------------
REM  Entradas    : nNumbers = intero a ser mostrado
REM                (123456789)
REM                cDigit = Tipo do digitos
REM                ("Green")
REM  Saidas      : 
REM -------------------------------------------------------------------------
Sub ShowNumbers(intNumbers, strFont)

 	If intNumbers = "" Then
    intNumbers = 0

  End If
	
	If UCase(strFont) = "ABACUS" Then
    ShowGraphChar "Left", strFont
		
	End If
	
  ShowGraphString intNumbers, strFont

	If UCase(strFont) = "ABACUS" Then
    ShowGraphChar "Left", strFont
		
	End If

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowNumbers

REM -------------------------------------------------------------------------
REM Fim do Character.inc
%>