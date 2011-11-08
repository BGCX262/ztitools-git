<%
REM -------------------------------------------------------------------------
REM  /ZTITools/String.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de String
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

Dim sstrPieceString
Dim sintPieceStringPosition
Dim sstrDateSeparator

  sstrDateSeparator = "/"

  sstrPieceString = ""
  sintPieceStringPosition = 1

REM -------------------------------------------------------------------------
REM Obtem a proxima string
REM -------------------------------------------------------------------------
Public Sub SetPiece(ByVal strString)
  Dim intPos

  Do
    intPos = InStr(strString, Chr(13))
    
    If intPos > 0 Then
      strString = Left(strString, intPos - 1) & Mid(strString, intPos + 1)
      
    End If
  Loop While intPos > 0

	sstrPieceString = strString

  sintPieceStringPosition = 1
	
End Sub

REM -------------------------------------------------------------------------
REM Obtem a proxima string
REM -------------------------------------------------------------------------
Public Sub ShowPiece()
	ShowMessage sstrPieceString
	
End Sub

REM -------------------------------------------------------------------------
REM Obtem a string inteira
REM -------------------------------------------------------------------------
Public Function GetEntirePiece()
	GetEntirePiece = sstrPieceString
	
End Function

REM -------------------------------------------------------------------------
REM Obtem a proxima string
REM -------------------------------------------------------------------------
Public Function GetPiece(ByVal intSize)
	GetPiece = Trim(Mid(sstrPieceString, sintPieceStringPosition, intSize))
	
	SkipPiece intSize

End Function

REM -------------------------------------------------------------------------
REM Obtem a proxima string
REM -------------------------------------------------------------------------
Public Function GetPieceRest()
	GetPieceRest = Trim(Mid(sstrPieceString, sintPieceStringPosition))
	
End Function

REM -------------------------------------------------------------------------
REM Obtem a proxima string colocando-a entre delimitadores de texto
REM -------------------------------------------------------------------------
Public Function GetPieceAsStr(ByVal intSize)
	GetPieceAsStr = "'" & GetPiece(intSize) & "'"
	
End Function

REM -------------------------------------------------------------------------
REM Obtem a proxima string formatando-a como data
REM -------------------------------------------------------------------------
Public Function GetPieceAsDate()
	GetPieceAsDate = chr(34) & GetPiece(8) & chr(34)
	
End Function

REM -------------------------------------------------------------------------
REM Salta a proxima string
REM -------------------------------------------------------------------------
Public Sub SkipPiece(ByVal intSize)

	sintPieceStringPosition = sintPieceStringPosition + intSize

End Sub

REM -------------------------------------------------------------------------
REM Obtem n-esima palavra de uma string
REM -------------------------------------------------------------------------
Public Function GetWord(strPhrase, n)
	Dim blnOnWord
	Dim nWordCount
	Dim strWord
	Dim nPos, nLength
	Dim cOneChar, strTempPhrase
	
	blnOnWord = false
	nWordCount = 0
	nPos = 1
	strTempPhrase = UCase(Trim(strPhrase))
	nLength = Len(strTempPhrase)
	
	do while nPos <= nLength
		cOneChar = mid(strTempPhrase, nPos, 1)
		if (cOneChar >= "A" and cOneChar <= "Z") or (cOneChar >= "0" and cOneChar <= "9") then
			if blnOnWord then
				strWord = strWord & cOneChar
			else
				blnOnWord = true
				nWordCount = nWordCount + 1
				strWord = cOneChar
			end if
		else
			if blnOnWord then
				blnOnWord = false
				if nWordCount = n then
					Exit Do					
				end if
			end if
		end if
		nPos = nPos + 1
	loop
	
	if nWordCount <> n then
		strWord = ""
	end if
	
	GetWord = strWord
end function

REM -------------------------------------------------------------------------
REM Checa as Strings e coloca os " ' " para a chamada SQL
REM -------------------------------------------------------------------------
Public Function CheckString(cString, cEndchar)
	nPos = InStr(cString, "'")

	While nPos > 0
		cString = Mid(cString, 1, nPos) & "'" & Mid(cString, nPos + 1)

		nPos = InStr(nPos + 2, cString, "'")

	Wend

	CheckString = "'" & cString & "'" & cEndchar & " "

End Function

REM -------------------------------------------------------------------------
REM Prepara valores numericos para serem adicionados a uma chamada SQL INSERT
REM -------------------------------------------------------------------------
Public Function CheckNumber(nValue, cSeparator)
	CheckNumber = nValue & cSeparator

End Function

REM -------------------------------------------------------------------------
REM Prepara Datas do Sistema para serem adicionados a uma chamada SQL INSERT
REM -------------------------------------------------------------------------
Public Function CheckSystemDate(cDate, cSeparator)
	CheckSystemDate = CheckString(cDate, cSeparator)

End Function

REM -------------------------------------------------------------------------
REM Prepara datas AAMMDD para serem adicionadas a uma chamada SQL INSERT
REM -------------------------------------------------------------------------
Public Function CheckDate(strDate, strSeparator)
	Dim strDia, strMes, strAno

	strAno = mid(strDate, 1, 2)
	strMes = mid(strDate, 3, 2)
	strDia = mid(strDate, 5, 2)

	strDate = strMes & sstrDateSeparator & strDia & sstrDateSeparator & strAno

	if isDate(strDate) then
		strDate = "'" & strDate & "'"
	else
		strDate = "NULL"
	end if

	CheckDate = strDate & cSeparator
End Function

Const conDateTypeAAMMDD   = 0
Const conDateTypeAAAAMMDD = 1
Const conDateTypeMMDDAA   = 2
Const conDateTypeMMDDAAAA = 3
Const conDateTypeDDMMAA   = 4
Const conDateTypeDDMMAAAA = 5

REM -------------------------------------------------------------------------
REM Prepara datas para serem utilizadas dentro de uma Query SQL
REM -------------------------------------------------------------------------
Public Function ImportDate(ByVal strDate, intDateType)
	Dim strDia, strMes, strAno
	Dim intDia, intMes, intAno, intAnoSize

  strDate = NormalizeString(Trim(strDate))
  
  intAnoSize = 2
  
  Select Case intDateType
    Case conDateTypeAAMMDD
      intDia = 5
      intMes = 3
      intAno = 1
      
    Case conDateTypeAAAAMMDD
      intDia = 7
      intMes = 5
      intAno = 1
      intAnoSize = 4
      
    Case conDateTypeMMDDAA
      intDia = 3
      intMes = 1
      intAno = 5

    Case conDateTypeMMDDAAAA
      intDia = 3
      intMes = 1
      intAno = 5
      intAnoSize = 4
      
    Case conDateTypeDDMMAA
      intDia = 1
      intMes = 3
      intAno = 5

    Case conDateTypeDDMMAAAA
      intDia = 1
      intMes = 3
      intAno = 5
      intAnoSize = 4
     
    Case Else
  End Select

	strAno = Mid(strDate, intAno, intAnoSize)
	strMes = Mid(strDate, intMes, 2)
	strDia = Mid(strDate, intDia, 2)

	strDate = strMes & sstrDateSeparator & strDia & sstrDateSeparator & strAno
	
	If Not isDate(strDate) Then
		strDate = Null

  Else
  	strDate = CDate(strDate)
		
	End If

  	ImportDate = strDate
	
End Function

Const conMoneyType9999D99  = 0
Const conMoneyType9D999D99 = 1
Const conMoneyType9999     = 2


REM -------------------------------------------------------------------------
REM Prepara Money para serem utilizadas dentro de uma Query SQL
REM -------------------------------------------------------------------------
Public Function ImportMoney(monValue, intMoneyType)

  monValue = TotalNormalizeString(AllTrim(monValue))
  
  If Len(monValue) > 0 Then
    If Fix(CDbl(monValue)) > 0 Then    
      Select Case intMoneyType
        Case conMoneyType9999D99, conMoneyType9D999D99
          ImportMoney = Fix(Left(monValue, Len(monValue) - 2)) & "." & Right(monValue, 2)
    
        Case Else
          ImportMoney = monValue
    
      End Select
    Else
        ImportMoney = 0.00
    End If
    
  Else
    ImportMoney = 0.00
    
  End If


End Function

REM -------------------------------------------------------------------------
REM Prepara um Int para ser importado
REM -------------------------------------------------------------------------
Public Function ImportInt(intValue)

  intValue = NormalizeString(AllTrim(intValue))
  
  If Len(intValue) > 0 Then
    If IsNumeric(intValue) Then    
      ImportInt = Fix(intValue)
  
    Else
      ImportInt = 0
    
    End If
  Else
    ImportInt = 0
    
  End If


End Function

REM -------------------------------------------------------------------------
REM Prepara um Int para ser importado
REM -------------------------------------------------------------------------
Public Function ImportLong(lngValue)

  lngValue = NormalizeString(AllTrim(lngValue))
  
  If Len(lngValue) > 0 Then
    If IsNumeric(lngValue) Then    
      ImportLong = CLng(lngValue)
  
    Else
      ImportLong = 0
    
    End If
  Else
    ImportLong = 0
    
  End If


End Function


REM -------------------------------------------------------------------------
REM Prepara um Int para ser importado
REM -------------------------------------------------------------------------
Public Function ImportDouble(lngValue)

  lngValue = NormalizeString(AllTrim(lngValue))
  
  If Len(lngValue) > 0 Then
    If IsNumeric(lngValue) Then    
      ImportLong = CDbl(lngValue)
  
    Else
      ImportLong = 0
    
    End If
  Else
    ImportLong = 0
    
  End If


End Function

REM -------------------------------------------------------------------------
REM Format number with leading zeroes
REM -------------------------------------------------------------------------
Public Function LeadingZeroes(Number, Places)             
  Number = Trim(Number)     
  If Len(Number) > Places Then
    LeadingZeroes = Number
    
  Else  
    LeadingZeroes = String(Places - Len(Number), "0") & Number
    
  End If  

End Function

REM -------------------------------------------------------------------------
REM Apaga os zeros a esquerda
REM -------------------------------------------------------------------------
Public Function LTrimZeroes(strValue)
  
  strValue = LTrim(strValue)
                                 
  Do While Len(strValue) > 0
    If Left(strValue, 1) = "0" Then
      strValue = Right(strValue, Len(strValue) - 1)
      
    Else
    
      Exit Do
      
    End If
                            
  Loop                            

	LTrimZeroes = strValue

end Function

REM ---------------------------------------------------------------------
REM Devolve um valor monetario completo R$#.###,##
REM ---------------------------------------------------------------------
Public Function ZTIMoney(ByVal monValue)
	ZTIMoney = 	"R$ " & ZTINumber(monValue)
	
End Function

REM ---------------------------------------------------------------------
REM Devolve um valor monetario completo R$#.###,##
REM ---------------------------------------------------------------------
Public Function ZTINumber(ByVal monValue)
	Dim Centavos, Reais, Negative, Centena
	Dim Target

	Centavos = LeadingZeroes(Fix(100 * (monValue - Fix(monValue)) + 0.5), 2)
	Reais = Fix(monValue)
	Negative = monValue < 0
	Target = ""
	
	do while Reais > 0
		Centena = Reais mod 1000
		Reais = Reais \ 1000
		
		if Reais then
			Target = "." & LeadingZeroes(Centena, 3) & Target
			
		else
			Target = Trim(CStr(Centena)) & Target
		end if
	loop
	
	if Target = "" then
		Target = "0"
	end if
	
	Target = Target & "," & Centavos
	
	if Negative then
		Target = "(" & Target & ")"
		
	end if
	
	ZTINumber = Target
	
End Function

REM ---------------------------------------------------------------------
REM Devolve um valor numerico completo #.###
REM ---------------------------------------------------------------------
Public Function ZTIInt(ByVal intValue)
	Dim Reais, Negative, Centena
	Dim Target
	
	If Trim(intValue) > "" Then
  	Reais = Round(intValue)
  	Negative = intValue < 0
  	Target = ""
  	
  	do while Reais > 0
  		Centena = Reais mod 1000
  		Reais = Reais \ 1000
  		
  		if Reais then
  			Target = "." & LeadingZeroes(Centena, 3) & Target
  			
  		else
  			Target = Trim(CStr(Centena)) & Target
  		end if
  	loop
  	
  	if Target = "" then
  		Target = "0"
  	end if
  	
  	Target = Target
  	
  	if Negative then
  		Target = "(" & Target & ")"
  		
  	end if
	
	Else
	  Target = 0

	End If  
	ZTIInt = Target
	
End Function

REM ---------------------------------------------------------------------
REM Devolve a data e a hora completos DD/MM/YYYY HH:MM:SS
REM ---------------------------------------------------------------------
Public Function ZTILongDateTime(ByVal numTime)
	ZTILongDateTime = ZTILongDate(numTime) & " " & ZTILongTime(numTime) 
End Function

REM ---------------------------------------------------------------------
REM Devolve a data longa DD/MM/YYYY
REM ---------------------------------------------------------------------
Public Function ZTILongDate(ByVal numTime)
  If IsDate(numTime) Then
    ZTILongDate = LeadingZeroes(Day(numTime), 2) & sstrDateSeparator & _
                  LeadingZeroes(Month(numTime), 2) & sstrDateSeparator & _  	    
                  LeadingZeroes(Year(numTime), 4)
                  
  End If          
    
End Function

REM ---------------------------------------------------------------------
REM Devolve a data longa MM/DD/YYYY
REM ---------------------------------------------------------------------
Public Function ZTIUSADate(ByVal numTime)
  If IsDate(numTime) Then
    ZTIUSADate = LeadingZeroes(Month(numTime), 2) & sstrDateSeparator & _
                  LeadingZeroes(Day(numTime), 2) & sstrDateSeparator & _  	    
                  LeadingZeroes(Year(numTime), 4)
                  
  End If          
    
End Function

REM ---------------------------------------------------------------------
REM Devolve a hora longa 24H
REM ---------------------------------------------------------------------
Function ZTILongTime(numTime)
  If IsDate(numTime) Then
    ZTILongTime = LeadingZeroes(Hour(numTime), 2) & ":" & _
  	    					LeadingZeroes(Minute(numTime), 2) & ":" & _
  			  				LeadingZeroes(Second(numTime), 2)
  			  				
	End If
	
End Function

REM ---------------------------------------------------------------------
REM Devolve a hora curta 24H
REM ---------------------------------------------------------------------
Function ZTIShortTime(numTime)
  If IsDate(numTime) Then
    ZTIShortTime = LeadingZeroes(Hour(numTime), 2) & ":" & _
  	    				   LeadingZeroes(Minute(numTime), 2)
  			  				
	End If
	
End Function

REM ---------------------------------------------------------------------
REM Devolve o numero de segundos de uma hora
REM ---------------------------------------------------------------------
Function ZTITimeSeconds(numTime)

	ZTITimeSeconds = Hour(numTime) * 3600 + Minute(numTime) * 60 + Second(numTime)
	
End Function

REM ---------------------------------------------------------------------
REM Devolve o numero de segundos de uma hora
REM ---------------------------------------------------------------------
Function ZTISecondsTime(numSeconds)

	ZTISecondsTime = TimeSerial(Int(numSeconds / 3600), Int(numSeconds / 3600 / 60), numSeconds mod 3600)
	
End Function

REM -------------------------------------------------------------------------
REM Prepara datas AAMMDD para serem adicionadas a uma chamada SQL INSERT
REM -------------------------------------------------------------------------
Public Function SQLLongDate(ByVal datDate)
  Dim strYear, strMonth, strDay

  If VarType(datDate) = vbString Then
    datDate = NormalizeString(datDate)

    strYear  = Mid(datDate, 5, 4)
    strMonth = Mid(datDate, 1, 2)
    strDay   = Mid(datDate, 3, 2)

    If Len(datDate) = 6 Then
      strYear  = "19" & Mid(datDate, 5, 2)
      
    End If
    
    datDate = strMonth & sstrDateSeparator & strDay & sstrDateSeparator & strYear
       
  End If

	SQLLongDate = "'" & LeadingZeroes(Month(datDate), 2) & sstrDateSeparator & LeadingZeroes(Day(datDate), 2) & sstrDateSeparator & LeadingZeroes(Year(datDate), 4) & "'"
	
End Function

REM -------------------------------------------------------------------------
REM Prepara datas MMDDAA para serem adicionadas a uma chamada SQL INSERT - Brasil
REM -------------------------------------------------------------------------
Public Function SQLLongDateBR(ByVal datDate)
  Dim strAux
  Dim strYear, strMonth, strDay

  If VarType(datDate) = vbString Then
    datDate = NormalizeString(datDate)

    strYear  = Mid(datDate, 5, 4)
    strMonth = Mid(datDate, 3, 2)
    strDay   = Mid(datDate, 1, 2)

    If Len(datDate) = 6 Then
      strYear  = "19" & Mid(datDate, 5, 2)
      
    End If
    
    datDate = strMonth & sstrDateSeparator & strDay & sstrDateSeparator & strYear
    
  End If

	SQLLongDateBR =  SQLLongDate(datDate)
	
End Function

REM -------------------------------------------------------------------------
REM Prepara data e hora para serem adicionadas a uma chamada SQL
REM -------------------------------------------------------------------------
Public Function SQLDateTime(ByVal dtmDate)
	SQLDateTime = "'" & CStr(dtmDate) & "'"
	
End Function

REM -------------------------------------------------------------------------
REM Prepara um Int para ser importado
REM -------------------------------------------------------------------------
Public Function SQLLong(ByVal lngValue)

  lngValue = NormalizeString(AllTrim(lngValue))
  
  If Len(lngValue) > 0 Then
    If IsNumeric(lngValue) Then    
      lngValue = CLng(lngValue)
  
    Else
      lngValue = 0
    
    End If
  Else
    lngValue = 0
    
  End If
  
  SQLLong = lngValue
                    
End Function

REM -------------------------------------------------------------------------
REM Prepara um Int para ser importado
REM -------------------------------------------------------------------------
Public Function SQLString(strString)

  SQLString = "'" & AllTrim(strString) & "'"
                    
End Function

REM -------------------------------------------------------------------------
REM Normaliza String
REM -------------------------------------------------------------------------
Public Function NormalizeString(strValue)
	dim target, tempstr
	dim i
	dim cOneChar

  If strValue > "" Then	
  	target = ""
  	tempstr = UCase(Trim(strValue))
  
  	for i = 1 to len(tempstr)
  		cOneChar = mid(tempstr, i, 1)
  		if (cOneChar >= "A" and cOneChar <= "Z") or (cOneChar >= "0" and cOneChar <= "9") then
  			target = target & cOneChar
  		end if
  	next
  	
  	NormalizeString = target
  Else
   	NormalizeString = strValue
  	
  End If
  
end function

REM -------------------------------------------------------------------------
REM Normaliza String
REM -------------------------------------------------------------------------
Public Function TotalNormalizeString(strValue)
	dim target, tempstr
	dim i
	dim cOneChar
	
	target = ""
	tempstr = UCase(Trim(strValue))
  If tempstr > "" Then
  	for i = 1 to len(tempstr)
  		cOneChar = mid(tempstr, i, 1)
  		
  		if (cOneChar >= "0" and cOneChar <= "9") then
  			target = target & cOneChar
  		end if
  		
  	next
	
  End If

	TotalNormalizeString = target
end function

REM -------------------------------------------------------------------------
REM Formata strings com o padrao com digito verificador 999.999-X
REM -------------------------------------------------------------------------
Function ZTIFormatDigito(ByVal strString)
  Dim strResult, i, strRest
  
  SetPiece ZTIReverse(strString)
  
  strResult = GetPiece(1) & "-"
  
  For i = 1 To Len(strString) Step 3
    strResult = strResult & GetPiece(3) & "."
  
  Next
  
  strRest = GetPieceRest
  
  If strRest > "" Then
    strResult = strResult & strRest
    
  Else
    strResult = Left(strResult, Len(strResult) - 1)
    
  End If
  
  
  ZTIFormatDigito = ZTIReverse(strResult)
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTIFormatDigito

REM -------------------------------------------------------------------------
REM Formata strings com o padrao 999.999
REM -------------------------------------------------------------------------
Function ZTIFormatNumber(ByVal strString)
  Dim strResult, i, strRest
  
  SetPiece ZTIReverse(strString)
  
  strResult = ""
  
  For i = 1 To Len(strString) Step 3
    strResult = strResult & GetPiece(3) & "."
  
  Next
  
  strRest = GetPieceRest
  
  If strRest > "" Then
    strResult = strResult & strRest
    
  Else
    strResult = Left(strResult, Len(strResult) - 1)
    
  End If
  
  
  ZTIFormatNumber = ZTIReverse(strResult)
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTIFormatNumber

REM -------------------------------------------------------------------------
REM Formata strings com o padrao de CGC 999.999.999/9999-99
REM -------------------------------------------------------------------------
Function ZTIFormatCGC(ByVal strString)
  Dim strResult, i, strRest
  
  strResult = ""

  strString = TotalNormalizeString(strString)
  SetPiece ZTIReverse(strString)
 
  If strString > "" Then  
  
    strResult = GetPiece(2) & "-" & GetPiece(4) & "/"

    For i = 7 To Len(strString) Step 3
      strResult = strResult & GetPiece(3) & "."
  
    Next
  
  
    strRest = GetPieceRest
  
    If strRest > "" Then
      strResult = strResult & strRest
     
    Else
      strResult = Left(strResult, Len(strResult) - 1)
    
    End If
  
  End If
  
  
  ZTIFormatCGC = ZTIReverse(strResult)
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTIFormatCGC

REM -------------------------------------------------------------------------
REM Formata strings com o padrao de CPF 999.999.999-99
REM -------------------------------------------------------------------------
Function ZTIFormatCPF(ByVal strString)
  Dim strResult, i, strRest
  
  strResult = ""
  
  strString = TotalNormalizeString(strString)
  SetPiece ZTIReverse(strString)
 
  If strString > "" Then  
  
    strResult = GetPiece(2) & "-"

    For i = 3 To Len(strString) Step 3
      strResult = strResult & GetPiece(3) & "."
  
    Next  
  
    strRest = GetPieceRest
  
    If strRest > "" Then
      strResult = strResult & strRest
     
    Else
      strResult = Left(strResult, Len(strResult) - 1)
    
    End If
  
  End If
  
  
  ZTIFormatCPF = ZTIReverse(strResult)
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTIFormatCPF

REM -------------------------------------------------------------------------
REM Formata strings com o padrao de telefone
REM -------------------------------------------------------------------------
Function ZTIFormatFone(ByVal strString) 

  If Not IsNull(strString) Then
    strString = NormalizeString(strString)
    
    Select Case Len(strString)
      Case 7
        ZTIFormatFone = Left(strString, 3) & "-" & Right(strString, 4)
        
      Case 8
        ZTIFormatFone = Left(strString, 4) & "-" & Right(strString, 4)    
        
      Case 10
        ZTIFormatFone = "(" & Left(strString, 3) & ") " & Mid(strString, 4, 3) & "-" & Right(strString, 4)
  
      Case 11
        ZTIFormatFone = "(" & Left(strString, 3) & ") " & Mid(strString, 4, 4) & "-" & Right(strString, 4)
        
      Case Else
        ZTIFormatFone = strString
              
    End Select
  End If
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTIFormatFone

REM -------------------------------------------------------------------------
REM Troca o separador de data
REM -------------------------------------------------------------------------
Sub SetDateSeparator(ByVal strSeparator)
  sstrDateSeparator = strSeparator

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetDateSeparator

REM -------------------------------------------------------------------------
REM Inverte o conteudo de uma string
REM -------------------------------------------------------------------------
Function ZTIReverse(ByVal strString)
  Dim i 
  Dim strTarget
    
  strTarget = ""

  If strString > "" Then
    For i = Len(strString) To 1 Step -1
      strTarget = strTarget & Mid(strString, i, 1)
    Next
    
  End If
    
  ZTIReverse = strTarget

End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTIReverse

REM -------------------------------------------------------------------------
REM Retorna o Dia da Semana de uma Data
REM -------------------------------------------------------------------------
Private Function DayofWeek(dtmDate, intType)
Dim strAux
  Select Case WeekDay(dtmDate)
    Case 1 
      strAux = "Domingo"
  
    Case 2
      strAux = "Segunda"
  
    Case 3 
      strAux = "Terça"
  
    Case 4 
      strAux = "Quarta"
  
    Case 5 
      strAux = "Quinta"
  
    Case 6 
      strAux = "Sexta"
  
    Case 7 
      strAux = "Sábado"
  
  End Select
  
  If intType = 0 Then
    DayofWeek = strAux
      
  Else
    DayofWeek = Left(strAux, intType)
  
  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub DayofWeek

REM -------------------------------------------------------------------------
REM Retorna o Nome de um Mes
REM -------------------------------------------------------------------------
Private Function MonthName(intMes)

  Select Case intMes
    Case 1
      MonthName = "Janeiro"
  
    Case 2
      MonthName = "Fevereiro"
  
    Case 3
      MonthName = "Março"
  
    Case 4
      MonthName = "Abril"
  
    Case 5
      MonthName = "Maio"
  
    Case 6
      MonthName = "Junho"
  
    Case 7
      MonthName = "Julho"
  
    Case 8
      MonthName = "Agosto"
  
    Case 9
      MonthName = "Setembro"
  
    Case 10
      MonthName = "Outubro"
  
    Case 11
      MonthName = "Novembro"
  
    Case 12
      MonthName = "Dezembro" 

  End Select
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function MonthName

REM -------------------------------------------------------------------------
REM Texto em Negrito
REM -------------------------------------------------------------------------
Public Function BoldText(ByVal strText)
  BoldText = "<B>" & strText & "</B>"

End Function
REM -------------------------------------------------------------------------
REM Final da Function BoldText
                
REM -------------------------------------------------------------------------
REM Texto em Italico
REM -------------------------------------------------------------------------
Public Function ItalicText(ByVal strText)
  ItalicText = "<I>" & strText & "</I>"

End Function
REM -------------------------------------------------------------------------
REM Final da Function ItalicText
                
REM -------------------------------------------------------------------------
REM Texto em Sublinhado
REM -------------------------------------------------------------------------
Public Function UnderlineText(ByVal strText)
  UnderlineText = "<U>" & strText & "</U>"

End Function
REM -------------------------------------------------------------------------
REM Final da Function UnderlineText

REM -------------------------------------------------------------------------
REM Retorna o Tamanho de um Mes
REM -------------------------------------------------------------------------
Function MonthSize(intMonth, intYear)
  Select Case intMonth
    Case 1,3,5,7,8,10,12
      MonthSize = 31

    Case 4,6,9,11
      MonthSize = 30

    Case 2
      If intYear Mod 4 = 0 Then
        MonthSize = 29
      
      Else
        MonthSize = 28
        
      End If
                                                
  End Select    
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function MonthSize

REM -------------------------------------------------------------------------
REM Fim do String.inc
%>