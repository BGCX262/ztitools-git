<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Validate.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de validacoes JavaScript 
REM  Cria‡ao     : 08:30h 28/12/1998
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM              : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

Dim sblnPutJSStartEnd

sblnPutJSStartEnd = True

Private Sub JavaScriptBegin
  If sblnPutJSStartEnd Then
    ShowHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"

  End If  

End Sub  

Private Sub JavaScriptEnd
  If sblnPutJSStartEnd Then
    ShowHTMLCR "</SCRIPT>"

  End If  

End Sub  

Private Sub SetPutJSStartEnd(blnAux)
  sblnPutJSStartEnd = blnAux

End Sub  

REM -------------------------------------------------------------------------
REM Entra FormJavaValidSepDate
REM -------------------------------------------------------------------------
Private Sub FormJavaValidSepDate(ByVal strField, strField2, strField3, strForm, strFunctionName, blnMaiorQueHoje, strMessage)
                                         
  JavaScriptBegin
  
  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "var dataatual, data, text;"
  
  ShowHTMLCR "if ((document." + strForm + "." & strField & ".value > """")  && (document." + strForm + "." & strField2 & ".value > """") && (document." + strForm + "." & strField3 & ".value > """")){"           
  ShowHTMLCR "text = document." + strForm + "." & strField & ".value + document." + strForm + "." & strField2 & ".value + document." + strForm + "." & strField3 & ".value;"
  ShowHTMLCR "if (!isdate(text)){" 
  ShowHTMLCR "alert(""Data Inválida"");"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}}" 

  If blnMaiorQueHoje Then
    ShowHTMLCR "dataatual  = new Date();"
    ShowHTMLCR "data       = todate(text);"
    ShowHTMLCR "if (data.valueOf() + 90000000 < dataatual.valueOf()) {"
    ShowHTMLCR "alert(""" & strMessage & """);"
    ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
    ShowHTMLCR "return false;"
    ShowHTMLCR "}"
  
  End If
  
REM  If sstrLastValidate > "" Then
REM    ShowHTMLCR "return " & sstrLastValidate & "()"  
    
REM  End If  
  ShowHTMLCR "}"
    
  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaValidSepDate

REM -------------------------------------------------------------------------
REM Entra FormJavaValidDate
REM -------------------------------------------------------------------------
Private Sub FormJavaValidDate(ByVal strField, strForm, strFunctionName, blnMaiorQueHoje, strMessage)
                                         
  JavaScriptBegin
  
  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "var dataatual, data, text;"
  
  ShowHTMLCR "if (document." + strForm + "." & strField & ".value > """"){"           
  ShowHTMLCR "text = document." + strForm + "." & strField & ".value;"
  ShowHTMLCR "if (!isdate(text)){" 
  ShowHTMLCR "alert(""Data Inválida"");"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}" 
  ShowHTMLCR "else" 
  ShowHTMLCR "{" 
  ShowHTMLCR "document." + strForm + "." & strField & ".value = datetostring(todate(text))"
  ShowHTMLCR "return true}"
  ShowHTMLCR "}" 

  If blnMaiorQueHoje Then
    ShowHTMLCR "dataatual  = new Date();"
    ShowHTMLCR "data       = todate(text);"
    ShowHTMLCR "if (data.valueOf() + 90000000 < dataatual.valueOf()) {"
    ShowHTMLCR "alert(""" & strMessage & """);"
    ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
    ShowHTMLCR "return false;"
    ShowHTMLCR "}"
    ShowHTMLCR "else{return true}"
  
  End If
  
  ShowHTMLCR "}"
    
  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaValidDate

REM -------------------------------------------------------------------------
REM Valida um campo de CGC
REM -------------------------------------------------------------------------
Function FormJavaValidCGC(ByVal strField, strForm, strFunctionName)

  JavaScriptBegin

  ShowHTMLCR "function " & strFunctionName & "(){"

  ShowHTMLCR "var cgc, soma, digito, digito2, multip, i, aux, numero;"
  ShowHTMLCR "soma   = 0;"
  ShowHTMLCR "multip = 5;"

  ShowHTMLCR "cgc    = document." + strForm + "." & strField & ".value;"
  ShowHTMLCR "cgc    = leadingzeroes(normalize(cgc), 14);"
  ShowHTMLCR "for(i = 0;i <= 11;i++){"
  ShowHTMLCR "number = parseInt(cgc.charAt(i));"
  ShowHTMLCR "soma  += (number * multip);"
  ShowHTMLCR "multip+= - 1;"
  ShowHTMLCR "if (multip < 2)"
  ShowHTMLCR "multip = 9;}"
  ShowHTMLCR "aux    = soma % 11;"
  ShowHTMLCR "if ((aux == 0) || (aux == 1))"
  ShowHTMLCR "digito = 0;"
  ShowHTMLCR "else"
  ShowHTMLCR "digito = 11 - aux;"
  ShowHTMLCR "if (parseInt(cgc.charAt(12)) == digito){"
  ShowHTMLCR "multip = 6;"
  ShowHTMLCR "soma   = 0;"
  ShowHTMLCR "for(i = 0;i <= 12;i++){"
  ShowHTMLCR "numero = parseInt(cgc.charAt(i));"
  ShowHTMLCR "soma  += (numero * multip);"
  ShowHTMLCR "multip+= - 1;"
  ShowHTMLCR "if (multip < 2)"
  ShowHTMLCR "multip = 9;"
  ShowHTMLCR "}"
  ShowHTMLCR "aux = soma % 11;"
  ShowHTMLCR "if ((aux == 0) || (aux == 1))"
  ShowHTMLCR "digito = 0;"
  ShowHTMLCR "else"
  ShowHTMLCR "digito = 11 - aux;"
  ShowHTMLCR "if (parseInt(cgc.charAt(13)) == digito)"
  ShowHTMLCR "return true;"
  ShowHTMLCR "else{"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "alert(""Este CGC é inválido"");"
  ShowHTMLCR "return false;}}"
  ShowHTMLCR "else{"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "alert(""Este CGC é inválido"");"
  ShowHTMLCR "return false;}"
  ShowHTMLCR "}"
  JavaScriptEnd

End Function
REM ---------------------------------------------------------------------
REM Final da function FormJavaValidCGC

REM -------------------------------------------------------------------------
REM Valida um campo de CPF
REM -------------------------------------------------------------------------
Function FormJavaValidCPF(ByVal strField, strForm, strFunctionName)

  JavaScriptBegin

  ShowHTMLCR "function " & strFunctionName & "(){"

  ShowHTMLCR "var cpf, soma, digito, digito2, multip, i, aux, numero;"
  ShowHTMLCR "soma    = 0;"
  ShowHTMLCR "multip  = 10;"

  ShowHTMLCR "cpf     = document." + strForm + "." & strField & ".value;"
  ShowHTMLCR "cpf     = leadingzeroes(normalize(cpf), 11);"
  ShowHTMLCR "for(i = 0;i <= 8;i++){"
  ShowHTMLCR "number = parseInt(cpf.charAt(i));"
  ShowHTMLCR "soma  += (number * multip);"
  ShowHTMLCR "multip+= - 1;}"
  ShowHTMLCR "aux = soma % 11;"
  ShowHTMLCR "if ((aux == 0) || (aux == 1))"
  ShowHTMLCR "digito = 0;"
  ShowHTMLCR "else"
  ShowHTMLCR "digito = 11 - aux;"
  ShowHTMLCR "if (parseInt(cpf.charAt(9)) == digito){"
  ShowHTMLCR "multip = 11;"
  ShowHTMLCR "soma   = 0;"
  ShowHTMLCR "for(i = 0;i <= 9;i++){"
  ShowHTMLCR "numero = parseInt(cpf.charAt(i));"
  ShowHTMLCR "soma  += (numero * multip);"
  ShowHTMLCR "multip+= - 1;"
  ShowHTMLCR "}"
  ShowHTMLCR "aux = soma % 11;"
  ShowHTMLCR "if ((aux == 0) || (aux == 1))"
  ShowHTMLCR "digito = 0;"
  ShowHTMLCR "else"
  ShowHTMLCR "digito = 11 - aux;"
  ShowHTMLCR "if (parseInt(cpf.charAt(10)) == digito)"
  ShowHTMLCR "return true;"
  ShowHTMLCR "else{"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "alert(""Este CPF é inválido"");"
  ShowHTMLCR "return false;}}"
  ShowHTMLCR "else{"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "alert(""Este CPF é inválido"");"
  ShowHTMLCR "return false;}"
  ShowHTMLCR "}"
  JavaScriptEnd

End Function
REM ---------------------------------------------------------------------
REM Final da function FormJavaValidCPF

REM -------------------------------------------------------------------------
REM Valida um campo de e-mail
REM -------------------------------------------------------------------------
Function FormJavaValidEMail(ByVal strField, strForm, strFunctionName)
  JavaScriptBegin

  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "var email;"
  ShowHTMLCR "email = document." + strForm + "." & strField & ".value;"
  ShowHTMLCR "if ((email == """") || ((email.indexOf(""."") != -1) && (email.indexOf(""@"") != -1)))"
  ShowHTMLCR "return true;"
  ShowHTMLCR "else{"
  ShowHTMLCR "document." + strForm + "." & strField & ".focus();"
  ShowHTMLCR "alert(""Este e-mail é inválido"");"
  ShowHTMLCR "return false;}"
  ShowHTMLCR "}"
  JavaScriptEnd

End Function
REM ---------------------------------------------------------------------
REM Final da function FormJavaValidEMail

REM -------------------------------------------------------------------------
REM Cria a Classe StrExt
REM -------------------------------------------------------------------------
Private Sub StrExtBegin

  JavaScriptBegin

  ShowHTMLCR "function mid(text, start, length){return text.substr(start - 1, length);}"

  ShowHTMLCR "function left(text, length){return text.substr(0, length);}"

  ShowHTMLCR "function midtoend(text, start){return text.substr(start - 1);}"

  ShowHTMLCR "function normalize(text)"
  ShowHTMLCR "{var i, text2;"
  ShowHTMLCR "  text2 = text;"
  ShowHTMLCR "  i = 0;"
  ShowHTMLCR "  while (i < text2.length)"
  ShowHTMLCR "    if ((text2.substr(i, 1) < ""0"") || (text2.substr(i, 1) > ""9""))"
  ShowHTMLCR "      text2 = text2.substr(0, i) + midtoend(text2, i + 2);"
  ShowHTMLCR "    else"
  ShowHTMLCR "      i++;"
  ShowHTMLCR "  return text2;"
  ShowHTMLCR "}"

  ShowHTMLCR "function leadingzeroes(text, tam){"
  ShowHTMLCR "  while (text.length < tam)"
  ShowHTMLCR "    text = ""0"" + text;"
  ShowHTMLCR "  return text;"
  ShowHTMLCR "}"

  ShowHTMLCR "function isdate(text){"
  ShowHTMLCR "var Dia, Mes, Ano, dataatual;"

  ShowHTMLCR "  blnis = true;"  
  ShowHTMLCR "  dataatual = new Date();"  
  ShowHTMLCR "  text = normalize(text);"
  ShowHTMLCR "  Mes = dataatual.getMonth() + 1"
  ShowHTMLCR "  Ano = dataatual.getFullYear()"

  ShowHTMLCR "  if (text.length > 1){"
  ShowHTMLCR "    Dia = parseFloat(left(text, 2));"

  ShowHTMLCR "    if (text.length > 3){"
  ShowHTMLCR "      Mes = parseFloat(mid(text, 3,2));"
  ShowHTMLCR "    }"
  
  ShowHTMLCR "    if (text.length > 5){"
  ShowHTMLCR "      Ano = parseFloat(midtoend(text, 5));}"
  
  ShowHTMLCR "    if ((Mes > 12) || (Mes < 1))"
  ShowHTMLCR "      blnis = false;"

  ShowHTMLCR "    switch (Mes) {"
  ShowHTMLCR "      case 1:"  
  ShowHTMLCR "      case 3:"  
  ShowHTMLCR "      case 5:"  
  ShowHTMLCR "      case 7:"  
  ShowHTMLCR "      case 8:"  
  ShowHTMLCR "      case 10:"  
  ShowHTMLCR "      case 12:"  
  ShowHTMLCR "        if ((Dia > 31) || (Dia < 1))"
  ShowHTMLCR "          blnis = false;break;"
  
  ShowHTMLCR "      case 4:"  
  ShowHTMLCR "      case 6:"  
  ShowHTMLCR "      case 9:"  
  ShowHTMLCR "      case 11:"  
  ShowHTMLCR "        if ((Dia > 30) || (Dia < 1))"
  ShowHTMLCR "          blnis = false;break;"
  
  ShowHTMLCR "      case 2:"
  ShowHTMLCR "        if (Ano % 4 > 0)"
  ShowHTMLCR "          if ((Dia > 28) || (Dia < 1))"
  ShowHTMLCR "            blnis = false;"            
  
  ShowHTMLCR "        else"                            
  ShowHTMLCR "          if ((Dia > 29) || (Dia < 1))"
  ShowHTMLCR "            blnis = false;break"            
  
  ShowHTMLCR "     }"
  ShowHTMLCR "    if (Ano > dataatual.getFullYear() + 60){" 
  ShowHTMLCR "      blnis = false;}"

  ShowHTMLCR "   }"  
  ShowHTMLCR "  else"
  ShowHTMLCR "    blnis = false;"
  ShowHTMLCR "  return blnis;"
  ShowHTMLCR "}"

  ShowHTMLCR "function todate(text){"
  ShowHTMLCR "var Dia, Mes, Ano, data;"

  ShowHTMLCR "   data = new Date();"  
  ShowHTMLCR "   text = normalize(text);"
  ShowHTMLCR "   Mes = data.getMonth() + 1"
  ShowHTMLCR "   Ano = data.getFullYear()"

  ShowHTMLCR "   if (text.length > 1){"
  ShowHTMLCR "     Dia = parseFloat(left(text, 2));}"

  ShowHTMLCR "   if (text.length > 3){"
  ShowHTMLCR "     Mes = parseFloat(mid(text, 3,2));}"
  
  ShowHTMLCR "   if (text.length > 5){"
  ShowHTMLCR "     Ano = parseFloat(midtoend(text, 5));}"

  ShowHTMLCR "   if (Ano < 100){"
  ShowHTMLCR "     if (Ano + 1900 > data.getFullYear() - 70)"
  ShowHTMLCR "       Ano += 1900;"

  ShowHTMLCR "     else"
  ShowHTMLCR "       Ano += 2000;"
  
  ShowHTMLCR "     }"  
  ShowHTMLCR "   data = new Date(Ano, Mes - 1, Dia);"
  ShowHTMLCR "   return data"
  ShowHTMLCR "  }"

  ShowHTMLCR "function datetostring(data){ return leadingzeroes(data.getDate().toString(), 2) + ""/"" + leadingzeroes((data.getMonth() + 1).toString(), 2) + ""/"" + leadingzeroes(data.getFullYear(), 2);}"

  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub StrExtBegin

REM -------------------------------------------------------------------------
REM Entra FormJavaVerEscolha
REM -------------------------------------------------------------------------
Private Sub FormJavaVerEscolha(ByVal strField, strForm, strFunctionName)
  JavaScriptBegin

  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "if (document." & strForm & "." & strField & ".value == ""00""){" 
  ShowHTMLCR "alert(""Uma das opções deve ser escolhida"");"
  ShowHTMLCR "document." & strForm & "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}"
  ShowHTMLCR "return true;"
  ShowHTMLCR "}"

  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaVerEscolha

REM -------------------------------------------------------------------------
REM Entra FormJavaVerEscolha
REM -------------------------------------------------------------------------
Private Sub FormJavaDataEscolha(ByVal strField, strField2, strField3, strForm, strFunctionName, strMessage)

  FormJavaVerEscolha strField3, strForm, strFunctionName & "VerEscolha"
  FormJavaCompareDate strField, strField2, strForm, strFunctionName & "CompareData", strMessage
  
  JavaScriptBegin

  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "return (" & strFunctionName & "VerEscolha()&&" & strFunctionName & "CompareData())"
  ShowHTMLCR "}"

  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaVerEscolha

REM -------------------------------------------------------------------------
REM Entra FormJavaCompareDate
REM -------------------------------------------------------------------------
Private Sub FormJavaCompareDate(ByVal strField, strField2, strForm, strFunctionName, strMessage)
                                         
  JavaScriptBegin
  
  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "var datainicio, datafim, text, text2;"
  ShowHTMLCR "datainicio = new Date();"
  ShowHTMLCR "datafim    = new Date();"           
  ShowHTMLCR "text       = document." & strForm & "." & strField & ".value;"           
  ShowHTMLCR "text2      = document." & strForm & "." & strField2 & ".value;"           
  ShowHTMLCR "if (text > """")"  
  ShowHTMLCR "if (!isdate(text)){" 
  ShowHTMLCR "alert(""Data Inválida"");"
  ShowHTMLCR "document." & strForm & "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}" 
  ShowHTMLCR "else{" 
  ShowHTMLCR "datainicio = todate(text);" 
  ShowHTMLCR "document." & strForm & "." & strField & ".value = datetostring(datainicio);"
  ShowHTMLCR "}"
  ShowHTMLCR "if (text2 > """")"
  ShowHTMLCR "if (!isdate(text2)){" 
  ShowHTMLCR "alert(""Data Inválida"");"
  ShowHTMLCR "document." & strForm & "." & strField2 & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}"
  ShowHTMLCR "else{" 
  ShowHTMLCR "datafim = todate(text2);" 
  ShowHTMLCR "document." & strForm & "." & strField2 & ".value = datetostring(datafim);"
  ShowHTMLCR "}"
  ShowHTMLCR "if ((text > """") && (text2 > """")){"
  ShowHTMLCR "if (datainicio.valueOf() > datafim.valueOf()) {"
  ShowHTMLCR "alert(""" & strMessage & """);"
  ShowHTMLCR "document." & strForm & "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}}"
  ShowHTMLCR "return true;"
  ShowHTMLCR "}"
  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaCompareDate

REM -------------------------------------------------------------------------
REM Entra FormJavaCompareDateSep
REM -------------------------------------------------------------------------
Private Sub FormJavaCompareDateSep(ByVal strField, strField2, strField3, strField4, strField5, strField6, strForm , strFunctionName, strMessage, blnMaiorQueHoje, strMessage2)
                                         
  JavaScriptBegin
  
  ShowHTMLCR "function " & strFunctionName & "(){"
  ShowHTMLCR "var datainicio, datafim, dataatual, text, text2;"
  ShowHTMLCR "datainicio = new Date();"
  ShowHTMLCR "datafim    = new Date();"
  
  ShowHTMLCR "if ((document." & strForm & "." & strField & ".value > """")  && (document." & strForm & "." & strField2 & ".value > """") && (document." & strForm & "." & strField3 & ".value > """")){"           
  ShowHTMLCR "text       = document." & strForm & "." & strField & ".value + document." & strForm & "." & strField2 & ".value + document." & strForm & "." & strField3 & ".value;"           
  ShowHTMLCR "if (!isdate(text)){" 
  ShowHTMLCR "alert(""Data Inválida"");"
  ShowHTMLCR "document." & strForm & "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}}" 
  
  ShowHTMLCR "if ((document." & strForm & "." & strField4 & ".value > """")  && (document." & strForm & "." & strField5 & ".value > """") && (document." & strForm & "." & strField6 & ".value > """")){"           
  ShowHTMLCR "text2      = document." & strForm & "." & strField4 & ".value + document." & strForm & "." & strField5 & ".value + document." & strForm & "." & strField6 & ".value;"           
  ShowHTMLCR "if (!isdate(text2)){" 
  ShowHTMLCR "alert(""Data Inválida"");"
  ShowHTMLCR "document." & strForm & "." & strField4 & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}}"                                           
  
  ShowHTMLCR "if ((text > """") && (text2 > """")){"
  ShowHTMLCR "datainicio = todate(text);" 
  ShowHTMLCR "datafim    = todate(text2);" 
  
  If blnMaiorQueHoje Then
    ShowHTMLCR "dataatual  = new Date();"
    ShowHTMLCR "if (datainicio.valueOf() + 90000000 < dataatual.valueOf()) {"
    ShowHTMLCR "alert(""" & strMessage2 & """);"
    ShowHTMLCR "return false;"
    ShowHTMLCR "}"
  
  End If
  
  ShowHTMLCR "if (datainicio.valueOf() > datafim.valueOf()) {"
  ShowHTMLCR "alert(""" & strMessage & """);"
  ShowHTMLCR "document." & strForm & "." & strField & ".focus();"
  ShowHTMLCR "return false;"
  ShowHTMLCR "}}"

  ShowHTMLCR "return true;"
  ShowHTMLCR "}"
    
  JavaScriptEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormJavaCompareDateSep

REM -------------------------------------------------------------------------
REM Testa se e um CPF
REM -------------------------------------------------------------------------
Function IsCPF(ByVal strCPF)

Dim intSoma, i, strNumber, IntDigit1, IntDigit2, intMultiply, intNet
  If strCPF > "" Then
    intSoma = 0
    intMultiply = 10
    
    strCPF = TotalNormalizeString(strCPF)
    strCPF = leadingzeroes(strCPF, 11)
    
    For i = 1 To 9
        strNumber   = Int(Mid(strCPF, i, 1))
        intSoma     = intSoma + (strNumber * intMultiply)
        intMultiply = intMultiply - 1
    
    Next
    
    intNet = intSoma Mod 11
    
    If (intNet = 0) Or (intNet = 1) Then
       IntDigit1 = 0
    
    Else
       IntDigit1 = CInt(11 - intNet)
    
    End If
    
    If IntDigit1 = Int(Mid(strCPF, 10, 1)) Then
       intSoma     = 0
       intMultiply = 11
    
       For i = 1 To 10
           strNumber   = Int(Mid(strCPF, i, 1))
           intSoma     = intSoma + (strNumber * intMultiply)
           intMultiply = intMultiply - 1
    
       Next
    
       intNet = intSoma Mod 11
    
       If (intNet = 0) Or (intNet = 1) Then
          IntDigit2 = 0
       Else
          IntDigit2 = 11 - intNet
       End If
    
       If IntDigit2 = Int(Mid(strCPF, 11, 1)) Then
         IsCPF = True
 
       Else
         IsCPF = False
 
       End If
    Else
      IsCPF = False

    End If
    
    strCPF = ZTIFormatCGC(strCPF)
  Else
    IsCPF = True    

  End If
End Function


REM -------------------------------------------------------------------------
REM Testa se e um CGC
REM -------------------------------------------------------------------------
Function IsCGC(strCGC)
Dim intSoma, i, strNumber, IntDigit1, IntDigit2, intMultiply, intNet
  
  If strCGC > "" Then
    intMultiply = 5
    intSoma = 0
  
    strCGC = TotalNormalizeString(strCGC)
    strCGC = leadingzeroes(strCGC, 14)
    For i = 1 To 12
        strNumber = Int(Mid(strCGC, i, 1))
        intSoma = intSoma + (intMultiply * strNumber)
        intMultiply = intMultiply - 1
        If intMultiply < 2 Then
           intMultiply = 9
        End If
    Next
    intNet = intSoma Mod 11
    If (intNet = 0) Or (intNet = 1) Then
       IntDigit1 = 0
    Else
       IntDigit1 = 11 - intNet
    End If
    If IntDigit1 = Int(Mid(strCGC, 13, 1)) Then
       intMultiply = 6
       intSoma = 0
  
       For i = 1 To 13
           strNumber = Int(Mid(strCGC, i, 1))
           intSoma = intSoma + (intMultiply * strNumber)
           intMultiply = intMultiply - 1
  
           If intMultiply < 2 Then
              intMultiply = 9
           End If
       Next
  
       intNet = intSoma Mod 11
  
       If (intNet = 0) Or (intNet = 1) Then
          IntDigit2 = 0
       Else
          IntDigit2 = 11 - intNet
       End If
  
       If IntDigit2 = Int(Mid(strCGC, 14, 1)) Then
          IsCGC = True
  
       Else
          IsCGC = False
       End If
  
     Else
       IsCGC = False
     End If           
  
  Else
    IsCGC = True
  
  End If

End Function
REM ---------------------------------------------------------------------
REM Fim do /ZTITools/CPF.inc

REM -------------------------------------------------------------------------
REM Testa um Email
REM -------------------------------------------------------------------------
Function IsEmail(ByVal strEmail)
  IsEmail = ((InStr(strEmail, "@") <> 0) And (InStr(strEmail, ".") <> 0) And (InStr(strEmail, " ") = 0)) Or (Not strEmail > "")
    
End Function
REM ---------------------------------------------------------------------
REM Fim do IsEmail

REM -------------------------------------------------------------------------
REM Testa uma Data
REM -------------------------------------------------------------------------
Function IsValidDate(ByVal strDate)
Dim Dia, Mes, Ano, blnIsDate
  
  blnIsDate = True
  If Not LCase(TypeName(strDate)) = "date" Then
    If strDate > "" Then
      strDate = NormalizeString(strDate)
      Mes = Month(Now)
      Ano = Year(Now)
      If Len(strDate) > 1  Then
        Dia = CInt(Left(strDate, 2))
        If Len(strDate) > 2 Then
          Mes = CInt(Mid(strDate, 3, 2))
      
        End If
    
        If Len(strDate) > 4 Then
          Ano = CInt(Mid(strDate, 5, 4))
      
        End If
    
        If (Mes > 12) Or (Mes < 1) Then
          blnIsDate = False
        
        End If
        
        If (Dia < 1) Or (Dia > MonthSize(Mes, Ano)) Then
          blnIsDate = False
            
        End If
     
        If (Ano > Year(Now) + 60) Then
          blnIsDate = False
        
        End If
    
      Else
        blnIsDate = False
      
      End If
    End If
  End If

  IsValidDate = blnIsDate

End Function
REM ---------------------------------------------------------------------
REM Fim do IsDate

REM -------------------------------------------------------------------------
REM Testa uma Hora
REM -------------------------------------------------------------------------
Function IsValidTime(ByVal strTime)
Dim Hora, Minuto, blnIsTime
  
  blnIsTime = True
  If strTime > "" Then
    strTime = NormalizeString(strTime)
    Minuto  = 0
    Hora    = CInt(Left(strTime, 2))
    If Len(strTime) > 2 Then
      Minuto = CInt(Mid(strTime, 3, 2))
    
    End If
 
    If (Hora > 23) Then
      blnIsTime = False
    
    End If
    
    If (Minuto > 59) Then
      blnIsTime = False
        
    End If
  
    strTime = LeadingZeroes(Hora, 2) & ":" & LeadingZeroes(Minuto, 2)

  End If

REM  ShowJavaScriptMessage strTime

  IsValidTime =blnIsTime

End Function
REM ---------------------------------------------------------------------
REM Fim do IsValidTime

REM -------------------------------------------------------------------------
REM Formata uma Hora
REM -------------------------------------------------------------------------
Function FormatTime(strTime)
Dim Hora, Minuto
  
  strTime = NormalizeString(strTime)
  Minuto  = 0
  Hora    = CInt(Left(strTime, 2))
  If Len(strTime) > 2 Then
    Minuto = CInt(Mid(strTime, 3, 2))
   
  End If
 
  strTime = LeadingZeroes(Hora, 2) & ":" & LeadingZeroes(Minuto, 2)

  FormatTime = strTime

End Function
REM ---------------------------------------------------------------------
REM Fim do FormatTime

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Function CompareDate(ByVal strDate, strDate2)
Dim dtmDataInicio, dtmDataFim, intCompare
  
  intCompare = 0

  If strDate > "" Then
    If IsValidDate(strDate) Then
      dtmDataInicio = CDate(strDate)

    Else
      intCompare = 1
    
    End If  
  End If

  If strDate2 > "" Then
    If IsValidDate(strDate2) Then
      dtmDataFim = CDate(strDate2)
    
    Else
      intCompare = intCompare + 2
    
    End If  
  End If  

  If intCompare = 0 Then
    If strDate > "" And strDate2 > "" Then
      If dtmDataInicio > dtmDataFim Then
        intCompare = 4
      
      End If  
    End If  
  End If

  CompareDate = intCompare

End Function
REM -------------------------------------------------------------------------
REM Final da Sub CompareDate

REM -------------------------------------------------------------------------
REM Retorna o Tamanho de um Mes
REM -------------------------------------------------------------------------
Function MonthSize(intMes, intAno)
  Select Case intMes
    Case 1,3,5,7,8,10,12
      MonthSize = 31

    Case 4,6,9,11
      MonthSize = 30

    Case 2
      If intAno Mod 4 = 0 Then
        MonthSize = 29
      
      Else
        MonthSize = 28
        
      End If
                                                
  End Select    
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function MonthSize

%>