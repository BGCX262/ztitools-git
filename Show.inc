<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Show.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Shows
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
Dim sblnShowComments
Dim sblnShowCR
  sblnShowComments = True
  sblnShowCR = True

REM -------------------------------------------------------------------------
REM Mostra um registro completo
REM -------------------------------------------------------------------------
Public Sub ShowCounterIni(ByVal strMessage, ByVal intModule)
  Session("CounterMsg") = strMessage
  Session("CounterMod") = intModule
  Session("CounterCnt") = 0
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ListFields

REM -------------------------------------------------------------------------
REM Mostra um registro completo
REM -------------------------------------------------------------------------
Public Sub ShowCounterAdd

  Session("CounterCnt") = Session("CounterCnt") + 1

  If Session("CounterCnt") > 0 Then
    If Session("CounterCnt") mod Session("CounterMod") = 0 Or Session("CounterCnt") = 1 Then
       ShowCounter
       
    End If 
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowCounter

REM -------------------------------------------------------------------------
REM Mostra um registro completo
REM -------------------------------------------------------------------------
Public Sub ShowCounter
  Response.Write Session("CounterMsg") & Session("CounterCnt")

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowCounter

REM -------------------------------------------------------------------------
REM Mostra um registro completo
REM -------------------------------------------------------------------------
Public Sub ShowField(ByVal objRSSL, ByVal i)
  ShowMessage "i=" & i & " - " & objRSSL(i).Name & " - " & objRSSL(i).Type & " - " & objRSSL(i).Value

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowRecord

REM -------------------------------------------------------------------------
REM Mostra resultado HTML
REM -------------------------------------------------------------------------
Public Sub ShowHTML(ByVal strHTML)
  Response.Write strHTML

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowHTML

Public Sub SetShowCR(blnShow)
  sblnShowCR = blnShow

End Sub

Public Function CRLF
  If sblnShowCR Then
    CRLF = Chr(13)

  End If

End Function

REM -------------------------------------------------------------------------
REM Mostra resultado HTML + CR
REM -------------------------------------------------------------------------
Public Sub ShowHTMLCR(ByVal strHTML)
  Response.Write strHTML & CRLF

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowHTMLCR

REM -------------------------------------------------------------------------
REM Mostra resultado CR + HTML
REM -------------------------------------------------------------------------
Public Sub ShowCRHTML(ByVal strHTML)
  Response.Write CRLF & strHTML

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowCRHTML

REM -------------------------------------------------------------------------
REM Mostra resultado CR + HTML
REM -------------------------------------------------------------------------
Public Sub ShowCRHTMLCR(ByVal strHTML)
  Response.Write CRLF & strHTML & CRLF

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowCRHTMLCR

Public Sub SetShowComment(blnShow)
  sblnShowComments = blnShow

End Sub

REM -------------------------------------------------------------------------
REM Mostra resultado CR + HTML
REM -------------------------------------------------------------------------
Public Sub ShowHTMLComment(ByVal strHTML)
  If sblnShowComments Then
    Response.Write "<!" & "-- " & strHTML & " --" & ">" & CRLF

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowCRHTMLCR

REM -------------------------------------------------------------------------
REM Mostra uma mensagem
REM -------------------------------------------------------------------------
Public Sub ShowMessage(ByVal strMessage)

  Response.Write "<BR>" & strMessage

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMessage

REM -------------------------------------------------------------------------
REM Mostra uma mensagem normal
REM -------------------------------------------------------------------------
Public Sub ShowMessageNormal(ByVal strMessage)

  ShowMessage strMessage

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMessageNormal

REM -------------------------------------------------------------------------
REM Mostra uma mensagem de Erro
REM -------------------------------------------------------------------------
Public Sub ShowMessageError(ByVal strMessage)

  ShowMessage "<font color=red><B>" & strMessage & "</B></font>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMessageError

REM -------------------------------------------------------------------------
REM Mostra uma mensagem de Alerta
REM -------------------------------------------------------------------------
Public Sub ShowMessageAlert(ByVal strMessage)

  ShowMessage "<font color=Orange><B>" & strMessage & "</B></font>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMessageAlert

REM -------------------------------------------------------------------------
REM Mostra uma mensagem de Erro
REM -------------------------------------------------------------------------
Public Sub ShowJavaScriptMessage(ByVal strMessage)
  If strMessage > "" Then
    Response.Write "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""" & strMessage & """);</SCRIPT>"
    
  Else
    Response.Write "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""Não há conteúdo neste ShowJavaScriptMessage"");</SCRIPT>"
    
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMessageError

REM -------------------------------------------------------------------------
REM Mostra o campo Link Rel
REM -------------------------------------------------------------------------
Public Sub ShowLinkRelStyleSheet(ByVal strHRef)

  ShowLink "REL", "STYLESHEET", strHRef, "", "text/css"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowLinkRel

REM -------------------------------------------------------------------------
REM Mostra a lista de variaveis do servidor
REM -------------------------------------------------------------------------
Public Sub ShowAllServerVariables
  Dim key

  ShowHTMLCR "<TABLE BORDER=1>"
  ShowHTMLCR "<TR><TD VALIGN=TOP><B>Variable</B></TD><TD VALIGN=TOP><B>Value</B></TD></TR>"
  
  For Each key in Request.ServerVariables
    ShowHTMLCR "<TR><TD>(" & key & "</TD>"
    ShowHTMLCR "<TD>"
  
    If Request.ServerVariables(key) = "" Then
      ShowHTMLCR "&nbsp;"
      
    Else 
      ShowHTMLCR Request.ServerVariables(key)
      
    End If
    ShowHTMLCR "</TD></TR>"
  
  Next

  ShowHTMLCR "</TABLE>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowAllServerVariables

REM -------------------------------------------------------------------------
REM Mostra o campo Link
REM -------------------------------------------------------------------------
Public Sub ShowLink(ByVal strTag, ByVal strName, ByVal strHRef, ByVal strTitle, ByVal strType)
  DIM strResult
  
  strResult = "<LINK"
  
  If strTag > "" Then
    strResult = strResult & " " & strTag

    If strName > "" Then
      strResult = strResult & "=""" & strName & """"
    
    End If
  End If

  If strHRef > "" Then
    strResult = strResult & " HREF=""" & strHRef & """"
    
  End If

  If strTitle > "" Then
    strResult = strResult & " TITLE=""" & strTitle & """"
    
  End If

  If strType > "" Then
    strResult = strResult & " TYPE=""" & strType & """"
    
  End If

  strResult = strResult & ">"

  ShowHTMLCR strResult
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowLink

REM -------------------------------------------------------------------------
REM Mostra o campo Link Rev
REM -------------------------------------------------------------------------
Public Sub ShowLinkRev(ByVal strName, ByVal strHRef, ByVal strTitle)

  ShowLink "REV", strName, strHRef, strTitle, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowLinkRev

REM -------------------------------------------------------------------------
REM Mostra o campo Link Rel
REM -------------------------------------------------------------------------
Public Sub ShowLinkRel(ByVal strName, ByVal strHRef, ByVal strTitle)

  ShowLink "REL", strName, strHRef, strTitle, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowLinkRel

REM -------------------------------------------------------------------------
REM Mostra o campo Meta
REM -------------------------------------------------------------------------
Public Sub ShowMeta(ByVal strField, ByVal strName, ByVal strContent)
  Dim strResult
  
  strResult = "<META " & strField & "=""" & strName & """ CONTENT="

  If InStr("'", strContent) = 0 Then
    strResult = strResult & """" & strContent & """"
    
  Else 
    StrResult = strResult & "'" & strContent & "'"
    
  End If
  
  ShowHTMLCR strResult & ">"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMeta

REM -------------------------------------------------------------------------
REM Mostra o campo Meta Name
REM -------------------------------------------------------------------------
Public Sub ShowMetaName(ByVal strName, ByVal strContent)
  ShowMeta "NAME", strName, strContent

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMetaName

REM -------------------------------------------------------------------------
REM Mostra o campo Meta HTTP-Equiv
REM -------------------------------------------------------------------------
Public Sub ShowMetaHTTPEquiv(ByVal strName, ByVal strContent)
  ShowMeta "HTTP-EQUIV", strName, strContent

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowMetaHTTPEquiv

REM -------------------------------------------------------------------------
REM Inicia a Lista de Itens
REM -------------------------------------------------------------------------
Public Sub ShowItemListBegin(ByVal strWidth, ByVal strBorder)
  ShowHTMLCR "<TABLE WIDTH=" & strWidth & " BORDER=" & strBorder & ">"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowItemListBegin

REM -------------------------------------------------------------------------
REM Inicia a Lista de Itens
REM -------------------------------------------------------------------------
Public Sub ShowItemListEnd()
  ShowHTMLCR "</TABLE>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowItemListEnd

REM -------------------------------------------------------------------------
REM Inicia a Lista de Itens
REM -------------------------------------------------------------------------
Public Sub ShowItemList(ByVal strTitle, ByVal strContent)
  ShowHTMLCR "<TR><TD NOWRAP=NOWRAP><B>" & strTitle & "</B></TD><TD>" & strContent & "</TD></TR>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowItemList

REM -------------------------------------------------------------------------
REM Inicia a Lista de Itens
REM -------------------------------------------------------------------------
Public Sub ShowItemListError(ByVal strTitle, ByVal strContent)
  ShowHTMLCR "<TR><TD NOWRAP=NOWRAP><FONT COLOR=RED><B>" & strTitle & "</B></FONT></TD><TD>" & strContent & "</TD></TR>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowItemList

REM -------------------------------------------------------------------------
REM Centraliza
REM -------------------------------------------------------------------------
Public Sub Center
  Response.Write "<CENTER>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Center

REM -------------------------------------------------------------------------
REM Centraliza
REM -------------------------------------------------------------------------
Public Sub CenterEnd
  Response.Write "</CENTER>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Center

REM -------------------------------------------------------------------------
REM Paragrafo seguido de Espaco
REM -------------------------------------------------------------------------
Public Sub PNbsp
  Response.Write "<P>&nbsp;"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub PNbsp
                
REM -------------------------------------------------------------------------
REM Espaco
REM -------------------------------------------------------------------------
Public Sub Nbsp
  Response.Write "&nbsp;"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Nbsp
                
REM -------------------------------------------------------------------------
REM Pula Linha
REM -------------------------------------------------------------------------
Public Sub BR
  Response.Write "<BR>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BR

REM -------------------------------------------------------------------------
REM Paragrafo
REM -------------------------------------------------------------------------
Public Sub Paragraph
  Response.Write "<P>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Paragraph

REM -------------------------------------------------------------------------
REM Texto em Negrito
REM -------------------------------------------------------------------------
Public Sub ShowBoldText(ByVal strText)
  Response.Write "<B>" & strText & "</B>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowBoldText
                
REM -------------------------------------------------------------------------
REM Texto em Italico
REM -------------------------------------------------------------------------
Public Sub ShowItalicText(ByVal strText)
  Response.Write "<I>" & strText & "</I>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowItalicText
                
REM -------------------------------------------------------------------------
REM Texto em Sublinhado
REM -------------------------------------------------------------------------
Public Sub ShowUnderlineText(ByVal strText)
  Response.Write "<U>" & strText & "</U>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowUnderlineText
                
REM -------------------------------------------------------------------------
REM Inicia Negrito
REM -------------------------------------------------------------------------
Public Sub Bold
  Response.Write "<B>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Bold
                
REM -------------------------------------------------------------------------
REM Finaliza Negrito
REM -------------------------------------------------------------------------
Public Sub BoldEnd
  Response.Write "</B>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Bold
                
REM -------------------------------------------------------------------------
REM Inicia Italico
REM -------------------------------------------------------------------------
Public Sub Italic
  Response.Write "<I>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Italic
                
REM -------------------------------------------------------------------------
REM Finaliza Italico
REM -------------------------------------------------------------------------
Public Sub ItalicEnd
  Response.Write "</I>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Bold
                
REM -------------------------------------------------------------------------
REM Inicia Sublinhado
REM -------------------------------------------------------------------------
Public Sub Underline
  Response.Write "<U>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Underline
                
REM -------------------------------------------------------------------------
REM Finaliza Sublinhado
REM -------------------------------------------------------------------------
Public Sub UnderlineEnd
  Response.Write "</U>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub UnderlineEnd

REM -------------------------------------------------------------------------
REM Rotina que emite uma Mensagem
REM -------------------------------------------------------------------------
Private Sub ShowMessageWindow(ByVal strCabecalho, strTexto, strRetorno)

  BR
  Center
  
  FormBegin strRetorno, "POST", "frmMensagem", "", ""

  TableBegin strCabecalho, 1, "40%"
  
  TableBeginRow 2
  TableBeginColumn
  Center
  ShowHTMLCR strTexto
  TableEndColumn
  TableEndRow

  TableBeginRow 2
  TableBeginColumn
  Center
  FormInputSubmit "bmtVolta", constrOK
  TableEndColumn
  TableEndRow

  TableEnd

  FormEnd

  ShowHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
  ShowHTMLCR "document.frmMensagem.bmtVolta.focus()"
  ShowHTMLCR "</SCRIPT>"
  
  CenterEnd
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Mensagem

REM -------------------------------------------------------------------------
REM Muda a formatacao de fonte
REM -------------------------------------------------------------------------
Private Function Font(ByVal intSize, ByVal strColor, ByVal strFace)
Dim strFont

  strFont = "<FONT"
  If CStr(intSize) > "" Then
    strFont = strFont & " SIZE=" & intSize

  End If

  If strColor > "" Then
    strFont = strFont & " COLOR=" & strColor

  End If

  If strFace > "" Then
    strFont = strFont & " FACE=" & strFace

  End If

  strFont = strFont & ">"

  Font = strFont

End Function
REM -------------------------------------------------------------------------
REM Final da Function Font

REM -------------------------------------------------------------------------
REM Finaliza Fonte
REM -------------------------------------------------------------------------
Private Function FontEnd
  FontEnd = "</FONT>"

End Function
REM -------------------------------------------------------------------------
REM Final da Function FontEnd
                
REM -------------------------------------------------------------------------
REM Fim do Show.inc
%>