<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Tip.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Edicao
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

Dim sSetTipHeadClass, sSetTipBodyClass
sSetTipHeadClass = ""
sSetTipBodyClass = ""

Dim sSetTipOverColor, sSetTipNormalColor
sSetTipOverColor   = ""
sSetTipNormalColor = ""

Dim sintTipCounter, sintTipLevel
sintTipCounter = 1
sintTipLevel   = 1

SetTipHeadClass "dtitle"
SetTipBodyClass "dbody"

SetTipOverColor "Black"
SetTipNormalColor "Blue"

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Sub AddTip(strTitle, strBody)

  ShowHTMLCR "<tr><td align=left>"
  ShowHTML "<DIV ID=F" & sintTipCounter

  If sSetTipHeadClass > "" Then
      ShowHTML " class=" & sSetTipHeadClass
      
  End If
  
  ShowHTMLCR " onMouseOver=""makeCool(F" & sintTipCounter & ");"" onMouseOut=""makeNormal(F" & sintTipCounter & ");"" onClick=""OpenDesc(T" & sintTipCounter & ");"">" & strTitle & "</DIV>"
  
  ShowHTMLCR "</td></tr><tr><td align=left>"
  ShowHTMLCR "<DIV ID=T" & sintTipCounter 

  If sSetTipBodyClass > "" Then
      ShowHTML " class=" & sSetTipBodyClass
      
  End If
  
  ShowHTML " style=""display:'none';"">"
  ShowHTMLCR strBody
  ShowHTMLCR "<hr color=white></DIV></td></tr>"

  sintTipCounter = sintTipCounter + 1

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub AddTip

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Sub BeginTip

%>
<SCRIPT TYPE="text/javascript">
<!--
function OpenDesc(text) {
  if (text.style.display =="") {
    text.style.display = "none";
  }
  else {
    text.style.display="";
  }
}

function OpenDescOne(text) {
  if (text.style.display =="") {
    text.style.display = "none";
  }
  else {
    text.style.display="";
  }
}

function makeCool(text) {
  text.style.color = "<%=sSetTipOverColor%>";
}

function makeNormal(text) {
  text.style.color = "<%=sSetTipNormalColor%>";
}
//-->
</script>

<%

  ShowHTMLCR "<TABLE width=100% border=0 CELLPADDING=0 CELLSPACING=0>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BeginTip

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Sub EndTip

  ShowHTML "</TABLE>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EndTip

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Sub SetTipHeadClass(ByVal strClass)

  sSetTipHeadClass = strClass

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetTipHeadClass

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Sub SetTipBodyClass(ByVal strClass)

  sSetTipBodyClass = strClass

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetTipBodyClass

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Sub SetTipOverColor(ByVal strColor)

  sSetTipOverColor = strColor

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetTipHeadClass

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Sub SetTipNormalColor(ByVal strColor)

  sSetTipNormalColor = strColor

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetTipHeadClass

REM -------------------------------------------------------------------------
REM Adiciona um Item na Lista
REM -------------------------------------------------------------------------
Private Sub AddTipBegin(ByVal strTitle, ByVal strBody)

  ShowHTML "<TR><TD align=left ID=LF" & sintTipLevel & sintTipCounter 

  If sSetTipHeadClass > "" Then
      ShowHTML " class=" & sSetTipHeadClass
      
  End If
  
  ShowHTMLCR " onMouseOver=""makeCool(LF" & sintTipLevel & sintTipCounter & ");"" onMouseOut=""makeNormal(LF" & sintTipLevel & sintTipCounter & ");"" onClick=""OpenDesc(LT" & sintTipLevel & sintTipCounter & ");"" STYLE='cursor: hand'>" & strTitle
  
  ShowHTMLCR "</TD></TR><TR><TD>"
    
  ShowHTMLCR "<TABLE border=" & Session("TableBorder") & " width=100% CELLPADDING=0 CELLSPACING=0>"

  ShowHTML "<TR><TD>&nbsp;</TD><TD ID=LT" & sintTipLevel & sintTipCounter 

  If sSetTipBodyClass > "" Then
      ShowHTML " class=" & sSetTipBodyClass
      
  End If
  
  ShowHTMLCR " style=""display:'none';"">"

  If strBody > "" Then
    ShowHTMLCR strBody
    ShowHTMLCR "<hr color=white>"
    
  End If

  ShowHTMLCR "<TABLE border=" & Session("TableBorder") & " width=100% CELLPADDING=0 CELLSPACING=0>"

  sintTipCounter = sintTipCounter + 1

  sintTipLevel = sintTipLevel + 1

End Sub
REM -------------------------------------------------------------------------
REM Final da AddTipBegin

REM -------------------------------------------------------------------------
REM Muda para o nivel anterior
REM -------------------------------------------------------------------------
Private Sub AddTipEnd

  sintTipLevel = sintTipLevel - 1

  ShowHTMLCR "</TABLE></TD></TR></TABLE></TR></TD>"

End Sub
REM -------------------------------------------------------------------------
REM Final da AddTipEnd

REM -------------------------------------------------------------------------
REM Fim do Tip.inc
%>

