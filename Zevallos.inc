<%
REM -------------------------------------------------------------------------
REM  /Zevallos.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do Zevallos
REM  Cria‡ao     : 11:23h 14/2/1998
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

Session("SourceCode") = True

Const conSiteName         = "ZTI"
Const conSiteTitle        = "Zevallos Tecnologia em Informacao"
Const conSiteMainSelector = "/Default.asp"
Const conSiteFirstPage    = "/News.asp"

REM -----------------------------------------------------------------------
REM Diversas
REM -----------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Definicao Global
REM -------------------------------------------------------------------------
Dim sstrArquivo
Dim sblnFirtTime

sstrArquivo = LCase(Request.QueryString("P"))

sblnFirtTime = False

If Not sstrArquivo > "" Then
  sstrArquivo = conSiteFirstPage

  If sstrThisScriptName = conSiteMainSelector Then
    sblnFirtTime = True
  
  End If
  
End If

InitSystem

REM -------------------------------------------------------------------------
REM Monta a Frame
REM -------------------------------------------------------------------------
Sub MountFrame
  ShowHTMLCR "<html>"
  ShowHTMLCR "<head>"
  ShowHTMLCR "<title>" & conSiteTitle & "</title>"
  ShowHTMLCR "</head>"
  ShowHTMLCR "<frameset rows=100% cols=""166,*"" framespacing=0>"
  ShowHTMLCR "<frame name=Menu src=menu.asp?xx=1 scrolling=no marginheight=0 marginwidth=0 frameborder=0 noresize>"
  ShowHTMLCR "<frame name=Body src=" & conSiteMainSelector & "?P=" & conSiteFirstPage & "&" & Request.ServerVariables("QUERY_STRING") & " scrolling=auto marginheight=0 marginwidth=0 frameborder=0>"
  ShowHTMLCR "</frameset>"
  ShowHTMLCR "</html>"
  
  Session(conSiteName)="frame"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub MountFrame

Sub InitSiteDefault
REM -------------------------------------------------------------------------
  Session("Default") 				  = conSiteName
  
  Session("MetaPragma")       = "NO-CACHE"
  Session("MetaContentType")  = "text/html; charset=iso-8859-1"
  Session("MetaReplyTo")      = "webmaster@zevallos.com.br"
  Session("MetaExpires")      = now + 90
  Session("MetaPicsLabel")    = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html""" &  _ 
                                " l gen true comment ""RSACi North America Server"" by" & _ 
                                " ""webmaster@zevallos.com.br"" for """ & _
                                "http://www.zevallos.com.br""" & _
                                " on ""1997.06.26T21:24-0500"" r (n 0 s 0 v 0 l 0))"
  Session("MetaKeywords")     = "JScript, Java, ActiveX, JavaScript, dynamic html, dhtml, html, html 4.0, style sheets, css, javascript, jscript, layers, dhtml web development, netscape 4.0, explorer 4.0, document object model, dhtml tutorials, dhtml help, dhtml template, " & _
                                "Dynamic HTML, DHTML, Style Sheets, CSS, Javascript, Jscript, Layers, DHTML Web Development, Netscape 4.0, Explorer 4.0, HTML 4.0, Document Object Model, DHTML Tutorials, DHTML Help, DHTML Template"

  Session("MetaDescription")  = "Atuando na  rea de tecnologia da informa‡ao, a Zevallos? ‚ especializada em transformar informa‡oes com o emprego de tecnologia disponibilizando-a em outros meios, preferencialmente a Internet. " & _
                                "Nossa missao: transformar a Internet em um meio onde a informa‡ao em qualquer formato possa ser acessada de forma r pida e desburocratizada. " & _
                                "Nossos profissionais, experientes e altamente capacitados, contam com mais de 5 anos de experiˆncia no uso de ferramentas para Internet. Nossas solu‡oes equilibram e harmonizam em seus projetos e propostas as mais modernas inova‡oes tecnol¢gicas com as reais necessidades de informa‡ao, simplicidade operacional e recursos do cliente. " & _
                                "No desenvolvimento utilizamos tecnologias de £ltima gera‡ao, tais como Java, JavaScript, ActiveX, VBScript, Active Server Pages, NetShow, RealVideo e RealAudio. " & _
                                "Diferente das demais empresas, sua metodologia de trabalho, a pesquisa continuada, a an lise das tecnologias emergentes e a aplica‡ao dos resultados cria seu grande diferencial. " & _
                                "A aplica‡ao desta metodologia e a busca por solu‡oes e melhorias na performance dos processos possibilitou o desenvolvimento do ZTITools?2.5, um conjunto de ferramentas que potencializam e reduzem o tempo de desenvolvimento de sistemas para menos de um quinto do per¡odo normal. " & _
                                "Zevallos?, sin“nimo de tecnologia da Informa‡ao aplicada a Internet ou intranets."
  
  
  Session("MetaProduct")      = "ZTITools v2.5 - Biblioteca ASP"
  Session("MetaLocale")       = "PO-BR"
  Session("MetaCharSet")      = "US-ASCII"
  Session("MetaCategory")     = "HOME PAGE"
  Session("MetaGenerator")    = "Tecnologias da Zevallos em ASP"
  Session("MetaAuthor")       = "ZTI - Zevallos(r) Tecnologia em Informacao - Brasilia - DF -" & _
                                " Brasil - webmaster@zevallos.com.br - http://www.zevallos.com.br"
  Session("MetaDistribution") = "global"
  Session("MetaFormatter")    = "ZTITools v2.0"
  Session("MetaTechnology")   = "SCRIPTING"

REM -------------------------------------------------------------------------
REM Constantes BASE FONT de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("BaseFont") = "Arial, Helvetica, Sans-Serif"

REM -------------------------------------------------------------------------
REM Constantes LINK de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("LinkHomeHRef")       = "/Default.asp"
  Session("LinkHomeTitle")      = "ZTITools v2.0"

  Session("LinkCopyrightHRef")  = "/copyright.htm"
  Session("LinkCopyrightTitle") = "Copyright da Zevallos"

  Session("LinkMadeHRef")       = "mailto:webmaster@zevallos.com.br"
  Session("LinkMadeTitle")      = "WebMaster da Zevallos"

  Session("LinkStyleSheetHRef")  = "/default.css"
  
REM -------------------------------------------------------------------------
REM Constantes BODY de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("BodyBGColor")    = "white"
  Session("BodyBackground") = "/ztitools/img/Assinatura75Back.Gif"
  Session("BodyStyle")      = ""
  Session("BodyText")       = ""
  Session("BodyLink")       = ""
  Session("BodyVLink")      = ""
  Session("BodyALink")      = ""
  Session("BodyTopMargin")  = 0
  Session("BodyLeftMargin") = 0
  Session("BodyWidth")      = 610

REM -------------------------------------------------------------------------
REM Constantes HEADER de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("HeaderIcon")       = ""
  Session("HeaderTitle")      = ""

REM -------------------------------------------------------------------------
REM Constantes FOOTER de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("FooterTitle")      = "Um produto da <A HREF=http://www.zevallos.com.br>" & _
                                "Zevallos&reg;</A> Tecnologia em Informa&ccedil;&atilde;o."
  Session("FooterSugestions") = "Sugest&otilde;es ou problemas encaminhar para o " & _
                                "<A HREF=mailto:webmaster@zevallos.com.br" & _
                                "<img src=/ztitools/img/icone/mailto.gif alt=""Mail To"" border=0> " & _
                                "<i>&lt;webmaster@zevallos.com.br&gt;</i></A>"
  Session("FooterCopyright")  = "&copy; 1997, 1998 <A HREF =default.asp?a=/copyright.htm>Zevallos&reg;</a>" & _
                                " todos os direitos reservados."
                                
REM Send questions or comments about this Website to the webmaster@isgsoft.com.
REM For all other inquiries, please see the Contact Us page.
REM Copyright (c)1997, 1998 ISG International Software Group, Ltd.
REM Current Time at the server is: 12:23:22 AM
REM Your IIS session expires 60 minutes from the time above.
                                
End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub InitSiteDefault

REM -------------------------------------------------------------------------
REM Inicializacao do sistema
REM -------------------------------------------------------------------------
Private Sub InitSystem

  If Not Session("Default") = conSiteName Or Request.QueryString("Site") > "" Or Not Session(conSiteName)="frame" Or sblnFirtTime Then
    InitSiteDefault
   
  End If

  If Not Request.ServerVariables("QUERY_STRING") > "" Then
    MountFrame
    
  End If
  
  REM -----------------------------------------------------------------------
  REM Inicializa os defaults do sistema
  REM -----------------------------------------------------------------------
  If Request.QueryString("Default") > "" Then
      InitDefault

  End If
  
  REM -----------------------------------------------------------------------
  REM Ativa o icone de ASP Source Code
  REM -----------------------------------------------------------------------
  If Request.QueryString("Source") > "" Then
    Select Case Request.QueryString("Source")
      Case 0
        Session("SourceCode") = False

      Case 1
        Session("SourceCode") = True

      Case Else
    End Select
  Else
    If IsEmpty(Session("SourceCode")) Then Session("SourceCode") = False

  End If

  REM -----------------------------------------------------------------------
  REM Define a borda das tabelas
  REM -----------------------------------------------------------------------
  If Request.QueryString("Border") > "" Then
    Session("TableBorder") = Request.QueryString("Border")

  Else
    If IsEmpty(Session("TableBorder")) Then Session("TableBorder") = 0

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub InicioSistema

REM -------------------------------------------------------------------------
REM Mostra o titulo das novidades com sobra via CCS
REM -------------------------------------------------------------------------
Sub ShowNewsTitle(txtMessage)

  ShowHTMLCR "<DIV CLASS=ntBack>" & txtMessage & "</DIV>"
  ShowHTMLCR "<DIV CLASS=ntFront>" & txtMessage & "</DIV>"

End Sub 
REM -------------------------------------------------------------------------
REM Fim da Sub ShowNewsTitle

REM -------------------------------------------------------------------------
REM Mostra o titulo das novidades com sobra via CCS
REM -------------------------------------------------------------------------
Sub ShowNewsBody(txtMessage)

  ShowHTMLCR "<SPAN CLASS=newsbody>" & txtMessage & "</SPAN>"

  ShowHTMLCR "<HR NOSHADE SIZE=1>"

End Sub 
REM -------------------------------------------------------------------------
REM Fim da Sub ShowNewsTitle

REM -------------------------------------------------------------------------
REM Mostra o titulo do Menu
REM -------------------------------------------------------------------------
Sub ShowMenuTitle(txtMessage)

  ShowHTMLCR "<TR><TD CLASS=Title VALIGN=TOP colspan=2>" & txtMessage
  ShowHTMLCR "</TD></TR>"

End Sub 
REM -------------------------------------------------------------------------
REM Fim da Sub ShowMenuTitle

REM -------------------------------------------------------------------------
REM Mostra o Item do Menu
REM -------------------------------------------------------------------------
Sub ShowMenuItem(txtURL, txtMessage)

  ShowHTMLCR "<TR><TD width=1% VALIGN=TOP>"
  ShowHTMLCR "<A HREF=" & txtURL & ">"
  ShowHTMLCR "<IMG BORDER=0 HSPACE=5 VSPACE=3 WIDTH=8 HEIGHT=7 "
  ShowHTMLCR "SRC=/ZTITools/img/icone/purpbullet.gif TARGET=Body></A></TD>"
  ShowHTMLCR "<TD VALIGN=TOP>"
  ShowHTMLCR "<A CLASS=Menu HREF=" & txtURL & " TARGET=Body>" & txtMessage & "</A><BR></TD></TR>"

End Sub 
REM -------------------------------------------------------------------------
REM Fim da Sub ShowMenuItem

REM -------------------------------------------------------------------------
REM Mostra o Espaco entre os menus
REM -------------------------------------------------------------------------
Sub ShowMenuSpacer()

  ShowHTMLCR "<TR><TD COLSPAN=2>"
  ShowHTMLCR "<IMG VSPACE=10 HEIGHT=1 WIDTH=1 SRC=/img/icone/blk.gif></TD></TR>"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowMenuSpacer

REM -------------------------------------------------------------------------
REM Mostra o Inicio do Menu
REM -------------------------------------------------------------------------
Sub ShowMenuBegin()

  ShowHTMLCR "<TABLE CLASS=Menu CELLPADDING=0 CELLSPACING=0 border=0 WIDTH=160>"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowMenuBegin

REM -------------------------------------------------------------------------
REM Mostra o Fim do Menu
REM -------------------------------------------------------------------------
Sub ShowMenuEnd()

  ShowHTMLCR "</TABLE>"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowMenuEnd

REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Mostra Inicio do Icone 
REM -------------------------------------------------------------------------
Sub ShowIconBegin()

  ShowHTML "<TR><TD CLASS=End COLSPAN=2>"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowIconBegin


REM -------------------------------------------------------------------------
REM Mostra Fim do Icone 
REM -------------------------------------------------------------------------
Sub ShowIconEnd()

  ShowHTML "</TD></TR><TR><TD COLSPAN=2>&nbsp;</TD></TR>"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowIconEnd

REM -------------------------------------------------------------------------
REM Mostra o Titulo da Pagina
REM -------------------------------------------------------------------------
Sub ShowPageTitle(txtSRC, txtAlt, txtTitle)

  ShowHTMLCR "<table width=100% border=0>"
  ShowHTMLCR "<tr><td width=1% valing=top>"
  ShowHTMLCR "<img src=""" & txtSRC & """ alt="" & txtAlt & ""></td><td>"
  ShowHTMLCR "<DIV CLASS=PgHeaderBack>" & txtTitle & "</DIV>"
  ShowHTMLCR "<DIV CLASS=PgHeaderMiddle>" & txtTitle & "</DIV>"
  ShowHTMLCR "<DIV CLASS=PgHeaderFront>" & txtTitle & "</DIV>"
  ShowHTMLCR "</table>"
  ShowHTMLCR "<hr>"

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub ShowPageTitle

REM -------------------------------------------------------------------------
REM Mostra o Titulo da Pagina
REM -------------------------------------------------------------------------
Sub ShowPageCookie
    
  ShowHTMLCR "Ol&aacute;"
  
  If LastVisited Then
    If Request.Cookies("zevallosasp").HasKeys Then
      ShowHTMLCR " " & Request.Cookies("zevallosasp")("NomePrimeiro") & ". "
      
    Else
      ShowHTMLCR ", que bom que voc&ecirc; voltou. "
      
    End If
  
    ShowHTMLCR "Sua &uacute;ltima visita foi "
    ShowHTMLCR Session("Last_Visited")
    
  Else
    ShowHTMLCR ", seja bem-vindo e aproveite a sua primeira visita."
    
  End If
          
End Sub

REM -------------------------------------------------------------------------
REM Fim da Sub ShowPageCookie

REM -------------------------------------------------------------------------
REM Fim do Zevallos.inc
%>