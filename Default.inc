<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Default.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Defaults
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

Dim sDefaultLimit
Dim sDefaultCenter

sDefaultLimit = True
sDefaultCenter = False

REM -------------------------------------------------------------------------
REM Inicio do HTML
REM -------------------------------------------------------------------------
Public Sub HTMLBegin

  ShowHTMLCR "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN" &_
             """ ""http://www.w3.org/TR/REC-html40/loose.dtd"">"
  ShowHTMLCR "<HTML>"
  ShowHTMLComment "ZTITools v2.5 - Library ASP" &  _ 
                  "1997, 1998 (c) Zevallos(r) Tecnologia em Informacao" & _ 
                  "Brasilia - DF - Brasil" & _ 
                  "webmaster@zevallos.com.br" & _ 
                  "http://www.zevallos.com.br" 

  If Session("UserID") > 0 Then
    ShowHTMLComment "User ID: " & Session("UserID")
    
  End If

  If Session("SiteCounter") > 0 Then
    ShowHTMLComment "Site Counter: " & Session("SiteCounter")
    
  End If
  
  Dim objFS, objFile
 
  Set objFS   = CreateObject("Scripting.FileSystemObject")
  Set objFile = objFS.GetFile(sstrThisScriptFull)
  
  ShowHTMLComment "File Name: " & sstrThisScriptURL
  ShowHTMLComment "Created: " & objFile.DateCreated
  ShowHTMLComment "Last Accessed: " & objFile.DateLastAccessed
  ShowHTMLComment "Last Modified: " & objFile.DateLastModified  
  
  Set objFile = Nothing
  Set objFS = Nothing

  ShowHTMLComment "Script Engine: " & ScriptEngine & " - " & _
                  ScriptEngineMajorVersion & "." & _
                  ScriptEngineMinorVersion & " (Build:" & _
                  ScriptEngineBuildVersion & ")"
  
  ShowHTMLComment "sstrThisHTTPHost    = " & sstrThisHTTPHost   
  ShowHTMLComment "sstrThisScriptName  = " & sstrThisScriptName 
  ShowHTMLComment "sstrThisScriptURL   = " & sstrThisScriptURL  
  ShowHTMLComment "sstrThisSiteRootDir = " & sstrThisSiteRootDir
  ShowHTMLComment "sstrThisScriptFull  = " & sstrThisScriptFull 
  ShowHTMLComment "sstrThisQueryString = " & sstrThisQueryString
                  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub HTMLBegin

REM -------------------------------------------------------------------------
REM Fim do HTML
REM -------------------------------------------------------------------------
Public Sub HTMLEnd

  ShowHTMLCR "</HTML>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub HTMLEnd

REM -------------------------------------------------------------------------
REM Inicio do Head com o Title
REM -------------------------------------------------------------------------
Public Sub HeadBegin(ByVal strTitle)
  
  ShowHTMLComment "Head Section Begin"
  ShowHTMLCR "<HEAD>"
  ShowHTMLCR "<TITLE>" & strTitle & "</TITLE>"
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub HeadBegin

REM -------------------------------------------------------------------------
REM Fim do Head
REM -------------------------------------------------------------------------
Public Sub HeadEnd

  ShowHTMLCR "</HEAD>"
  ShowHTMLComment "Head Section End"
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub HeadEnd

REM -------------------------------------------------------------------------
REM Inicio do Body
REM -------------------------------------------------------------------------
Public Sub BodyBegin
  Dim strResult

  strResult = "<BODY"

  If Session("BodyBGColor") > "" Then
    strResult = strResult & " BGCOLOR=" & Chr(34) & Session("BodyBGColor") & Chr(34)
    
  End If

  If Session("BodyBackground") > "" Then
    strResult = strResult & " BACKGROUND=" & Chr(34) & Session("BodyBackground") & Chr(34)
    
  End If

  If Session("BodyText") > "" Then
    strResult = strResult & " TEXT=" & Session("BodyText")
    
  End If

  If Session("BodyLink") > "" Then
    strResult = strResult & " LINK=" & Session("BodyLink")
    
  End If

  If Session("BodyVLink") > "" Then
    strResult = strResult & " VLINK=" & Session("BodyVLink")
    
  End If

  If Session("BodyALink") > "" Then
    strResult = strResult & " ALINK=" & Session("BodyALink")
    
  End If
  
  If Session("BodyTopMargin") > "" Then
    strResult = strResult & " TOPMARGIN=" & Session("BodyTopMargin")
    
  End If

  If Session("BodyLeftMargin") > "" Then
    strResult = strResult & " LEFTMARGIN=" & Session("BodyLeftMargin")
    
  End If

  If Session("BodyStyle") > "" Then
    strResult = strResult & " STYLE=" & Session("BodyStyle")
    
  End If

  If Session("BodyBGProperties") > "" Then
    strResult = strResult & " BGPROPERTIES=" & Session("BodyBGProperties")
    
  End If

  If Session("BodyJavaScript") > "" Then
    strResult = strResult & " " & Session("BodyJavaScript")
    
  End If

  ShowHTMLComment "Main Body Section Begin"
  ShowHTMLCR strResult & ">"

  If sDefaultCenter Then
    ShowHTMLCR "<CENTER>"
    
  End If
  
  If sDefaultLimit Then
    strResult = "<TABLE"
    
    If Session("BodyWidth") > "" Then
      strResult = strResult & " WIDTH=" & Session("BodyWidth")
      
    End If
  
    If Session("TableBorder") > "" Then
      strResult = strResult & " BORDER=" & Session("TableBorder")
  
    Else
      strResult = strResult & " BORDER=0"
        
    End If
    
    ShowHTMLCR strResult & " cellpadding=0 cellspacing=0><TR><TD>"
    
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BodyBegin

REM -------------------------------------------------------------------------
REM Inicio do Body e parametros
REM -------------------------------------------------------------------------
Public Sub BodyLimit(ByVal blnState)

  sDefaultLimit = blnState
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BodyLimit

REM -------------------------------------------------------------------------
REM Inicio do Body e parametros
REM -------------------------------------------------------------------------
Public Sub BodyCenter(ByVal blnState)

  sDefaultCenter = blnState
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BodyCenter

REM -------------------------------------------------------------------------
REM Inicio do Body e parametros
REM -------------------------------------------------------------------------
Public Sub BodyEnd
  
  If sDefaultLimit Then
    ShowHTMLCR "</TD></TR></TABLE>"
    
  End If

  If sDefaultCenter Then
    ShowHTMLCR "</CENTER>"
    
  End If

REM  On Error Resume Next
  
REM  ZTISession.ZTI  
  
REM  ZTISession.AddPage
  
  ShowHTMLCR "</BODY>"
  ShowHTMLComment "Main Body Section End - ("
  
 REM & ZTISession.Sequence & "/" & ZTISession.PageCode & ") " & DateValue(Now)

  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BodyEnd

REM -------------------------------------------------------------------------
REM Todos os tags necessarios no HEAD
REM -------------------------------------------------------------------------
Public Sub HeadAll(ByVal strTitle)
  HeadBegin strTitle
  MetaDefault
  HeadEnd
  BaseFontDefault
  LinkDefault

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub HeadAll

REM -------------------------------------------------------------------------
REM Campos META Default
REM -------------------------------------------------------------------------
Public Sub MetaDefault

  ShowHTMLComment "Meta Section Begin"

  If Session("MetaPragma") > "" Then
    ShowMetaHTTPEquiv "PRAGMA", Session("MetaPragma")

  End If

  If Session("MetaRefresh") > "" Then
    ShowMetaHTTPEquiv "REFRESH", Session("MetaRefresh")

  End If

  If Session("MetaContentType") > "" Then
    ShowMetaHTTPEquiv "Content-Type", Session("MetaContentType")

  End If

  If Session("MetaReplyTo") > "" Then
    ShowMetaHTTPEquiv "REPLY-TO", Session("MetaReplyTo")

  End If

  If Session("MetaExpires") > "" Then
    ShowMetaHTTPEquiv "EXPIRES", Session("MetaExpires")
  End If

  If Session("MetaPicsLabel") > "" Then
    ShowMetaHTTPEquiv "PICS-Label", Session("MetaPicsLabel")

  End If

  If Session("MetaKeywords") > "" Then
    ShowMetaName "KEYWORDS", Session("MetaKeywords") & _
                             ", ASP, VBScript, Active Server Pages, ZTI, ZTITools, Zevallos, Zevallos Tecnologia em Informacao, Brasil, Brazil, Brasilia, DF"

  End If

  If Session("MetaKeyPhrase") > "" Then
    ShowMetaName "KEYPHASE", Session("MetaKeywords")

  End If

  If Session("MetaDescription") > "" Then
    ShowMetaName "DESCRIPTION", Session("MetaDescription")

  End If

  If Session("MetaProduct") > "" Then
    ShowMetaName "PRODUCT", Session("MetaProduct")

  End If

  If Session("MetaLocale") > "" Then
    ShowMetaName "LOCALE", Session("MetaLocale")

  End If

  If Session("MetaCharSet") > "" Then
    ShowMetaName "CHARSET", Session("MetaCharSet")

  End If

  If Session("MetaCategory") > "" Then
    ShowMetaName "CATEGORY", Session("MetaCategory")
  End If

  If Session("MetaGenerator") > "" Then
    ShowMetaName "GENERATOR", Session("MetaGenerator")

  End If

  If Session("MetaAuthor") > "" Then
    ShowMetaName "AUTHOR", Session("MetaAuthor")

  End If

  If Session("MetaDistribution") > "" Then
    ShowMetaName "DISTRIBUTION", Session("MetaDistribution")

  End If

  If Session("MetaFormatter") > "" Then
    ShowMetaName "FORMATTER", Session("MetaFormatter")

  End If

  If Session("MetaTechnology") > "" Then
    ShowMetaName "TECHNOLOGY", Session("MetaTechnology")

  End If

  ShowHTMLComment "Meta Section End"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MetaDefault

REM -------------------------------------------------------------------------
REM Campos Base Font Default
REM -------------------------------------------------------------------------
Public Sub BaseFontDefault
  If Session("BaseFont") > "" Then
    ShowHTMLCR "<BASEFONT FACE=""" & Session("BaseFont") & """>"

  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub BaseFontDefault

REM -------------------------------------------------------------------------
REM Campos Link Default
REM -------------------------------------------------------------------------
Public Sub LinkDefault

  ShowHTMLComment "Link Section Begin"

  If Session("LinkHomeHRef") > "" Then
    ShowLinkRel "Home", Session("LinkHomeHRef"), Session("LinkHomeTitle")

  End If

  If Session("LinkCopyrightHRef") > "" Then
    ShowLinkRel "Copyright", Session("LinkCopyrightHRef"), Session("LinkCopyrightTitle")

  End If
  
  If Session("LinkStyleSheetHRef") > "" Then
    ShowLinkRelStyleSheet Session("LinkStyleSheetHRef")

  End If

  If Session("LinkMadeHRef") > "" Then
    ShowLinkRev "Made", Session("LinkMadeHRef"), Session("LinkMadeTitle")

  End If

  ShowHTMLComment "Link Section End"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub LinkDefault

REM -------------------------------------------------------------------------
REM Cabecalho da Pagina
REM -------------------------------------------------------------------------
Public Sub PageHeaderDefault(ByVal strHeaderTitle)

  If strHeaderTitle > "" Then
    Session("HeaderTitle") = strHeaderTitle
  End If
  
  If Session("HeaderTitle") > "" Then
    ShowHTMLComment "Main Header Begin"
    ShowHTMLCR "<TABLE WIDTH=100% BORDER=0>"
    ShowHTMLCR "<TR>"
    
    If Session("HeaderIcon") > "" Then
      ShowHTMLCR "<TD WIDTH=""1%"">" & Session("HeaderIcon") & "</TD>"
      
    End If
    
    ShowHTMLCR "<TD ALIGN=CENTER>" & Session("HeaderTitle") & "</TD></TR>"
    ShowHTMLCR "</TABLE>"
    ShowHTMLComment "Main Header End"
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub PageHeaderDefault

REM -------------------------------------------------------------------------
REM Rodape da Pagina
REM -------------------------------------------------------------------------
Public Sub PageFooterDefault
  ShowHTMLCR "<br><HR>"
  ShowHTMLCR "<table width=100% border=" & Session("TableBorder") + 0 & " cellpadding=0 cellspacing=0>"
  ShowHTML "<tr>"

  If Session("SourceCode") Then
    ShowHTMLCR "<td width=1% align=left>"
    ShowHTMLCR "<A HREF=""/ZTITools/Code/code.asp?SourceFile=" & Request.ServerVariables("PATH_INFO") & """>"
    ShowHTMLCR "<IMG SRC=/ZTITools/Code/Vsource.GIF WIDTH=85 HEIGHT=45 ALT=""View Active Server Page Source"" ALIGN=CENTER BORDER=0></A>"
    ShowHTMLCR "</td>"
  End If

  ShowHTMLCR "<td align=center><font size=-3>"
  ShowHTMLCR Session("FooterTitle") & "<br>" & Session("FooterSugestions") & _
             "<br>" & Session("FooterCopyright")

  
  ShowHTMLCR "</font></td></tr></table>"
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub PageFooterDefault

REM -------------------------------------------------------------------------
REM Fim do Default.inc
%>