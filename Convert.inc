<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Convert.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Conversores
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

REM -------------------------------------------------------------------------
REM Lista para conversao
REM -------------------------------------------------------------------------
Function ConvList

  ConvList = "#160. .space;" ' space
  ConvList = ConvList & "#000.a.a;"   '
  ConvList = ConvList & "#000.b.b;"   '
  ConvList = ConvList & "#000.c.c;"   '
  ConvList = ConvList & "#000.d.d;"   '
  ConvList = ConvList & "#000.e.e;"   '
  ConvList = ConvList & "#000.f.f;"   '
  ConvList = ConvList & "#000.g.g;"   '
  ConvList = ConvList & "#000.h.h;"   '
  ConvList = ConvList & "#000.i.i;"   '
  ConvList = ConvList & "#000.j.j;"   '
  ConvList = ConvList & "#000.k.k;"   '
  ConvList = ConvList & "#000.l.l;"   '
  ConvList = ConvList & "#000.m.m;"   '
  ConvList = ConvList & "#000.n.n;"   '
  ConvList = ConvList & "#000.o.o;"   '
  ConvList = ConvList & "#000.p.p;"   '
  ConvList = ConvList & "#000.r.r;"   '
  ConvList = ConvList & "#000.s.s;"   '
  ConvList = ConvList & "#000.t.t;"   '
  ConvList = ConvList & "#000.u.u;"   '
  ConvList = ConvList & "#000.v.v;"   '
  ConvList = ConvList & "#000.w.w;"   '
  ConvList = ConvList & "#000.x.x;"   '
  ConvList = ConvList & "#000.y.y;"   '
  ConvList = ConvList & "#000.z.z;"   '
  ConvList = ConvList & "#000.A.A;"   '
  ConvList = ConvList & "#000.B.B;"   '
  ConvList = ConvList & "#000.C.C;"   '
  ConvList = ConvList & "#000.D.D;"   '
  ConvList = ConvList & "#000.E.E;"   '
  ConvList = ConvList & "#000.F.F;"   '
  ConvList = ConvList & "#000.G.G;"   '
  ConvList = ConvList & "#000.H.H;"   '
  ConvList = ConvList & "#000.I.I;"   '
  ConvList = ConvList & "#000.J.J;"   '
  ConvList = ConvList & "#000.K.K;"   '
  ConvList = ConvList & "#000.L.L;"   '
  ConvList = ConvList & "#000.M.M;"   '
  ConvList = ConvList & "#000.N.N;"   '
  ConvList = ConvList & "#000.O.O;"   '
  ConvList = ConvList & "#000.P.P;"   '
  ConvList = ConvList & "#000.Q.Q;"   '
  ConvList = ConvList & "#000.R.R;"   '
  ConvList = ConvList & "#000.S.S;"   '
  ConvList = ConvList & "#000.Y.Y;"   '
  ConvList = ConvList & "#000.U.U;"   '
  ConvList = ConvList & "#000.V.V;"   '
  ConvList = ConvList & "#000.W.W;"   '
  ConvList = ConvList & "#000.X.X;"   '
  ConvList = ConvList & "#000.Y.Y;"   '
  ConvList = ConvList & "#000.Z.Z;"   '
  ConvList = ConvList & "#000.0.0;"   '
  ConvList = ConvList & "#000.1.1;"   '
  ConvList = ConvList & "#000.2.2;"   '
  ConvList = ConvList & "#000.3.3;"   '
  ConvList = ConvList & "#000.4.4;"   '
  ConvList = ConvList & "#000.5.5;"   '
  ConvList = ConvList & "#000.6.6;"   '
  ConvList = ConvList & "#000.7.7;"   '
  ConvList = ConvList & "#000.8.8;"   '
  ConvList = ConvList & "#000.9.9;"   '
  ConvList = ConvList & "#160, ,nbsp;" 'no-break space
  ConvList = ConvList & "#000.|.brkbar;" ' broken vertical bar
  ConvList = ConvList & "#000._.mdash;"  ' em dash
  ConvList = ConvList & "#000.-.ndash;"  ' en dash
  ConvList = ConvList & "#000.>.gt;"     ' greater than
  ConvList = ConvList & "#000.<.lt;"     ' less than
  ConvList = ConvList & "#000.&.amp;"    ' ampersand
  ConvList = ConvList & "#000."".quot;" ' quotation mark
  ConvList = ConvList & "#000.'.acute;"  ' acute accent
  ConvList = ConvList & "#000.^.circ;"   ' circ accent
  ConvList = ConvList & "#000.~.tilde;"  ' tilde accent
  ConvList = ConvList & "#161.?.iexcl;"  ' inverted exclamation mark
  ConvList = ConvList & "#162.?.cent;"   ' cent sign
  ConvList = ConvList & "#163.#.pound;"  ' pound sterling sign
  ConvList = ConvList & "#164.$.curren;" ' general currency sign
  ConvList = ConvList & "#165.?.yen;"    ' yen sign
  ConvList = ConvList & "#166.?.brvbar;" ' broken (vertical) bar
  ConvList = ConvList & "#167.?.sect;"   ' section sign
  ConvList = ConvList & "#168.?.uml;"    ' umlaut (dieresis)
  ConvList = ConvList & "#169.?.copy;"   ' copyright sign
  ConvList = ConvList & "#170.?.ordf;"   ' ordinal indicator, feminine
  ConvList = ConvList & "#171.?.laquo;"  ' angle quotation mark, left
  ConvList = ConvList & "#172.?.not;"    ' not sign
  ConvList = ConvList & "#173.-.shy;"    ' soft hyphen
  ConvList = ConvList & "#174.?.reg;"    ' registered sign
  ConvList = ConvList & "#175.?.macr;"   ' macron
  ConvList = ConvList & "#176.?.deg;"    ' degree sign
  ConvList = ConvList & "#177.?.plusmn;" ' plus-or-minus sign
  ConvList = ConvList & "#178.?.sup2;"   ' superscript two
  ConvList = ConvList & "#179.?.sup3;"   ' superscript three
  ConvList = ConvList & "#180.'.acute;"  ' acute accent
  ConvList = ConvList & "#181.æ.micro;"  ' micro sign
  ConvList = ConvList & "#182.?.para;"   ' pilcrow (paragraph sign)
  ConvList = ConvList & "#183.?.middot;" ' middle dot
  ConvList = ConvList & "#184.,.cedil;"  ' cedilla
  ConvList = ConvList & "#185.?.sup1;"   ' superscript one
  ConvList = ConvList & "#186.?.ordm;"   ' ordinal indicator, masculine
  ConvList = ConvList & "#187.?.raquo;"  ' angle quotation mark, right
  ConvList = ConvList & "#188.?.frac14;" ' fraction one-quarter
  ConvList = ConvList & "#189.?.frac12;" ' fraction one-half
  ConvList = ConvList & "#190.?.frac34;" ' fraction three-quarters
  ConvList = ConvList & "#191.?.iquest;" ' inverted question mark
  ConvList = ConvList & "#192.À.Agrave;" ' capital A, grave accent
  ConvList = ConvList & "#193.Á.Aacute;" ' capital A, acute accent
  ConvList = ConvList & "#194.Â.Acirc;"  ' capital A, circumflex accent
  ConvList = ConvList & "#195.Ã.Atilde;" ' capital A, tilde
  ConvList = ConvList & "#196.Ä.Auml;"   ' capital A, dieresis or umlaut mark
  ConvList = ConvList & "#197.A.Aring;"  ' capital A, ring
  ConvList = ConvList & "#198.A.AElig;"  ' capital AE diphthong (ligature)
  ConvList = ConvList & "#199.Ç.Ccedil;" ' capital C, cedilla
  ConvList = ConvList & "#200.È.Egrave;" ' capital E, grave accent
  ConvList = ConvList & "#201.É.Eacute;" ' capital E, acute accent
  ConvList = ConvList & "#202.Ê.Ecirc;"  ' capital E, circumflex accent
  ConvList = ConvList & "#203.Ë.Euml;"   ' capital E, dieresis or umlaut mark
  ConvList = ConvList & "#204.Ì.Igrave;" ' capital I, grave accent
  ConvList = ConvList & "#205.Í.Iacute;" ' capital I, acute accent
  ConvList = ConvList & "#206.Î.Icirc;"  ' capital I, circumflex accent
  ConvList = ConvList & "#207.Ï.Iuml;"   ' capital I, dieresis or umlaut mark
  ConvList = ConvList & "#208.?.ETH;"    ' capital Eth, Icelandic
  ConvList = ConvList & "#209.Ñ.Ntilde;" ' capital N, tilde
  ConvList = ConvList & "#210.Ò.Ograve;" ' capital O, grave accent
  ConvList = ConvList & "#211.Ó.Oacute;" ' capital O, acute accent
  ConvList = ConvList & "#212.Ô.Ocirc;"  ' capital O, circumflex accent
  ConvList = ConvList & "#213.Õ.Otilde;" ' capital O, tilde
  ConvList = ConvList & "#214.Ö.Ouml;"   ' capital O, dieresis or umlaut mark
  ConvList = ConvList & "#215.*.times;"  ' multiply sign
  ConvList = ConvList & "#216.O.Oslash;" ' capital O, slash
  ConvList = ConvList & "#217.Ù.Ugrave;" ' capital U, grave accent
  ConvList = ConvList & "#218.Ú.Uacute;" ' capital U, acute accent
  ConvList = ConvList & "#219.Û.Ucirc;"  ' capital U, circumflex accent
  ConvList = ConvList & "#220.Ü.Uuml;"   ' capital U, dieresis or umlaut mark
  ConvList = ConvList & "#221.Ý.Yacute;" ' capital Y, acute accent
  ConvList = ConvList & "#222.?.THORN;"  ' capital THORN, Icelandic
  ConvList = ConvList & "#223.?.szlig;"  ' small sharp s, German (sz ligature)
  ConvList = ConvList & "#224.à.agrave;" ' small a, grave accent
  ConvList = ConvList & "#225.á.aacute;" ' small a, acute accent
  ConvList = ConvList & "#226.â.acirc;"  ' small a, circumflex accent
  ConvList = ConvList & "#227.ã.atilde;" ' small a, tilde
  ConvList = ConvList & "#228.ä.auml;"   ' small a, dieresis or umlaut mark
  ConvList = ConvList & "#229.a.aring;"  ' small a, ring
  ConvList = ConvList & "#230.a.aelig;"  ' small ae diphthong (ligature)
  ConvList = ConvList & "#231.ç.ccedil;" ' small c, cedilla
  ConvList = ConvList & "#232.è.egrave;" ' small e, grave accent
  ConvList = ConvList & "#233.é.eacute;" ' small e, acute accent
  ConvList = ConvList & "#234.ê.ecirc;"  ' small e, circumflex accent
  ConvList = ConvList & "#235.ë.euml;"   ' small e, dieresis or umlaut mark
  ConvList = ConvList & "#236.ì.igrave;" ' small i, grave accent
  ConvList = ConvList & "#237.í.iacute;" ' small i, acute accent
  ConvList = ConvList & "#238.î.icirc;"  ' small i, circumflex accent
  ConvList = ConvList & "#239.ï.iuml;"   ' small i, dieresis or umlaut mark
  ConvList = ConvList & "#240.?.eth;"    ' small eth, Icelandic
  ConvList = ConvList & "#241.ñ.ntilde;" ' small n, tilde
  ConvList = ConvList & "#242.ò.ograve;" ' small o, grave accent
  ConvList = ConvList & "#243.ó.oacute;" ' small o, acute accent
  ConvList = ConvList & "#244.ô.ocirc;"  ' small o, circumflex accent
  ConvList = ConvList & "#245.õ.otilde;" ' small o, tilde
  ConvList = ConvList & "#246.ö.ouml;"   ' small o, dieresis or umlaut mark
  ConvList = ConvList & "#247.?.divide;" ' divide sign
  ConvList = ConvList & "#248.?.oslash;" ' small o, slash
  ConvList = ConvList & "#249.ù.ugrave;" ' small u, grave accent
  ConvList = ConvList & "#250.ú.uacute;" ' small u, acute accent
  ConvList = ConvList & "#251.û.ucirc;"  ' small u, circumflex accent
  ConvList = ConvList & "#252.ü.uuml;"   ' small u, dieresis or umlaut mark
  ConvList = ConvList & "#253.ý.yacute;" ' small y, acute accent
  ConvList = ConvList & "#254.?.thorn;"  ' small thorn, Icelandic
  ConvList = ConvList & "#255.?.yuml;"   ' small y, dieresis or umlaut mark

End Function
REM -------------------------------------------------------------------------
REM Final da Function ConvList

REM -------------------------------------------------------------------------
REM Converte os Acentos ANSI Latin 1 para Acentos ASCII
REM -------------------------------------------------------------------------
Public Function ConvLatin12Ascii(ByVal strString)
  Dim strConvList
  Dim strResult
  Dim intPos
  
  strResult = ""
  
  strConvList = ConvList

  strString = Left(strString, 1) & LCase(Mid(strString, 2))

  intPos = InStr(strConvList, strString)

  If intPos > 0 Then
    If Left(strString, 1 ) = "#" Then
      intPos = intPos + 5

    Else
      intPos = intPos - 2
      
    End If

    strResult = Mid(strConvList, intPos, 1)

  End If
 
  ConvLatin12Ascii = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ConvLatin12Ascii

REM -------------------------------------------------------------------------
REM Converte os Acentos ANSI Latin 1 para Acentos ASCII
REM -------------------------------------------------------------------------
Public Function ConvAscii2Latin1(ByVal strChar)
  Dim strConvList
  Dim strResult
  Dim intPos, intRef

  strResult = ""
  
  strConvList = ConvList

  intPos = InStr(strConvList, "." & strChar & ".") + 2
  
  intRef = InStr(";.?", strChar)
  
  If intPos > 0 And intRef = 0 Then
    strResult = Mid(strConvList, intPos + 1, 10)
    
    intPos = InStr(strResult, ";")
    
    strResult = Left(strResult, intPos - 1)
  End If
  
  ConvAscii2Latin1 = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ConvAscii2Latin1

REM -------------------------------------------------------------------------
REM Limpa o os brancos do inicio e do fim
REM -------------------------------------------------------------------------
Public Function AllTrim(ByVal strString)
  AllTrim = Trim(LTrim(strString))

End Function
REM -------------------------------------------------------------------------
REM Final da Function AllTrim

REM -------------------------------------------------------------------------
REM Converte os Acentos ANSI Latin 1 para Acentos ASCII
REM -------------------------------------------------------------------------
Public Function ConvFirstCaps(ByVal strFrase)
  Dim strResult
  Dim intPos
  Dim strPiece
  Dim intFraseLen
  Dim blnNotEnd

  strFrase = AllTrim(LCase(strFrase))
  
  strResult = ""

  intPos = 1
  
  blnNotEnd = True
  
  Do While intPos > 0 And strFrase > "" And blnNotEnd
    intPos      = InStr(strFrase, " ")
    intFraseLen = Len(strFrase)
  
    If intPos = 0 Then
      intPos = intFraseLen
      
      blnNotEnd = False
    
    End If

    If intPos > 0 Then
      strPiece = Left(strFrase, intPos)
      
      If Len(strPiece) < 5 Then
        If InStr("~da~de~do~das~dos~um~uma~uns~umas~em~e~a~em~não~sim~", "~" & Trim(LCase(strPiece)) & "~") > 0 Then
          strPiece = LCase(strPiece)
          
        ElseIf InStr("~ha~há~rio~sao~são~av~alm~rua~rod~est~r~pq~", "~" & LCase(strPiece) & "~") > 0 Or InStr(strPiece, ".") > 0 Then
          strPiece = UCase(Left(strPiece, 1)) & Mid(strPiece, 2, intPos - 1)
          
        End If

      Else
        strPiece = UCase(Left(strPiece, 1)) & Mid(strPiece, 2, intPos - 1)
          
      End If
      
      strResult = strResult & strPiece

      strFrase = AllTrim(Mid(strFrase, intPos + 1))

    End If

  Loop
  
  ConvFirstCaps = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function ConvFirstCaps

REM -------------------------------------------------------------------------
REM Fim do Convert.inc
%>
