<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Initializer.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do Inicializador
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

Option Explicit 
Response.Expires = 0

REM -------------------------------------------------------------------------
REM Constantes dos Tipos de Conexoes
REM -------------------------------------------------------------------------
Const conConnSQL     = 1
Const conConnDBase   = 2
Const conConnAccess  = 3
Const conConnExcel   = 4
Const conConnFoxPro  = 5
Const conConnText    = 6
Const conConnParadox = 7
      
Const conConnDriverAccess   = 25 
Const conConnDriverdBaseIII = 21 
Const conConnDriverdBaseIV  = 277
Const conConnDriverdBase5   = 533
Const conConnDriverExcel3   = 534
Const conConnDriverExcel4   = 278
Const conConnDriverExcel5   = 22 
Const conConnDriverExcel7   = 22 
Const conConnDriverExcel97  = 790
Const conConnDriverFoxPro20 = 24 
Const conConnDriverFoxPro25 = 280
Const conConnDriverFoxPro26 = 536
Const conConnDriverParadox3 = 26 
Const conConnDriverParadox4 = 282
Const conConnDriverParadox5 = 538
Const conConnDriverText     = 27 

Dim sobjBT
Set sobjBT = Server.CreateObject("MSWC.BrowserType")

Dim sblnVDHTML, sstrBBrowser, sintBVersion
sblnVDHTML = False

sstrBBrowser = LCase(RTrim(sobjBT.browser))
sintBVersion = sobjBT.version + 0

SystemOptions

If sstrBBrowser = "ie" And sintBVersion >= 4.0 Then 
  sblnVDHTML = True
  
End If

If Not Session("TableBorder") > 0 Then
  Session("TableBorder") = 0
  
End If

REM -----------------------------------------------------------------------
REM Host Variables
REM -----------------------------------------------------------------------
Public sstrThisScriptURL, sstrThisScriptName, sstrThisHTTPHost, sstrThisScriptFull
Public sstrThisSiteRootdir, sstrThisQueryString

sstrThisHTTPHost    = LCase(Request.ServerVariables("HTTP_HOST"))
sstrThisScriptName  = LCase(Request.ServerVariables("SCRIPT_NAME"))

sstrThisScriptURL   = LCase("http://" & sstrThisHTTPHost & sstrThisScriptName)

sstrThisSiteRootDir = LCase(Request.ServerVariables("APPL_PHYSICAL_PATH"))

sstrThisScriptFull  = LCase(Request.ServerVariables("PATH_TRANSLATED"))

sstrThisQueryString = Request.ServerVariables("QUERY_STRING")

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
REM DSN	     = The name of the data source as returned by SQLDataSources. The 
REM            DSN keyword is not used if DRIVER is used.
REM DRIVER 	 = The name of the driver as returned by SQLDrivers. The DRIVER 
REM            keyword is not used if DSN is used.The SQL Server driver name
REM            is {SQL Server}.
REM SERVER	 = The name of the server on the network on which the data source 
REM            resides. On a Microsoft Windows NT computer, "(local)" can be 
REM            entered as the server, in which case a local copy of SQL Server
REM            can be used, even when this is a non-networked version. Note 
REM            that when the 16-bit SQL Server driver is using "(local)" without
REM            a network, the "MS Loopback Adapter" must be installed.
REM UID	     = The user login ID.
REM PWD	     = The user-specified password.
REM APP	     = The name of the application calling SQLDriverConnect (optional).
REM WSID	   = The workstation ID. Typically, this is the network name of the 
REM            computer on which the application resides (optional).
REM DATABASE = The name of the SQL Server database (optional).
REM LANGUAGE = The national language to be used by SQL Server (optional).
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Stored                Procedure	Returns
REM -------------------------------------------------------------------------
REM sp_column_privileges	Information about column privileges for the specified 
REM                       table(s)
REM sp_columns	          Information about columns for the specified table(s)
REM sp_databases	        A list of databases
REM sp_datatype_info      Information about the supported datatypes
REM sp_fkeys	            Information about logical foreign keys
REM sp_pkeys	            Information about primary keys
REM sp_server_info	      A list of attribute names and matching values for the 
REM                       server
REM sp_special_columns	  Information for a single table about columns in the 
REM                       table that have special attributes
REM sp_sproc_columns	    Column information for a stored procedure
REM sp_statistics	        A list of indexes for a single table
REM sp_stored_procedures	A list of stored procedures
REM sp_table_privileges	  Information about table privileges for the specified
REM                       table(s)
REM sp_tables	            A list of objects that can be queried
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Direct Data Access
REM -------------------------------------------------------------------------
REM Driver	            Keywords Required	Examples
REM -------------------------------------------------------------------------
REM Microsoft Access	  Driver, DBQ	      Driver={Microsoft Access Driver (*.mdb)};
REM                                       DBQ=c:\temp\sample.mdb
REM dBASE	              Driver, DriverID	Driver={Microsoft dBASE Driver (*.dbf)};
REM                                       DBQ=c:\temp; DriverID=277
REM Microsoft Excel 3.0	Driver, DriverID	Driver={Microsoft Excel Driver (*.xls)};
REM or 4.0                                DBQ=c:\temp; DriverID=278
REM Microsoft Excel	    Driver, DriverID, Driver={Microsoft Excel Driver (*.xls)};
REM  5.0/7.0            DBQ	              DBQ=c:\temp\sample.xls; DriverID=22
REM Microsoft Excel 97	Driver, DriverID, Driver={Microsoft Excel Driver (*.xls)};
REM                     DBQ               DBQ=c:\temp\sample.xls; DriverID=790
REM Microsoft FoxPro	  Driver, DriverID	Driver={Microsoft FoxPro Driver (*.dbf)};
REM                                       DBQ=c:\temp; DriverID=536
REM Paradox	            Driver, DriverID	Driver={Microsoft Paradox Driver (*.db )};
REM                                       DBQ=c:\temp; DriverID=26
REM Text	              Driver	          Driver={Microsoft Text Driver (*.txt;*.csv)};
REM                                       DefaultDir=c:\temp
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM DRIVERID	An integer ID for the driver.
REM -------------------------------------------------------------------------
REM 25        (Microsoft Access)
REM 21        (dBASE III)
REM 277       (dBASE IV)
REM 533       (dBASE 5.0)
REM 534       (Microsoft Excel 3.0)
REM 278       (Microsoft Excel 4.0)
REM 22        (Microsoft Excel 5.0/7.0)
REM 790       (Microsoft Excel 97)
REM 24        (Microsoft FoxPro 2.0)
REM 280       (Microsoft FoxPro 2.5)
REM 536       (Microsoft FoxPro 2.6)
REM 26        (Paradox 3.x)
REM 282       (Paradox 4.x)
REM 538       (Paradox 5.x)
REM 27        (Text)

REM -------------------------------------------------------------------------
REM Convencao para nomeclatura de variaveis
REM -------------------------------------------------------------------------
REM Subtype     Prefix  Example 
REM -------------------------------------------------------------------------
REM Boolean     bln     blnFound 
REM Byte        byt     bytRasterData 
REM Date (Time) dtm     dtmStart 
REM Double      dbl     dblTolerance 
REM Error       err     errOrderint 
REM Integer     int     intQuantity 
REM Long        lng     lngDistance 
REM Object      obj     objCurrent 
REM Single      sng     sngAverage 
REM String      str     strFirstName 

REM Constant    con     conAppName 
REM Binary      bin     binImage

REM -------------------------------------------------------------------------
REM Escopo das variaveis
REM -------------------------------------------------------------------------
REM Scope           Prefix  Example 
REM -------------------------------------------------------------------------
REM Procedure-level None    dblVelocity 
REM Script-level    s       sblnCalcInProgress 

REM -------------------------------------------------------------------------
REM Convencao para nomeclatura de objetos
REM -------------------------------------------------------------------------
REM Object type         Prefix   Example 
REM -------------------------------------------------------------------------
REM 3D Panel                pnl     pnlGroup 
REM Animated button         ani     aniMailBox 
REM Check box               chk     chkReadOnly 
REM Combo box, 
REM drop-down list box      cbo     cboEnglish 
REM Command button          cmd     cmdExit 
REM Common dialog           dlg     dlgFileOpen 
REM Frame                   fra     fraLanguage 
REM Horizontal scroll bar   hsb     hsbVolume 
REM Image                   img     imgIcon 
REM Label                   lbl     lblHelpMessage 
REM Line                    lin     linVertical 
REM List Box                lst     lstPolicyCodes 
REM Spin                    spn     spnPages 
REM Text box                txt     txtLastName 
REM Vertical scroll bar     vsb     vsbRate 
REM Slider                  sld     sldScale 
REM Parameter               par     parConvertionSource

REM Radio Button            rad     radSourceType

REM ----------------------------------------------------------------------------
REM  Sistema de padronizacao da documentacao
REM ----------------------------------------------------------------------------
REM 
REM  - [ colchetes ]      = os itens dentro de colchetes sao opcionais. Para a
REM                       inclucao de informacoes opcionais digite somente a
REM                       informacao entre os colchetes; nao digite os colchetes;
REM 
REM  - ...( reticencias ) = as reticencias ( ... ) significam que um item podera
REM                         ser repetido tantas vezes quanto forem necessarias;
REM 
REM  - separadores      = a menos que seja especificado o contrario, os comandos
REM                       devem ser separados das opcoes utilizando espacos.
REM 
REM  - Nomes com as iniciais maiusculas, bem detalhados, utilizando o bom senso;
REM      Ex: SaveScreen() ( mais facil de utilizar e vizualizar )
REM 
REM  - Cabecalho seguindo os passos abaixo:
REM 
REM ----------------------------------------------------------------------------
REM  Nome_da_Funcao()
REM ----------------------------------------------------------------------------
REM  Proposito  : Descri‡ao curta de onde pode-se utilizar a funcao.
REM 
REM  Sintaxe    : Prototipo completo. Mostrando os argumentos e tambem as
REM               declaracoes.
REM 
REM  Parƒmetros : Detalhamento de cada um dos parametros, Descricao do seu
REM               formato, possiveis acoes, enfim, tudo sobre o mesmo.
REM 
REM  Assume     : Situacao que a fun‡ao assume ja existir quando da sua
REM               execucao.
REM 
REM  Efeitos    : Efeitos que a funcao podera causar no meio ambiente.
REM 
REM  Chamadas   : Que outras procedures e Funcoes a mesma necessita.
REM 
REM  Autor      : O nome do autor ou dos autores da funcao.
REM 
REM  Criacao    : Hora e data da Criacao.
REM 
REM  Alteracao  : Hora e data da ultima Alteracao.
REM 
REM  Descricao  : Como a funciona a fun‡ao e como devera utiliza-la.
REM 
REM  Usos Comuns: Situacoes em que a funcao podera ser util.
REM 
REM  Retornos   : O valor retornado pela funcao, se este existir.
REM 
REM  Comentarios: Observacoes especiais para ajudar evitar confusoes no uso
REM               efetivo da mesma.
REM 
REM  Veja Tambem: Outras Funcoes relacionadas e como estas se relacionam.
REM 
REM  Exemplos   : Um ou mais programas de exemplos completos ilustrando como a
REM               funcao pode ser utilizada.
REM ----------------------------------------------------------------------------
REM 
REM  PADRONIZACAO DOS PROGRAMAS
REM 
REM  - Comandos com todas as letras maiusculas
REM      Ex: WHILE, END, IF, ELSE, ENDIF, FOR, NEXT, DO CASE, ENDCASE etc
REM 
REM  - Colocar em () todos os parametros dos comandos WHILE, IF, CASE etc
REM      Ex: WHILE ( pTRUE )
REM            DO CASE
REM              CASE ( !EofRel() )
REM                IF ( Empty( var ) )
REM                  cVar := "X"
REM                ENDIF
REM 
REM              OTHERWISE
REM                GOTO BOTTOM
REM            ENDCASE
REM          END
REM 
REM  - Funcoes somente com as iniciais maiusculas
REM      Ex: SubStr(), Left(), Chr(), Iif() etc
REM 
REM  - Identacao sempre com dois espacos, sem excecao
REM      Ex: WHILE ( pTRUE )
REM            IF ( Empty( var ) )
REM              cVar := "X"
REM            ENDIF
REM          END
REM 
REM  - Linha de programa, com no maximo 78 posicoes, utilizando o ";" sempre
REM    que necessario.
REM      Ex: IF ( Trim( Str( nNumero, 5, 2 ) ) == "123" .AND.;
REM            SubStr( Str( nNumero2, 5, 2 ), 2 ) == "123" )
REM            lErro := pFALSE
REM          ENDIF
REM 
REM  - Separar com um espaco todos os argumentos das expressoes aritmeticas e
REM    fun‡oes, sem exce‡ao.
REM      Ex: cAux := Str( nA + nB + Val( nX, 2 ), 3 )
REM 
REM  - Todas as letras das variaveis locais com as iniciais maiusculas
REM    precedidas com uma letra minuscula indicando o tipo do dado, utilizando
REM    a notacao Hungara
REM      Ex: cCaracter  := "XXX"
REM          nNumero    := 0
REM          lLogico    := pTRUE
REM          dData      := Date()
REM          aArray     := { { "xxx" }, { 120 } }
REM          oObjeto    := Janela:New()
REM          bCodeBlock := {| cParametro |, Codigo( cParametro ) }
REM 
REM  - Todas as variaveis nao locais, com as iniciais maiusculas precedidas
REM    com uma letra minuscula indicando o tipo da variavel e do dado, 
REM    utilizando a notacao Hungara.
REM      Ex: pcCaracter  := "XXX"                                   -> Public
REM          snNumero    := 0                                       -> Static
REM          plLogico    := pTRUE                                   -> Public
REM          sdData      := Date()                                  -> Static
REM          saArray     := { { "xxx" }, { 120 } }                  -> Static
REM          poObjeto    := Janela:New()                            -> Public
REM          sbCodeBlock := {| cParametro |, Codigo( cParametro ) } -> Static
REM 
REM  - As variaveis de trabalho dos arquivos, com todas as letras minusculas e
REM    sempre precedidas pelo o identificador de variavel de memoria "M->",
REM    precedidas com uma letra minuscula indicando o tipo do dado.
REM      Ex: M->cCampo, M->cNome, M->cEndereco
REM      minusculo, indicando que e' um parametro.
REM 
REM  - As variaveis locais, mas de uso por outras procedures e Funcoes, exceto
REM    as variaveis de trabalho dos arquivos, deverao estar precedidas por um
REM    identificador da funcao que gerou.
REM      Ex: PS_Var -> variavel da funcao PrtScr()
REM ----------------------------------------------------------------------------

If Not Session("Default") > "" Then
  InitDefault
  
End If

REM -------------------------------------------------------------------------
REM Inicialisa todos os padroes do sistema
REM -------------------------------------------------------------------------
Public Sub InitDefault
  Session("Default") = "ZTI"
  
  Session("MetaPragma")       = "NO-CACHE"
  Session("MetaContentType")  = "text/html; charset=iso-8859-1"
  Session("MetaReplyTo")      = "webmaster@zevallos.com.br"
  Session("MetaExpires")      = now + 90
  Session("MetaPicsLabel")    = "(PICS-1.1 ""http://www.rsac.org/ratingsv01.html""" &  _ 
                                " l gen true comment ""RSACi North America Server"" by" & _ 
                                " ""webmaster@zevallos.com.br"" for """ & _
                                "http://www.zevallos.com.br""" & _
                                " on ""1997.06.26T21:24-0500"" r (n 0 s 0 v 0 l 0))"
  Session("MetaKeywords")     = "Zevallos"
  Session("MetaKeyPhrase")    = ""  
  Session("MetaDescription")  = ""
  Session("MetaProduct")      = "ZTITools v2.05 - Biblioteca ASP"
  Session("MetaLocale")       = "PO-BR"
  Session("MetaCharSet")      = "US-ASCII"
  Session("MetaCategory")     = "HOME PAGE"
  Session("MetaGenerator")    = "Tecnologias da Zevallos em ASP"
  Session("MetaAuthor")       = "ZTI - Zevallos(r) Tecnologia em Informacao - Brasilia - DF -" & _
                                " Brasil - webmaster@zevallos.com.br - http://www.zevallos.com.br"
  Session("MetaDistribution") = "global"
  Session("MetaFormatter")    = "ZTITools v2.05"
  Session("MetaTechnology")   = "SCRIPTING"

REM -------------------------------------------------------------------------
REM Constantes BASE FONT de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("BaseFont") = "Arial, Helvetica, Sans-Serif"

REM -------------------------------------------------------------------------
REM Constantes LINK de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("LinkHomeHRef")       = "/Default.asp"
  Session("LinkHomeTitle")      = "ZTITools v1.0"

  Session("LinkCopyrightHRef")  = "/copyright.htm"
  Session("LinkCopyrightTitle") = "Copyright da Zevallos"

  Session("LinkMadeHRef")       = "mailto:webmaster@zevallos.com.br"
  Session("LinkMadeTitle")      = "WebMaster da Zevallos"

  Session("LinkStyleSheetHRef") = ""
  Session("LinkStyleSheetTitle") = ""
  
REM -------------------------------------------------------------------------
REM Constantes BODY de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("BodyBGColor")    = "white"
  Session("BodyBackground") = "/ZTITools/img/Assinatura75Back.Gif"
  Session("BodyBGProperties") = ""
  Session("BodyStyle")      = ""
  Session("BodyText")       = ""
  Session("BodyLink")       = ""
  Session("BodyVLink")      = ""
  Session("BodyALink")      = ""
  Session("BodyTopMargin")  = 0
  Session("BodyLeftMargin") = 0
  Session("BodyWidth")      = 780

REM -------------------------------------------------------------------------
REM Constantes HEADER de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("HeaderIcon")       = "<IMG SRC=/ZTITools/img/Assinatura75.Gif>"
  Session("HeaderTitle")      = "<h2></h2>"

REM -------------------------------------------------------------------------
REM Constantes FOOTER de configuracao do sistema
REM -------------------------------------------------------------------------
  Session("FooterTitle")      = "Um produto da <A HREF=http://www.zevallos.com.br>" & _
                                "Zevallos&reg;</A> Tecnologia em Informa&ccedil;&atilde;o."
  Session("FooterSugestions") = "Sugest&otilde;es encaminhar para o " & _
                                "<A HREF=mailto:webmaster@zevallos.com.br>" & _
                                "<img src=/ztitools/img/mailto.gif alt=""Mail To"" border=0> " & _
                                "<i>&lt;webmaster@zevallos.com.br&gt;</i></A>"
  Session("FooterCopyright")  = "&copy; 1997, 1998, 1999 <A HREF =/copyright.asp>Zevallos&reg;</a>" & _
                                " todos os direitos reservados."

End Sub
REM -------------------------------------------------------------------------
REM Fim da Sub InitDefault

REM -------------------------------------------------------------------------
REM Inicializador do sistema
REM -------------------------------------------------------------------------
Private Sub SystemOptions
  REM -----------------------------------------------------------------------
  REM Inicializa os defaults do sistema
  REM -----------------------------------------------------------------------
  If Request.QueryString("Default") > "" Then
    Session("Default") = ""

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
REM Salva a QueryString Corrente
REM -------------------------------------------------------------------------
Public Sub QueryStringSave
  If sstrThisScriptName & "?" & sstrThisQueryString <> Session("QueryStringCurrent") Then
    Session("QueryStringLast") = Session("QueryStringCurrent")
    Session("QueryStringCurrent") = sstrThisScriptName & "?" & sstrThisQueryString

  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub QueryStringSave

REM -------------------------------------------------------------------------
REM Recupera a QueryString Corrent
REM -------------------------------------------------------------------------
Public Function QueryStringCurrent
  QueryStringCurrent = Session("QueryStringCurrent")

End Function
REM -------------------------------------------------------------------------
REM Final da Function QueryStringSave

REM -------------------------------------------------------------------------
REM Recupera a QueryString Anterior
REM -------------------------------------------------------------------------
Public Function QueryStringLast
  QueryStringLast = Session("QueryStringLast")

End Function
REM -------------------------------------------------------------------------
REM Final da Function QueryStringLast

REM -------------------------------------------------------------------------
REM Fim do Initializer.inc
%>