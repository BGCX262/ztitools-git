<%@ LANGUAGE="VBSCRIPT" %>
<%
REM -------------------------------------------------------------------------
REM  /Cadastro/Inicia.ASP
REM -------------------------------------------------------------------------
REM  Descricao   : Cria o ambiente inicial do Cadastro Fecomercio
REM  Cria‡ao     : 1/17/99 11:32PM
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------
%>
<!--#INCLUDE VIRTUAL="/ZTITools/All.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/Data.inc"-->
<%

REM -------------------------------------------------------------------------
REM Constantes do sistema
REM -------------------------------------------------------------------------
  Const conAppName = "ZEUSGenerator"

  Const conScriptTimeout = 300

REM -----------------------------------------------------------------------
REM Objetos
REM -----------------------------------------------------------------------
  Public sobjConn
  Public sobjDBase
  Public sobjRS, sobjRS2
  Public sintOldTimeOut
  Public sobjFS
  Public sobjCMD
  Public sobjFile
REM -----------------------------------------------------------------------
REM Diversas
REM -----------------------------------------------------------------------
  Public sblnAtividade
  Public sblnSetCookies, sblnGetCookies

  sblnSetCookies = False
  sblnGetCookies = True

  sblnAtividade = False

REM -----------------------------------------------------------------------
REM Parametros
REM -----------------------------------------------------------------------
  Public sparOption

  sparOption  = lCase(Request.QueryString("O"))

REM -------------------------------------------------------------------------
REM Chama a Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Main

REM -------------------------------------------------------------------------
REM Rotina Principal do Sistema
REM -------------------------------------------------------------------------
Private Sub Main
  Server.ScriptTimeOut = conScriptTimeout

  Set sobjConn      = Server.CreateObject("ADODB.Connection")
  sobjConn.ConnectionTimeout = Session("ConnectionTimeout")
  sobjConn.CommandTimeout    = Session("CommandTimeout")

  sobjConn.Open Session("ConnectionString"), _
                Session("RuntimeUserName"), _
                Session("RuntimePassword")

  Dim strPubPath

  Set sobjRS = Server.CreateObject("ADODB.RecordSet")
  sobjRS.CacheSize = 150
  sobjRS.CursorType = adOpenDynamic
  sobjRS.LockType = adLockPessimistic

  Set sobjRS2 = Server.CreateObject("ADODB.RecordSet")
  sobjRS2.CacheSize = 150
  sobjRS2.CursorType = adOpenDynamic
  sobjRS2.LockType = adLockPessimistic

  Set sobjFS = CreateObject("Scripting.FileSystemObject")

  Set sobjCMD = Server.CreateObject("ADODB.Command")
  Set sobjCMD.ActiveConnection = sobjConn
  sobjCMD.Prepared = True
  sobjCMD.CommandType = adCmdText

  MainBody

  Server.ScriptTimeOut = Session("ScriptTimeOut")

  On Error Resume Next

  Set sobjDbase = nothing

  sobjConn.Close
  Set sobjConn = nothing

  sobjCMD.Close
  Set sobjCMD = nothing

  sobjRS.Close
  Set sobjRS   = nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Main

REM -------------------------------------------------------------------------
REM  Cria os Programas das Tabelas
REM -------------------------------------------------------------------------
Public Sub CreateTablePrograms
  Dim strFile, strConnectionString
  Dim i, x, arrTable(20), sql
  Dim intType, intDriver
  Dim objFS
  Dim strLocalDir
  Dim intNameTemp, intTypeTemp, intSizeTemp
  Dim intNameLen, intTypeLen, intSizeLen
  Dim intFileCounter
  Dim strFormList, strFormUnit, strFormFind
  Dim strResultFile
  Dim strDBDir, strRoot, strASPDir, strCurrentFile
  Dim intFieldLen, strFieldType
  
  Dim strAux

  intFileCounter = 14

REM  strDBDir = "BasedeDados\"
REM  strASPDir = "Data\"
REM  strFile = "GeapNew.mdb"
REM  strRoot = "D:\Inetpub\WWWRoot\Clients\Geap\"

  strDBDir = "BD\"
  strASPDir = "Data\"
  strFile = "Base3.mdb"
  strRoot = "D:\Inetpub\WWWRoot\Clients\Loreno\Sac"

  strResultFile = strRoot & strASPDir & "\Saida.asp"

  Dim sobjZTIConn

  Set sobjZTIConn = Server.CreateObject("ZTITools.Connection")

  sobjZTIConn.ConnectionType = conConnAccess
  sobjZTIConn.DBQ = strRoot & strDBDir & strFile

REM  arrTable(1) = "E_PROD"
REM  arrTable(2) = "E_CLAS"
REM  arrTable(3) = "V_CLI"
REM  arrTable(4) = "V_ITEM"
REM  arrTable(5) = "V_ORDM"
REM  arrTable(6) = "V_PLPG"
REM  arrTable(7) = "ARECEBER"

REM  arrTable(8) = "MOEDA"
REM  arrTable(9) = "QTDPRDLJ"
REM  arrTable(10) = "RECEBIDO"
REM  arrTable(11) = "TIPOMOVI"
REM  arrTable(12) = "TIPOPGTO"
REM  arrTable(14) = "V_VEND"
REM  arrTable(15) = "LOJAS"

REM  arrTable(1) = "slClientes"
REM  arrTable(2) = "slContasAReceber"
REM  arrTable(3) = "slPedidoCabecalho"
REM  arrTable(4) = "slPedidoItens"
REM  arrTable(5) = "slPlanosPagamento"
REM  arrTable(6) = "slProdutos"
REM  arrTable(7) = "slProdutosClasse"
REM  arrTable(8) = "slProdutosSaldo"

REM  arrTable(1) = "D010_TELEFONE_TITULAR"
REM  arrTable(2) = "D011_SITUACAO_CONVENIADA"
REM  arrTable(3) = "D012_CONVENIADA_TITULAR"
REM  arrTable(4) = "D013_SUSPENSAO_CLIENTE"
REM  arrTable(5) = "D020_INSCRICAO"
REM  arrTable(6) = "D021_CARTEIRA_CLIENTE"
REM  arrTable(7) = "D027_CONVENIADA"
REM  arrTable(8) = "D030_CLIENTE"
REM  arrTable(9) = "D034_OBS_CLIENTE"
REM  arrTable(10) = "D037_VINCULO"
REM  arrTable(11) = "D038_ORGAO_LOCAL"
REM  arrTable(12) = "D044_CIDADE"
REM  arrTable(13) = "D059_PLANO"
REM  arrTable(14) = "D089_CONTROLE_CLIENTE"

  arrTable(1) = "slClientes"
  arrTable(2) = "slContasAReceber"
  arrTable(3) = "slLojas"
  arrTable(4) = "slMoeda"
  arrTable(5) = "slPedidoCabecalho"
  arrTable(6) = "slPedidoItens"
  arrTable(7) = "slPlanosPagamento"
  arrTable(8) = "slProdutos"
  arrTable(9) = "slProdutosClasse"
  arrTable(10) = "slProdutosSaldo"
  arrTable(11) = "slVendedor"
  arrTable(12) = "slTipoMovimentacao"
  arrTable(13) = "slTipoPagamento"

  REM -----------------------------------------------------------------------
  Set objFS   = CreateObject("Scripting.FileSystemObject")

  Set sobjFile = objFS.CreateTextFile(strResultFile, True)

  ShowMessageError "Criando o arquivo " & strResultFile

  REM -----------------------------------------------------------------------
  FSWriteLineCR "<" & "%"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM  /" & strResultFile
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM  Descricao   : Includes gerais da " & strResultFile
  FSWriteLineCR "REM  Criacao     : " & Now
  FSWriteLineCR "REM  Local       : Brasilia/DF"
  FSWriteLineCR "REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>"
  FSWriteLineCR "REM  Versao      : 1.0.0"
  FSWriteLineCR "REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "%" & ">"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/All.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/Validate.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/Edit.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/Browse.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/Security.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/EditForm.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/Find.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/ZTITools/Data.inc""-->"
  FSWriteLineCR "<!--#INCLUDE VIRTUAL=""/Data/sgGeapData.inc""-->"

  ShowMessageError "Gerada a Sessão dos Includes"

  FSWriteLineCR "<" & "%"

  FSWriteLineCR "  QueryStringSave"
  FSWriteLineCR ""
  FSWriteLineCR "  TableSetHeadRowBGColor ""#CCCCFF"""
  FSWriteLineCR "  TableSetRowBGColor     ""#ECECFF"""
  FSWriteLineCR ""
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Constantes do Sistema"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "  Public sparOption"
  FSWriteLineCR "  Public sparTarget"
  FSWriteLineCR "  Public sparType"
  FSWriteLineCR "  Public sparLess"
  FSWriteLineCR "  Public sparGreater"
  FSWriteLineCR ""
  FSWriteLineCR "  Const conScriptTimeout = 30"
  FSWriteLineCR ""
  FSWriteLineCR "  Const conPOption = ""O"""
  FSWriteLineCR "  Const conPTarget = ""T"""
  FSWriteLineCR "  Const conPType   = ""Y"""
  FSWriteLineCR "  Const conLess    = ""L"""
  FSWriteLineCR "  Const conGreater = ""G"""
  FSWriteLineCR ""
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Variaveis globais com parametros"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "  Public sobjCMD"
  FSWriteLineCR "  Public sobjConn"
  FSWriteLineCR "  Public sobjRS"
  FSWriteLineCR "  Public sobjRS2"
  FSWriteLineCR "  Public sobjErr"
  FSWriteLineCR ""
  FSWriteLineCR "  sparOption  = lCase(Request.QueryString(conPOption))"
  FSWriteLineCR "  sparTarget  = lCase(Request.QueryString(conPTarget))"
  FSWriteLineCR "  sparType    = lCase(Request.QueryString(conPType))"
  FSWriteLineCR "  sparLess    = lCase(Request.QueryString(conLess))"
  FSWriteLineCR "  sparGreater = lCase(Request.QueryString(conGreater))"
  FSWriteLineCR ""
  FSWriteLineCR "main"
  FSWriteLineCR ""
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Rotina Principal do Sistema"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "Private Sub Main"
  FSWriteLineCR "  Server.ScriptTimeOut = conScriptTimeout"
  FSWriteLineCR ""
  FSWriteLineCR "  Set sobjConn      = Server.CreateObject(""ADODB.Connection"")"
  FSWriteLineCR "  sobjConn.ConnectionTimeout = Session(""ConnectionTimeout"")"
  FSWriteLineCR "  sobjConn.CommandTimeout    = Session(""CommandTimeout"")"
  FSWriteLineCR ""
  FSWriteLineCR "  sobjConn.Open Session(""ConnectionString""), _"
  FSWriteLineCR "	Session(""RuntimeUserName""), _"
  FSWriteLineCR "	Session(""RuntimePassword"")"
  FSWriteLineCR ""
  FSWriteLineCR "  Set sobjRS = Server.CreateObject(""ADODB.RecordSet"")"
  FSWriteLineCR "  Set sobjRS2 = Server.CreateObject(""ADODB.RecordSet"")"
  FSWriteLineCR ""
  FSWriteLineCR "  Set sobjCMD = Server.CreateObject(""ADODB.Command"")"
  FSWriteLineCR "  Set sobjCMD.ActiveConnection = sobjConn"
  FSWriteLineCR "  sobjCMD.Prepared = True"
  FSWriteLineCR "  sobjCMD.CommandType = adCmdText"
  FSWriteLineCR ""
  FSWriteLineCR "  MainBody"
  FSWriteLineCR ""
  FSWriteLineCR "  Server.ScriptTimeOut = Session(""ScriptTimeOut"")"
  FSWriteLineCR ""
  FSWriteLineCR "  On Error Resume Next"
  FSWriteLineCR ""
  FSWriteLineCR "  sobjRS2.Close"
  FSWriteLineCR "  Set sobjRS2   = nothing"
  FSWriteLineCR ""
  FSWriteLineCR "  sobjRS.Close"
  FSWriteLineCR "  Set sobjRS   = nothing"
  FSWriteLineCR ""
  FSWriteLineCR "  sobjCMD.Close"
  FSWriteLineCR "  Set sobjCMD = nothing"
  FSWriteLineCR ""
  FSWriteLineCR "  sobjConn.Close"
  FSWriteLineCR "  Set sobjConn = nothing"
  FSWriteLineCR ""
  FSWriteLineCR "End Sub"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Final da Sub Main"
  FSWriteLineCR ""
  FSWriteLineCR ""
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Mostra a Primeira Pagina"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "Private Sub ShowFirstPage"
  FSWriteLineCR ""
  FSWriteLineCR "  ShowHTMLCR ""<H3 ALIGN=CENTER>Teste do Sistema de Edição</H3>"""
  FSWriteLineCR ""
  FSWriteLineCR "  ShowHTMLCR ""<HR>"""
  FSWriteLineCR ""
  FSWriteLineCR "  Center"
  FSWriteLineCR ""
  FSWriteLineCR "  TableSetSpacing 2"
  FSWriteLineCR "  TableSetPadding 2"
  FSWriteLineCR "  TableSetColumnNoWrap True"
  FSWriteLineCR ""
  FSWriteLineCR "  TableNormalBegin ""60%"""
  FSWriteLineCR ""
  FSWriteLineCR "  TableSetColumnColSpan """""
  FSWriteLineCR ""
  FSWriteLineCR "  TableBeginHeadRow 3"
  FSWriteLineCR ""
  FSWriteLineCR "  TableSetColumnAlign ""CENTER"""
  FSWriteLineCR ""
  FSWriteLineCR "  TableHeadColumn ""Tabelas do Sistema"""
  FSWriteLineCR ""
  FSWriteLineCR "  TableSetColumnAlign """""
  FSWriteLineCR ""
  FSWriteLineCR "  TableEndHeadRow"
  FSWriteLineCR ""
  FSWriteLineCR "  TableBeginRow 2"
  FSWriteLineCR "  TableSetColumnVAlign ""Top"""
  FSWriteLineCR ""
  FSWriteLineCR "  REM -----------------------------------------------------------------------"
  FSWriteLineCR "  REM Coluna com todas as tabelas"
  FSWriteLineCR "  REM -----------------------------------------------------------------------"
  FSWriteLineCR "  TableBeginColumn"
  FSWriteLineCR ""
  FSWriteLineCR "  EditCreateURLBegin sstrThisScriptName"
  FSWriteLineCR ""

  For x = 1 To intFileCounter
    FSWriteLineCR "  EditCreateURL conOption" & arrTable(x) & ", """ & arrTable(x) & """"
    FSWriteLineCR "  BR"
    
  Next

  FSWriteLineCR ""
  FSWriteLineCR "  EditCreateURLEnd"
  FSWriteLineCR ""
  FSWriteLineCR "  TableEndRow"
  FSWriteLineCR "  TableNormalEnd"
  FSWriteLineCR ""
  FSWriteLineCR "  TableSetColumnNoWrap False"
  FSWriteLineCR ""
  FSWriteLineCR "  CenterEnd"
  FSWriteLineCR ""
  FSWriteLineCR "End Sub"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Final da Sub FistPage"
  FSWriteLineCR ""

  REM -----------------------------------------------------------------------

  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Constantes da Aplicacao"
  FSWriteLineCR "REM -------------------------------------------------------------------------"

  For x = 1 To intFileCounter
    FSWriteLineCR "  Const conOption" & arrTable(x) & " = """ & x & """"

  Next

  FSWriteLineCR ""
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Cria todas as tabelas do sistema"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "Public Sub CriaTabelasAll"

  REM -----------------------------------------------------------------------
  For x = 1 To intFileCounter
    FSWriteLineCR ""
    FSWriteLineCR "  " & arrTable(x)
    FSWriteLineCR "  TableCreate """ & arrTable(x) & """"

  Next

  REM -----------------------------------------------------------------------
  FSWriteLineCR ""
  FSWriteLineCR "End Sub"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Final da Sub CriaTabelasAll"
  FSWriteLineCR ""

  ShowMessageError "Gerando do Create Table ALL"


  For x = 1 To intFileCounter
    strCurrentFile = arrTable(x)

    ShowMessageError sobjZTIConn.Connection

    sobjConn.Close

    sobjConn.Open sobjZTIConn.Connection, _
                  Session("RuntimeUserName"), _
                  Session("RuntimePassword")

REM    ShowMessageError strConnectionString

    Set sobjCMD.ActiveConnection = sobjConn

    sobjCMD.Prepared = True
    sobjCMD.CommandType = adCmdText

    sql = "SELECT * FROM " & strCurrentFile

    sobjCMD.CommandText = sql

    ShowMessageError sql

    sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

    ShowMessageError "Pega os tamanhos dos campos!"

    REM ---------------------------------------------------------------------

    intNameLen = 0
    intTypeLen = 0
    intSizeLen = 0

    strFormList = ""
    strFormUnit = ""
    strFormFind = ""

    For i = 0 To sobjRS.Fields.Count - 1
      intNameTemp = Len(Trim(sobjRS(i).Name))
      intTypeTemp = Len(Trim(GetFieldType(sobjRS(i).Type)))
      intSizeTemp = sobjRS(i).DefinedSize
      
      strAux = Trim(GetFieldType(sobjRS(i).Type))

      Select Case sobjRS(i).Type
        Case adDouble, adCurrency, adDate, adNumeric, adDBDate, adDBTime, adDBTimeStamp, adInteger, adUnsignedInt, adTinyInt, adUnsignedTinyInt, adSmallInt, adUnsignedSmallInt
          intSizeTemp = 1

        Case Else

      End Select

      intSizeTemp = Len(LTrim(CStr(intSizeTemp)))

      If intNameTemp > intNameLen Then
        intNameLen = intNameTemp

      End If

      If intTypeTemp > intTypeLen Then
        intTypeLen = intTypeTemp

      End If

      If intSizeTemp > intSizeLen Then
        intSizeLen = intSizeTemp

      End If

      strFormList = strFormList & "," & sobjRS(i).Name
      strFormUnit = strFormUnit & ";" & sobjRS(i).Name
      strFormFind = strFormFind & "," & sobjRS(i).Name

    Next

    FSWriteLineCR "REM -------------------------------------------------------------------------"
    FSWriteLineCR "REM  Tabela do " & strCurrentFile
    FSWriteLineCR "REM -------------------------------------------------------------------------"
    FSWriteLineCR "Public Sub " & strCurrentFile
    FSWriteLineCR ""

    strFormList = Mid(strFormList, 2)
    strFormUnit = Mid(strFormUnit, 2)
    strFormFind = Mid(strFormFind, 2)

    FSWriteLineCR "  DataBegin """ & strCurrentFile & """"
    FSWriteLineCR ""

    ShowMessageError intTypeLen & " - " & intTypeTemp
    
    REM ---------------------------------------------------------------------
    For i = 0 to sobjRS.Fields.Count - 1
      intFieldLen = sobjRS(i).DefinedSize
      strFieldType = GetFieldType(sobjRS(i).Type)

      Select Case sobjRS(i).Type
        Case adDouble, adCurrency, adDate, adNumeric, adDBDate, adDBTime, adDBTimeStamp
          intFieldLen = 8

        Case adInteger, adUnsignedInt
          intFieldLen = 4

        Case adTinyInt, adUnsignedTinyInt
          intFieldLen = 1

        Case adSmallInt, adUnsignedSmallInt
          intFieldLen = 2

        Case Else

      End Select

      FSWriteLineCR "  DataAddField """ & FixSpaces(sobjRS(i).Name & """,", intNameLen + 1, True) & _
                    " " & FixSpaces(strFieldType & ",", intTypeLen + 1, True) & _
                    FixSpaces(LTrim(CStr(intFieldLen)), intSizeLen - 1, False) & ", conDataNULL"

    Next

    FSWriteLineCR ""

    FSWriteLineCR "REM  DataAddPrimaryKey ""cliCodigo"""

    FSWriteLineCR "REM  DataIndexClustered ""Codigo"", ""cliCodigo"""
    FSWriteLineCR "REM  DataAddIndex ""Nome"", ""cliNome"""

    REM ---------------------------------------------------------------------
  	FSWriteLineCR "End Sub"
    FSWriteLineCR "REM -------------------------------------------------------------------------"
    FSWriteLineCR "REM Final da Sub " & strCurrentFile
    FSWriteLineCR ""
    FSWriteLineCR "REM -------------------------------------------------------------------------"
    FSWriteLineCR "REM  Guarda os dados do form " & strCurrentFile
    FSWriteLineCR "REM -------------------------------------------------------------------------"
    FSWriteLineCR "Public Sub Get" & strCurrentFile
    FSWriteLineCR ""
    FSWriteLineCR "  " & strCurrentFile
    FSWriteLineCR ""
    FSWriteLineCR "  EditFormBegin """ & strCurrentFile & _
                  """, """ & strCurrentFile & _
                  """, 1, conOption" & strCurrentFile & ", conClientValidation"
    FSWriteLineCR ""
    FSWriteLineCR "  EditFormList """ & strFormList & """"
    FSWriteLineCR ""
    FSWriteLineCR "  EditFormUnit """ & strFormUnit & """"
    FSWriteLineCR ""
    FSWriteLineCR "  EditFormFind """ & strFormFind & """"
    FSWriteLineCR ""

    For i = 0 to sobjRS.Fields.Count - 1
      FSWriteLineCR "  EditAddFormField """ & FixSpaces(sobjRS(i).Name & """, ", intNameLen + 2, True) & _
                    """" & FixSpaces(sobjRS(i).Name & """, ", intNameLen + 2, True) & _
                    "conTextField, 0, 0, """", """""

    Next

    FSWriteLineCR ""
    FSWriteLineCR "  EditFieldInternalLink """ & sobjRS(0).Name & """, conOption" & strCurrentFile & ", """ & sobjRS(0).Name & """"
    FSWriteLineCR ""

    For i = 0 to sobjRS.Fields.Count - 1
      FSWriteLineCR "  EditFormFieldHint """ & FixSpaces(sobjRS(i).Name & """, ", intNameLen + 2, True) & _
                    """" & sobjRS(i).Name & """"

    Next

    FSWriteLineCR ""
    FSWriteLineCR "  EditFormEnd"
    FSWriteLineCR ""
    FSWriteLineCR "End Sub"
    FSWriteLineCR "REM -------------------------------------------------------------------------"
    FSWriteLineCR "REM Final da Sub " & strCurrentFile
    FSWriteLineCR ""

    sobjRS.Close

  Next

  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Corpo Principal do sistema"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "Private Sub MainBody"
  FSWriteLineCR ""
  FSWriteLineCR "  EditQueryString"
  FSWriteLineCR ""
  FSWriteLineCR "  EditExeOptions sparEditOption"
  FSWriteLineCR ""
  FSWriteLineCR "  Select Case sparEditWhat"

  REM -----------------------------------------------------------------------
  For x = 1 To intFileCounter
    FSWriteLineCR ""
    FSWriteLineCR "    Case conOption" & arrTable(x)
    FSWriteLineCR "      Get" & arrTable(x)

  Next

  FSWriteLineCR ""
  FSWriteLineCR "  End Select"
  FSWriteLineCR ""
  FSWriteLineCR "  HTMLBegin"
  FSWriteLineCR "  HeadAll ""Sistema de Entrada de Dados"""
  FSWriteLineCR "  BodyBegin"
  FSWriteLineCR ""
  FSWriteLineCR "  EditShowOptions sparEditOption"
  FSWriteLineCR ""
  FSWriteLineCR "  PageFooterDefault"
  FSWriteLineCR ""
  FSWriteLineCR "  BodyEnd"
  FSWriteLineCR "  HTMLEnd"
  FSWriteLineCR ""
  FSWriteLineCR "End Sub"
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Final da Sub MainBody"
  FSWriteLineCR ""
  FSWriteLineCR "REM -------------------------------------------------------------------------"
  FSWriteLineCR "REM Fim do " & strResultFile

  FSWriteLineCR "%" & ">"

  sobjFile.Close
  Set sobjFile = Nothing

	Set sObjFS = Nothing

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub CreateTableProgramas

REM -------------------------------------------------------------------------
REM Grava uma linha no arquivo aberto
REM -------------------------------------------------------------------------
Private Sub FSWriteLineCR(ByVal strLine)
  sobjFile.WriteLine strLine

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FSWriteLineCR

REM -------------------------------------------------------------------------
REM Retorna o descritivo do tipo do campo
REM -------------------------------------------------------------------------
Private Function FixSpaces(ByVal strString, ByVal intMaxSpaces, ByVal blnAlign)
  Dim intLen, strSpaces     
  
  intMaxSpaces = intMaxSpaces + 1

  If IsNull(strString) Then
    strString = ""

  Else
    strString = Trim(LTrim(strString))

  End If

  If blnAlign Then
    FixSpaces = Left(strString & Space(intMaxSpaces), intMaxSpaces)

  Else
    FixSpaces = Right(Space(intMaxSpaces) & strString, intMaxSpaces)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub FSWriteLineCR

REM -------------------------------------------------------------------------
REM Retorna o descritivo do tipo do campo
REM -------------------------------------------------------------------------
Private Function GetFieldType(ByVal intValue)
  Dim strResult

REM  conDataInt
REM  conDataLong
REM  conDataMoney

  Select Case Fix(intValue)
    Case adVarChar
      strResult = "conDataVarChar"

    Case adDouble
      strResult = "conDataFloat"

    Case adDBDate, adDate, adDBTime, adDBTimeStamp
      strResult = "conDataDateTime"

    Case adUnsignedTinyInt, adTinyInt
      strResult = "conDataTinyInt"

    Case adSmallInt, adUnsignedSmallInt
      strResult = "conDataSmallInt"

    Case adBoolean
      strResult = "conDataBit"

    Case adInteger, adUnsignedInt
      strResult = "conDataInt"

    Case adCurrency
      strResult = "conDataMoney"

    Case adChar
      strResult = "conDataChar"

    Case Else
      strResult = CStr(intValue)

  End Select

  GetFieldType = strResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function GetFieldType

REM -------------------------------------------------------------------------
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------

Private Sub MainBody
  BodyLimit False

  HTMLBegin
  HeadAll "Cria Estrutura Inicial"
  BodyBegin

  PageHeaderDefault "<H2>Cria Estrutura Inicial</H2>"

  If sparOption > "" Then
    Select Case sparOption
      Case "900"
        CreateTablePrograms

      Case Else
        ShowMessageError "<center>Op&ccedil;&atilde;o (" & sparOption & ") Inexistente!!!</center>"

    End Select
  Else
    ShowMessageError "<center>Nenhuma Oção escolhida!!!</center>"

  End If

  PageFooterDefault

  BodyEnd
  HTMLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody

REM -------------------------------------------------------------------------
REM Fim do Inicia.asp
%>
  