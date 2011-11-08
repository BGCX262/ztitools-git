<%@ LANGUAGE="VBSCRIPT" %>
<%
REM -------------------------------------------------------------------------
REM  /sql/isql.asp
REM -------------------------------------------------------------------------
REM  Descricao   : ISQL Web
REM  Criacao     : 12:00h 20/01/97
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.1.a
REM  Copyright   : 1997 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  Site        : Zevallos(r) Tecnologia em Informacao
REM   URL        : http://www.zevallos.com.br
REM  Responsavel : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM ADODB Constants
REM -------------------------------------------------------------------------
'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset      = 1
Const adOpenDynamic     = 2
Const adOpenStatic      = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly        = 1
Const adLockPessimistic     = 2
Const adLockOptimistic      = 3
Const adLockBatchOptimistic = 4

REM -------------------------------------------------------------------------
REM Definicoes Globais do sistema
REM -------------------------------------------------------------------------
Dim objConn, objRS

Const constConnectionTimeout = 86400
Const constCommandTimeout    = 86400
Const constScriptTimeout     = 86400

numOldTimeOut        = Server.ScriptTimeOut
Server.ScriptTimeOut = constScriptTimeout

Set ObjConn = Server.CreateObject("ADODB.Connection")
Set ObjRS   = Server.CreateObject("ADODB.RecordSet")

objConn.ConnectionTimeout = constConnectionTimeout
objConn.CommandTimeout    = constCommandTimeout

Const ConstcmdExecuteQuery   = "Execute Query"
Const ConstcmdClearForm      = "Clear Form"
Const ConstcmdAbandonSession = "Abandon Session"

REM -------------------------------------------------------------------------
REM Iniciar o sistema
REM -------------------------------------------------------------------------
InicioSistema

REM -------------------------------------------------------------------------
REM Inicio do Form
REM -------------------------------------------------------------------------
Sub FormBegin(txtAction, txtMethod, txtName, txtTarget, txtOnSubmit)
  Dim txtHTML

  txtHTML = "<form action=" & txtAction

  If txtMethod > "" Then
    txtHTML = txtHTML & " method=" & txtMethod

  End If

  If txtName > "" Then
    txtHTML = txtHTML & " name=" & txtName

  End If

  If txtTarget > "" Then
    txtHTML = txtHTML & " target=" & Chr(34) & txtTarget & Chr(34)

  End If

  If txtOnSubmit > "" Then
    txtHTML = txtHTML & " onsubmit=" & Chr(34) & txtOnSubmit & Chr(34)

  End If

  Response.Write(txtHTML & ">" & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormBegin

REM -------------------------------------------------------------------------
REM Fim do Form
REM -------------------------------------------------------------------------
Sub FormEnd

  Response.Write("</form>" & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormEnd

REM -------------------------------------------------------------------------
REM Entra Check Box
REM -------------------------------------------------------------------------
Sub FormInputCheckBox(txtName, txtValue, txtAlign, numState)

  FormInput "checkbox", txtName, 0, txtValue, 0, txtAlign, numState

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputCheckBox

REM -------------------------------------------------------------------------
REM Entra Radio Buttom
REM -------------------------------------------------------------------------
Sub FormInputRadio(txtName, txtValue, txtAlign, numState)

  FormInput "radio", txtName, 0, txtValue, 0, txtAlign, numState

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputRadio

REM -------------------------------------------------------------------------
REM Entra Text Area
REM -------------------------------------------------------------------------
Sub FormInputTextArea(txtName, numSize, numMaxLength, numRows, numCols, txtValue)
  Dim txtHTML

  txtHTML = "<TEXTAREA type=text"

  If txtName > "" Then
    txtHTML = txtHTML & " name=" & txtName

  End If

  If numSize > 0 Then
    txtHTML = txtHTML & " size=" & numSize

  End If

  If numMaxLength > 0 Then
    txtHTML = txtHTML & " maxlength=" & numMaxLength

  End If

  If numRows > 0 Then
    txtHTML = txtHTML & " rows=" & numRows

  End If

  If numCols > 0 Then
    txtHTML = txtHTML & " cols=" & numCols

  End If

  txtHTML = txtHTML & ">"

  If txtValue > "" Then
     txtHTML = txtHTML & txtValue

  End If

  Response.Write(txtHTML & "</TEXTAREA>" & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputTextArea

REM -------------------------------------------------------------------------
REM Entra Botao Reset
REM -------------------------------------------------------------------------
Sub FormInputReset(txtValue)

  FormInput "reset", "", 0, txtValue, 0, "", -1

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputReset

REM -------------------------------------------------------------------------
REM Entra Botao Submit
REM -------------------------------------------------------------------------
Sub FormInputSubmit(txtName, txtValue)

  FormInput "submit", txtName, 0, txtValue, 0, "", -1

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputSubmit

REM -------------------------------------------------------------------------
REM Entra Campo Senha
REM -------------------------------------------------------------------------
Sub FormInputPassword(txtName, numSize)

  FormInput "password", txtName, numSize, "", 0, "", -1

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputPassword

REM -------------------------------------------------------------------------
REM Entra Campo Texto
REM -------------------------------------------------------------------------
Sub FormInputText(txtName, numSize, txtValue)

  FormInput "text", txtName, numSize, txtValue, 0, "", -1

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputText

REM -------------------------------------------------------------------------
REM Entra Campo Texto
REM -------------------------------------------------------------------------
Sub FormInputTextMaxLength(txtName, numSize, txtValue, numMaxLength)

  FormInput "text", txtName, numSize, txtValue, numMaxLength, "", -1

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputTextMaxLength

REM -------------------------------------------------------------------------
REM Entra Campo Invisivel
REM -------------------------------------------------------------------------
Sub FormInputHidden(txtName, txtValue)

  FormInput "hidden", txtName, 0, txtValue, 0, "", -1

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInputHidden

REM -------------------------------------------------------------------------
REM Entra InputDefault
REM -------------------------------------------------------------------------
Sub FormInput(txtType, txtName, numSize, txtValue, numMaxLength, txtAlign, numState)
  Dim txtHTML

  txtHTML = "<input type=" & txtType

  If txtName > "" Then
    txtHTML = txtHTML & " name=" & txtName

  End If

  If numSize > 0 Then
    txtHTML = txtHTML & " size=" & numSize

  End If

  If txtValue > "" Then
     txtHTML = txtHTML & " value=" & Chr(34) & txtValue & Chr(34)

  End If

  If numMaxLength > 0 Then
    txtHTML = txtHTML & " maxlength=" & numMaxLength

  End If

  If txtAlign > "" Then
    txtHTML = txtHTML & " align=" & txtAlign

  End If

  If numState = 1 Then
    txtHTML = txtHTML & " checked"

  End If

  Response.Write(txtHTML & ">" & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub FormInput

REM -------------------------------------------------------------------------
REM Mostra Fonte Arial e Tamanho
REM -------------------------------------------------------------------------
Sub ShowFontArialBegin(numSize)

  ShowFontBegin "Arial, Helvetica, Sans-Serif, Verdana", numSize

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowFontArialBegin

REM -------------------------------------------------------------------------
REM Mostra Fonte Escolhido
REM -------------------------------------------------------------------------
Sub ShowFontBegin(txtFace, numSize)
  Dim txtHTML

  If numSize > "" Or txtFace > "" Then
    txtHTML = "<font"

  End If

  If numSize <> 0 Then
    txtHTML = txtHTML & " font size=" & numSize

  End If

  If txtFace > "" Then
    txtHTML = txtHTML & " face=" & Chr(34) & txtFace  & Chr(34) & ">"

  Else
    txtHTML = txtHTML & ">"

  End If

  Response.Write(txtHTML & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowFontBegin

REM -------------------------------------------------------------------------
REM Mostra o texto Arial e Normal
REM -------------------------------------------------------------------------
Sub ShowFontEnd

  Response.Write("</font>" & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowFontEnd

REM -------------------------------------------------------------------------
REM Mostra o texto Arial e Normal
REM -------------------------------------------------------------------------
Sub ShowTextArial(txtText, numSize, txtFontShape)

  ShowText txtText, numSize, "Arial, Helvetica, Sans-Serif, Verdana", txtFontShape

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowTextArial

REM -------------------------------------------------------------------------
REM Mostra o texto Arial e Normal
REM -------------------------------------------------------------------------
Sub ShowTextArialNormal(txtText, numSize)

  ShowTextArial txtText, numSize, ""

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowTextArialNormal

REM -------------------------------------------------------------------------
REM Mostra o texto Arial e Strong
REM -------------------------------------------------------------------------
Sub ShowTextArialStrong(txtText, numSize)

  ShowTextArial txtText, numSize, "Strong"

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowTextArialStrong

REM -------------------------------------------------------------------------
REM Mostra o texto Arial e Strong
REM -------------------------------------------------------------------------
Sub ShowTextArialBold(txtText, numSize)

  ShowTextArial txtText, numSize, "B"

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowTextArialBold

REM -------------------------------------------------------------------------
REM Mostra o texto Generico
REM -------------------------------------------------------------------------
Sub ShowText(txtText, numSize, txtFace, txtFontShape)
  Dim txtHTML

  txtHTML = ""

  If numSize > "" Or txtFace > "" Then
    txtHTML = "<font"

  End If

  If numSize <> 0 Then
    txtHTML = txtHTML & " font size=" & numSize

  End If

  If txtFace > "" Then
    txtHTML = txtHTML & " face=" & Chr(34) & txtFace  & Chr(34) & ">"

  Else
    txtHTML = txtHTML & ">"

  End If

  If txtFontShape > "" Then
    txtHTML = txtHTML & "<" & txtFontShape & ">"

  End If

  txtHTML = txtHTML &  txtText

  If txtFontShape > "" Then
    txtHTML = txtHTML & "</" & txtFontShape & ">"

  End If

  If txtFace > "" Then
    txtHTML = txtHTML & "</font>"

  End If

  Response.Write(txtHTML & Chr(13))

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowText

REM -------------------------------------------------------------------------
REM Inicio do Script do Input
REM -------------------------------------------------------------------------
Sub ScriptInputBegin(txtForm, txtFunction)
  %>
  <SCRIPT LANGUAGE="JavaScript">
  <!--
  function <% =txtForm %><% =txtFunction %>() {
  <%
End Sub
REM -------------------------------------------------------------------------
REM End Sub ScriptInputBegin

REM -------------------------------------------------------------------------
REM Valida o Value.Length.Input Operator
REM -------------------------------------------------------------------------
Sub ScriptInputValueLength(txtOperator, txtHandle, txtVariable, numLength, txtMessage)
  %>
  if (document.<% =txtHandle %>.<% =txtVariable %>.value.length <% =txtOperator %> <% =numLength %>) {
    alert("<% =txtMessage %>");
    return false;
    }
  <%
End Sub
REM -------------------------------------------------------------------------
REM End Sub ScriptInputValueLength

REM -------------------------------------------------------------------------
REM Valida o Value.
REM -------------------------------------------------------------------------
Sub ScriptInputValueCompare(txtOperator, txtHandle, txtVariable1, txtVariable2, txtMessage)
  %>
  if (document.<% =txtHandle %>.<% =txtVariable1 %>.value <% =txtOperator %> document.<% =txtHandle %>.<% =txtVariable2 %>.value) {
    alert("<% =txtMessage %>");
    return false;
    }
  <%
End Sub
REM -------------------------------------------------------------------------
REM End Sub ScriptInputValueCompare

REM ---------------------------------------------------------------------
REM Final do Script do Input
REM ---------------------------------------------------------------------
Sub ScriptInputEnd
  %>
  return true; }
  //-->
  </SCRIPT>
  <%
End Sub
REM -------------------------------------------------------------------------
REM End Sub ScriptInputEnd

REM -------------------------------------------------------------------------
REM Repete String
REM -------------------------------------------------------------------------
Function RepeatString(cString, nTimes)
  Dim x, txtResult

  txtResult = ""

  For x = 0 To nTimes
    txtResult = txtResult & cString

  Next

  RepeatString = txtResult

End Function

REM -------------------------------------------------------------------------
REM End Function RepeatString

REM -------------------------------------------------------------------------
REM Ajuste do tamanho de strings
REM -------------------------------------------------------------------------
Function FitLen(cString, nLen)

  If isNull(cString) Then
    cString = ""

  End If

  FitLen = Left(cString & Space(nLen), nLen)

End Function

REM -------------------------------------------------------------------------
REM End Function FitLen

REM -------------------------------------------------------------------------
REM Ajuste do tamanho de strings
REM -------------------------------------------------------------------------
Function FitNumber(cNumber, nLen)

  FitNumber = Right(Space(nLen) & cNumber, nLen)

End Function

REM -------------------------------------------------------------------------
REM End Function FitLen

REM -------------------------------------------------------------------------
REM Entra com o formulario principal
REM -------------------------------------------------------------------------
Sub GetFormQuery(cDSN, cDatabase, cUID, nPageSize, cComm)
  Const constFormName = "frmISQL"

  ScriptInputBegin       constFormName, "ValidateInput"

  ScriptInputValueLength "<=", constFormName, "cDSN", 0, _
    "Enter the ODBC DSN!"

  ScriptInputValueLength "<=", constFormName, "cDatabase", 0, _
    "Enter the SQL Database Name!"

  ScriptInputValueLength "<=", constFormName, "cUID", 0, _
    "Enter the SQL User Name!"

  ScriptInputValueLength "<=", constFormName, "nPageSize", 0, _
    "Enter the lenght of the Page!"

  ScriptInputValueLength "<=", constFormName, "cComm", 0, _
    "Enter any valid the SQL Statement!"

  ScriptInputEnd

  FormBegin "isql.asp", "post", constFormName, "", ("return " & constFormName & "ValidateInput()")

  Response.Write("<table border=" & Session("TableBorder") & ">")
  Response.Write("<tr><td>")
  ShowTextArialBold "DSN:", 0
  Response.Write("<td>")

  FormInputText "cDSN", 15, cDSN

  ShowTextArialBold " Database:", 0
  FormInputText "cDatabase", 15, cDatabase

  Response.Write("</td></tr>")

  Response.Write("<tr><td>")
  ShowTextArialBold "UID:", 0
  Response.Write("<td>")

  FormInputText "cUID", 15, cUID

  ShowTextArialBold " PWD:", 0
  FormInputText "cPWD", 15, cPWD

  ShowTextArialBold " Page Size:", 0
  FormInputText "nPageSize", 8, nPageSize

  Response.Write("</td></tr>")

  Response.Write("<tr><td valign=top>")
  ShowTextArialBold "Command:", 0
  Response.Write("<td>")
  FormInputTextArea "cComm", 1024, 1024, 4, 70, cComm

  Response.Write("</td></tr></table>")

  Response.Write("<hr>")
  FormInputHidden "cFormSource", "Executa"
  FormInputSubmit "bMove", ConstcmdExecuteQuery
  FormInputReset ConstcmdClearForm

  Response.Write("<p>")

  FormInputSubmit "bMove", ConstcmdAbandonSession

  FormEnd

End Sub

REM -------------------------------------------------------------------------
REM End Sub GetForm

REM -------------------------------------------------------------------------
REM Rodape da Pagina
REM -------------------------------------------------------------------------
Sub PageFooter
  Set  objDirFooter = Server.CreateObject("ZTITools.Directory")

  %><HR>
  <table width=100% border=<% =Session("TableBorder") %>>
  <tr><%

  If Session("SourceCode") Then
    %><td width=1% align=left>
    <!--#include virtual="include/code/srcform.inc"-->
    </td>
  <% End If

  %><td align=center>
  <% ShowFontArialBegin -3 %>
  Um produto da <A HREF = "http://www.zevallos.com.br">ZTI - Zevallos&reg; Tecnologia em Informa&ccedil;&atilde;o.</A>.
  <br>
  Sugest&otilde;es e problemas encaminhar para o
  <A HREF="mailto:webmaster@zevallos.com.br">
  <img src="/img/icone/mailto.gif" alt="Mail To" border=0>
  <i>&lt;webmaster@zevallos.com.br&gt;</i></A>
  <br>
  &copy; 1997 <A HREF = "/copyright.asp">Zevallos&reg;</a> todos os direitos reservados.
  <%
  If Not Session("SourceCode") Then Response.Write("<!--")

  Response.Write("<br>&Uacute;ltima atualiza&ccedil;&atilde;o,")

  Response.Write( _
  Server.HTMLEncode(ObjDirFooter.LastDate(Request.ServerVariables("PATH_TRANSLATED"))))

  If Not Session("SourceCode") Then Response.Write("-->")

  ShowFontEnd

  %></td></tr>
  </table>
  <% ShowFontEnd
  %></BODY>
  </HTML><%
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub PageFooter

REM -------------------------------------------------------------------------
REM Inicio do sistema
REM -------------------------------------------------------------------------
Sub InicioSistema

  REM -----------------------------------------------------------------------
  REM Ativa o icone de ASP Source Code
  REM -----------------------------------------------------------------------
  If Request.QueryString("Source") = 1 Then
    Session("SourceCode") = True

  ElseIf Request.QueryString("Source") = 0 Then
    Session("SourceCode") = False

  End If

  If IsEmpty(Session("SourceCode")) Then Session("SourceCode") = False

  REM -----------------------------------------------------------------------
  REM Define a borda das tabelas
  REM -----------------------------------------------------------------------
  If Request.QueryString("Border") > "" Then
    Session("TableBorder") = Request.QueryString("Border")

  End If

  If IsEmpty(Session("TableBorder")) Then Session("TableBorder") = 0


End Sub

REM -------------------------------------------------------------------------
REM End Sub InicioSistema

REM -------------------------------------------------------------------------
REM Executa o resultado do SQL
REM -------------------------------------------------------------------------
Sub ResultadoSQL(objConn, objRS, nPageSize)

  lBOFEOF = (Not objRS.Bof And Not objRS.Eof)

  If lBOFEOF Then

    objRS.MoveFirst

    %>
    <P>
    <HR NOSHADE SIZE=1>
    <table border=1 width=1% cellpadding=1 cellspacing=0 rules=cols bordercolor=black>
    <tr bgcolor=black><td width=1%><font color=white>Reg<% =RepeatString("&nbsp;", 3) %></font></td><%

    For i = 0 to ObjRS.Fields.Count - 1
      cName  = objRS(i).Name
      cField = objRS(i).Value

      If isNull(cField) Then
        cField = ""

      End If

      nLenName = Len(cName)
      nLenField = objRS(i).DefinedSize

      If nLenName > nLenField Then
        nFieldLen = nLenName

      Else
        nFieldLen = nLenField

      End If

      %><td width="1%"><font color=white>&nbsp;<% =cName %><% =RepeatString("&nbsp;", nFieldLen - nLenName) %>&nbsp;</font></td><%

    Next

    Response.Write("</tr>")

    If Session("nPage") = "" Or cFormSource = "Executa" Then
      nPage = 1

    Else
      nPage = Session("nPage")

    End If

    If cBrowse = "PgDn" Then
      nPage = nPage + 1

      If nPage > objRS.PageCount Then
         nPage = objRS.PageCount

      End If
    ElseIf cBrowse = "PgUp" Then
      nPage = nPage - 1

      If nPage < 1 Then
         nPage = 1

      End If
    End If

    objRS.PageSize = nPageSize

    objRS.AbsolutePage = nPage

    Session("nPage") = nPage

    nRecordPosition = (nPage - 1) * objRS.PageSize

    nRecordCount = 1

    Do While nRecordCount <= objRS.PageSize And Not objRS.EOF
      ShowTRColor nRecordCount

      %><td align=right><% =nRecordCount %>&nbsp;</td><%

      For i = 0 to objRS.Fields.Count - 1
          cItem = RTrim(objRS(i).Value)

          If isNull(cItem) Then
            cItem = "Null"

          ElseIf cItem = "" Then
            cItem = "Empty"

          End If

          %><td>&nbsp;<% =cItem %>&nbsp;</td><%
      Next
      %>
      </tr>
      <%

      objRS.MoveNext

      nRecordCount = nRecordCount + 1

    Loop

    objRS.Close

    %>
    </table>
    <P><B>(<% =nRecordCount - 1 %> Linha(s) afetadas)</B>
    <form method=POST action="isql.asp">
    <input type=hidden name="cFormSource" Value="Browse">
    <input type=submit Name="bBrowse" value="PgUp">
    <input type=submit Name="bBrowse" value="PgDn">
    </form>
    <%

  Else
    Response.Write "<P><B>N&atilde;o h&aacute; retorno de mensagem!!!</B>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM End Sub ResultadoSQL

REM -------------------------------------------------------------------------
REM Criacao dos cookies
REM -------------------------------------------------------------------------
Sub CriaCookie(cDatabase, cDSN, cPWD, nPageSize, cOPEN, cLOCK, cComm)

  Response.Cookies("iSQL")("cDataBase") = cDataBase
  Response.Cookies("iSQL")("cDSN")      = cDSN
  Response.Cookies("iSQL")("cUID")      = cUID
  Response.Cookies("iSQL")("cPWD")      = cPWD
  Response.Cookies("iSQL")("nPageSize") = nPageSize
  Response.Cookies("iSQL")("cOPEN")     = cOPEN
  Response.Cookies("iSQL")("cLOCK")     = cLOCK
  Response.Cookies("iSQL")("cComm")     = cComm

  Response.Cookies("iSQL").Domain = Request.ServerVariables("HTTP_HOST")
  Response.Cookies("iSQL").Expires = Now() + 90
  Response.Cookies("iSQL").Secure = FALSE

End Sub
REM -------------------------------------------------------------------------
REM End Sub CriaCookie

REM -------------------------------------------------------------------------
REM Mensagem de erro de digitacao
REM -------------------------------------------------------------------------
Sub  ShowFormError(nError, cMessage, cLast)
  %>
  <h1 align=center>Erro</h1>
  <hr>
  <%
  If nError > 1 Then
    Response.Write Left(cMessage, Len(cMessage) - 2) & _
                   " e o " & Left(cLast, Len(cLast) - 2) & _
                   " s&atilde;o obrigat&oacute;rios"

  Else
    Response.Write cMessage & _
                   Left(cLast, Len(cLast) - 2) & _
                   " &eacute; obrigat&oacute;rio!"

  End If

  %>
  <p><a href="isql.asp">Volta para o ISQL</a>
  <%

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowFormError


REM -------------------------------------------------------------------------
REM Muda a cor da linha de acordo com o contador
REM -------------------------------------------------------------------------
Sub ShowTRColor(numCounter)

  If Int(numCounter Mod 2) = 0 Then
    %><TR BGCOLOR=lightyellow><%

  Else
    %><tr bgcolor=white><%

  End If

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowTRColor

REM -------------------------------------------------------------------------
REM Inicio do ISQL
REM -------------------------------------------------------------------------
lOk      = True
cLast     = ""
cMessage = "O "
nError   = 0

cFormSource = Request.Form("cFormSource")
cRequest    = Request("REQUEST_METHOD") & cFormSource
cButton     = Request.Form("bMove")
cBrowse     = Request.Form("bBrowse")

If cRequest = "POSTExecuta" Then
  cDataBase = Request.Form("cDataBase")
  cDSN      = Request.Form("cDSN")
  cUID      = Request.Form("cUID")
  cPWD      = Request.Form("cPWD")
  nPageSize = Request.Form("nPageSize")
  cComm     = Request.Form("cComm")

  If cDataBase = "" Then
    lOK = False
    nError = nError + 1

    cMessage = cMessage & cLast
    cLast    = "DataBase, "

  End If

  If cDSN = "" Then
    lOK = False
    nError = nError + 1

    cMessage = cMessage & cLast
    cLast    = "DSN, "

  End If

  If cUID = "" Then
    lOK = False
    nError = nError + 1

    cMessage = cMessage & cLast
    cLast    = "UID, "

  End If

  If nPageSize = "" Then
    lOK = False
    nError = nError + 1

    cMessage = cMessage & cLast
    cLast    = "Page Size, "

  End If

  If cComm = "" Then
    lOK = False
    nError = nError + 1

    cMessage = cMessage & cLast
    cLast    = "Comm, "

  End If

  If lOK Then CriaCookie cDatabase, cDSN, cPWD, nPageSize, cOPEN, cLOCK, cComm

End If
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<!--
ZTI - Zevallos(r) Tecnologia em Informacao
Brasilia - DF - Brasil
webmaster@zevallos.com.br
http://www.zevallos.com.br

Ruben Zevallos Jr. - zevallos@zevallos.com.br
-->

<META NAME="ROBOTS" CONTENT="NOINDEX">
<META HTTP-EQUIV="PRAGMA" CONTENT="NO-CACHE">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">
<META HTTP-EQUIV="REPLY-TO" CONTENT="webmaster@zevallos.com.br">
<META HTTP-EQUIV="EXPIRES" CONTENT="23 Aug 1999 00:01 GMT">

<META http-equiv="PICS-Label" content='(PICS-1.1 "http://www.rsac.org/ratingsv01.html" l gen true comment "RSACi North America Server" by "webmaster@zevallos.com.br" for "http://www.zevallos.com.br" on "1997.06.26T21:24-0500" r (n 0 s 0 v 0 l 0))'>

<META NAME="KEYWORDS" CONTENT="SQL; ISQL; Microsoft ISQL; Brasil; Brazil; ZTI; Zevallos Tecnologia em Informacao;ASP; Active Server Pages">
<META NAME="DESCRIPTION" CONTENT="."
<META NAME="PRODUCT" CONTENT="ISQL Web 1.1c">
<META NAME="LOCALE" CONTENT="PO-BR">
<META NAME="CHARSET" CONTENT="US-ASCII">
<META NAME="CATEGORY" CONTENT="HOME PAGE">
<META NAME="GENERATOR" CONTENT="Tecnologias da ZTI em ASP">
<META NAME="AUTHOR" CONTENT="ZTI - Zevallos(r) Tecnologia em Informacao - Brasilia - DF - Brasil - webmaster@zevallos.com.br - http://www.zevallos.com.br">

<LINK REL="Home" HREF="/Default.asp" TITLE="Zevallos Tecnologia em Informacao">
<LINK REL="Copyright" HREF="/copyright.htm" TITLE="ZTI - Zevallos(r) Tecnologia em Informacao">
<LINK REV="Made" HREF="mailto:webmaster@zevallos.com.br" TITLE="WebMaster da Zevallos(r)">

<title>ISQL Web</title>
</head>
<body bgcolor=white>
<%

ShowFontArialBegin 0 

%><h2 align=center>ISQL Web 1.1c</h2>
<hr>
<p><%

If lOk Then
REM ------------------------------------------------------------------------
REM Variaveis locais
REM ------------------------------------------------------------------------
  cDataBase = Request.Cookies("iSQL")("cDataBase")
  cDSN      = Request.Cookies("iSQL")("cDSN")
  cUID      = Request.Cookies("iSQL")("cUID")
  cPWD      = Request.Cookies("iSQL")("cPWD")
  nPageSize = Request.Cookies("iSQL")("nPageSize")
  cOPEN     = Request.Cookies("iSQL")("cOPEN")
  cLOCK     = Request.Cookies("iSQL")("cLOCK")
  cComm     = Request.Cookies("iSQL")("cComm")

  If cDataBase = "" Then cDataBase = Session("SQLDatabase")
  If cDSN = ""       Then cDSN      = Session("SQLDSN")
  If cUID = ""       Then cUID      = Session("SQLUID")
  If nPageSize = "" Then nPageSize = 20

  GetFormQuery cDSN, cDatabase, cUID, nPageSize, cComm

REM ------------------------------------------------------------------------
REM Entrada dos dados
REM ------------------------------------------------------------------------

  cButtReq = cButton & cRequest

  If cButtReq = (ConstcmdAbandonSession & "POSTExecuta") Then
     Session.Abandon
     
    %><h1 align=center>Abandono de Sess&atilde;o</h1>
    <hr>
    <p>
    <p align=center><a href="default.asp">Volta para o ISQL</a>
    <p>
    <%

  ElseIf cButtReq = (ConstcmdExecuteQuery & "POSTExecuta") Or cFormSource = "Browse" Then

       On Error Resume Next

      txtDatabase = "database=" & cDatabase & ";" & _
                    "dsn=" & cDSN & ";" & _
                    "uid=" & cUID & ";" & _
                    "pwd=" & cPWD

       objConn.Open txtDatabase

      objConn.BeginTrans

      objRS.Open cComm, objConn, adOpenKeySet, adLockOptimistic

      If Not objConn.Errors.Count = 0  Then
        ErrorHandler objConn, "CREATE TABLE - Do", sql

      Else

        ResultadoSQL objConn, objRS, nPageSize

      End If

      objConn.CommitTrans

      objConn.Close

  End if
Else

  ShowFormError nError, cMessage, cLast

End If

Server.ScriptTimeOut = numOldTimeOut

PageFooter

REM ---------------------------------------------------------------------
REM Fim do ISQL.asp
%>
