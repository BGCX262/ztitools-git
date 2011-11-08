<%@ LANGUAGE="VBSCRIPT" %>
<%
REM -------------------------------------------------------------------------
REM  /sql/isql.asp
REM -------------------------------------------------------------------------
REM  Descricao   : Teste com o ISQL
REM  Cria‡ao     : 12:00h 20/01/97
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao			 : 1.0.0
REM  Copyright   : 1997 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  Site        : Zevallos(r) Tecnologia em Informacao
REM	 URL         : http://www.zevallos.com.br
REM  Responsavel : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------
%>

<!--#include virtual="/include/ADOvbs.inc"-->

<!--#include virtual="/include/ZTIDBLib.inc"-->

<!--#include virtual="/include/PageDefault.inc"-->

<%
const constConnectionTimeout = 86400
const constCommandTimeout		 = 86400
const constScriptTimeout		 = 86400

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
REM End Funcion FitLen

REM -------------------------------------------------------------------------
REM Ajuste do tamanho de strings
REM -------------------------------------------------------------------------
Function FitNumber(cNumber, nLen)

	FitNumber = Right(Space(nLen) & cNumber, nLen)
	
End Function

REM -------------------------------------------------------------------------
REM End Funcion FitLen

lOk 		 = True
cLast		 = ""
cMessage = "O "
nError	 = 0

numOldTimeOut = Server.ScriptTimeOut
Server.ScriptTimeOut = constScriptTimeout

cFormSource = Request.Form("cFormSource")
cRequest = Request("REQUEST_METHOD") & cFormSource
cButton = Request.Form("bMove")
cBrowse = Request.Form("bBrowse")

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
		cLast 	 = "DataBase, "

	End If

	If cDSN = "" Then
		lOK = False
		nError = nError + 1

		cMessage = cMessage & cLast
		cLast 	 = "DSN, "

	End If

	If cUID = "" Then
		lOK = False
		nError = nError + 1

		cMessage = cMessage & cLast
		cLast 	 = "UID, "

	End If

	If nPageSize = "" Then
		lOK = False
		nError = nError + 1

		cMessage = cMessage & cLast
		cLast 	 = "Page Size, "

	End If

	If cComm = "" Then
		lOK = False
		nError = nError + 1

		cMessage = cMessage & cLast
		cLast 	 = "Comm, "

	End If

	If lOK Then
REM -------------------------------------------------------------------------
REM Criacao dos cookies
REM -------------------------------------------------------------------------
		Response.Buffer = True
	
		Response.Cookies("iSQL").Domain = Request.ServerVariables("HTTP_HOST")
		Response.Cookies("iSQL").Secure = FALSE
		Response.Cookies("iSQL").Expires = #August 1, 1999 1:00:00 AM#
	
		Response.Cookies("iSQL")("cDataBase") = cDataBase
		Response.Cookies("iSQL")("cDSN")      = cDSN     
		Response.Cookies("iSQL")("cUID")      = cUID     
		Response.Cookies("iSQL")("cPWD")      = cPWD     
		Response.Cookies("iSQL")("nPageSize") = nPageSize
		Response.Cookies("iSQL")("cOPEN")     = cOPEN    
		Response.Cookies("iSQL")("cLOCK")     = cLOCK    
		Response.Cookies("iSQL")("cComm")     = cComm    
	
		Response.Expires = 0
		Response.ExpiresAbsolute = #August 1, 1999 1:00:00 AM#
	
		Response.Flush

	End If
End If
%>

<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<title>INPI - ISQL</title>

<!--#include virtual="/include/MetaDefault.inc"-->

</head>
<body bgcolor=<% =Session("PageBGColor") %> TOPMARGIN=0 LEFTMARGIN=0 background=<% =Session("PageBackground") %>>

<!-- Inicio Tabela Geral da tela -->
<% BeginPageTable %>

	<!-- Cabecalho da pagina -->
	<% PageHeader %>

	<!-- Inicio do corpo da pagina -->
	<% BeginBodyArea %>

		<!-- Menu da pagina -->
		<% PageMainMenu %>

		<!-- Inicio da Area Livre -->
		<% BeginPageFreeArea %>
					
			<!-- Texto para novidades -->
			<tr><td>
<%
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

	If cDataBase = "" Then
		cDataBase = Session("SQLDatabase")
		
	End If
	
	If cDSN = "" Then
		cDSN = Session("SQLDSN")
		
	End If
	
	If cUID = "" Then
		cUID = Session("SQLUID")
		
	End If

	If nPageSize = "" Then
		nPageSize = 20
		
	End If

%>	
  <form
    method=post
    action="isql.asp"
    >
  <table border = 0>
  <tr><td>DSN:
      <td><input
            type=text
            size=10
            maxlength=30
            name="cDSN"
            Value="<% = cDSN %>"
            >
  Database:
  <input
            type=text
            size=10
            maxlength=30
            name="cDatabase"
            Value="<% = cDatabase %>"
            >
  <tr><td>UID:
      <td><input
           type=text
           size=10
           maxlength=30
           name="cUID"
           Value="<% = cUID %>"
         >
  PWD:
      <input
           type=text
           size=10
           maxlength=30
           name="cPWD"
           Value="<% = cPWD %>"
         >
Page Size:
      <input
           type=text
           size=10
           maxlength=30
           name="nPageSize"
           Value="<% = nPageSize %>"
         >
           
  <tr><td valign=top>Command:
      <td><TEXTAREA
           type=text
           size=1024
           maxlength=1024
           rows=4
           cols=70
           name="cComm"
           ><% = cComm %></TEXTAREA>
  </table>
  <hr>
	<input
		type=hidden
		name="cFormSource"
		Value="Executa"
		>
  <input
    type=submit
    Name="bMove"
    value="Executa"
    >
  <input
    type=reset
     value="Apaga Formul&aacute;rio"
     >
  <p>
  <input
    type=submit
    Name="bMove"
    value="Abandon Session"
    >
  </form>

<%
REM ------------------------------------------------------------------------
REM Entrada dos dados
REM ------------------------------------------------------------------------

	cButtReq = cButton & cRequest

	If cButtReq = "Abandon SessionPOSTExecuta" Then
	 	Session.Abandon
%>
		<h1 align=center>Abandono de Sess&atilde;o</h1>
		<hr>
	  <p>
		<p align=center><a href="default.asp">Volta para o ISQL</a>
	  <p>
<%

	ElseIf cButtReq = "ExecutaPOSTExecuta" Or cFormSource = "Browse" Then
		 	On Error Resume Next
		
		  Set ObjConn = Server.CreateObject("ADODB.Connection")
		  Set ObjRS = Server.CreateObject("ADODB.RecordSet")		  
		
			objConn.ConnectionTimeout = constConnectionTimeout
			objConn.CommandTimeout		 = constCommandTimeout		
			
		 	objConn.Open "database=" & cDatabase & ";" & _
				"dsn=" & cDSN & ";" & _
				"uid=" & cUID & ";" & _
				"pwd=" & cPWD
		
			objConn.BeginTrans
			
			objRS.Open cComm, objConn, _
          adOpenKeySet, _
          adLockOptimistic

			If Not objConn.Errors.Count = 0  Then
				ErrorHandler objConn, "CREATE TABLE - Do", sql
		
			Else
				lBOFEOF = (Not objRS.Bof And Not objRS.Eof)
				
				If lBOFEOF Then
					Dim nFieldsLen(1000)
	
					Response.Write("<P><HR NOSHADE SIZE=1>")
					Response.Write "<DIV CLASS=result><PRE>"
	
					objRS.MoveFirst
					
					Response.Write "Reg   "
	
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
						
						nFieldsLen(i) = nFieldLen
						
						Response.Write FitLen(cName, nFieldLen) & " "
						
					Next
	
					Response.Write "<BR>----- "
	
					For i = 0 to ObjRS.Fields.Count - 1
						Response.Write String( nFieldsLen(i), "-") & " "
						
					Next
					
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
						Response.Write "<BR>" & FitNumber(nRecordCount + nRecordPosition, 5) & " "
						
						For i = 0 to objRS.Fields.Count - 1
								cItem = RTrim(objRS(i).Value)
								
								If isNull(cItem) Then 
									cItem = "Null"
									
								ElseIf cItem = "" Then
									cItem = "Empty"
							
								End If	
								
								Response.Write FitLen(cItem, nFieldsLen(i)) & " "
						Next
	
						objRS.MoveNext
						
						nRecordCount = nRecordCount + 1
			
					Loop

					Response.Write "<BR>----- "
	
					For i = 0 to ObjRS.Fields.Count - 1
						Response.Write String( nFieldsLen(i), "-") & " "
						
					Next

					Response.Write "<P><B>(" & nRecordCount - 1 & " Linha(s) afetadas)</B>"
					Response.Write "</PRE></DIV>"

					objRS.Close
%>
					<form
						method=POST
						action="isql.asp"
						>
					<input
						type=hidden
						name="cFormSource"
						Value="Browse"
						>
					<input
						type=submit
						Name="bBrowse"
						value="PgUp"
						>
					<input
						type=submit
						Name="bBrowse"
						value="PgDn"
						>
					</form>
<%

				Else
					Response.Write "<P><B> N&atilde;o h&aacute; retorno de mensagem!!</B>"

				End If
		  End If

			objConn.CommitTrans
		
		  objConn.Close

	End if
Else
REM -------------------------------------------------------------------------
REM Mensagem de erro de digitacao
REM -------------------------------------------------------------------------
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

  Response.Write "<p><a href=" & _
 								 Chr(34) & "isql.asp" & _
 								 Chr(34) & ">Volta para o ISQL</a>"
End If

Server.ScriptTimeOut = numOldTimeOut
	
%>
<p><a href="/default.asp">Volta para o menu</a>

		</td></tr>
		<!-- Fim da Area Livre -->
		<% EndPageFreeArea %>

	<!-- Fim do corpo da pagina -->
	<% EndBodyArea %>

	<!-- Rodape da pagina -->
	<% PageFooter %>

<!-- Fim Tabela Geral da tela -->
<% EndPageTable %>

</body>
</html>

<%
REM ---------------------------------------------------------------------
REM Fim do ISQL.asp
%>
