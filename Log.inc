<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Log.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Logs
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
REM Retorna o contador do form
REM -------------------------------------------------------------------------
Function AddLog(ByVal intCode, ByVal strMessage)
  Dim objConn, objRS, objCMD, objBC
  Dim strBrowser, strVersion

	Set objBC = Server.CreateObject("MSWC.BrowserType")

	strBrowser = objBC.browser
	strVersion = objBC.version

	If  strBrowser = "Unknown" Then
		strBrowser = Null
		
	End If	

	If  strVersion = "Unknown" Then
		strVersion = Null
		
	End If	
  
  Server.ScriptTimeOut = conScriptTimeout

  Set objConn      = Server.CreateObject("ADODB.Connection")
  objConn.ConnectionTimeout = Session("ConnectionTimeout")
  objConn.CommandTimeout    = Session("CommandTimeout")

  objConn.Open Session("ConnectionString"), _
               Session("RuntimeUserName"), _
               Session("RuntimePassword")

  Set objRS = Server.CreateObject("ADODB.RecordSet")

  Set objCMD = Server.CreateObject("ADODB.Command")
  Set objCMD.ActiveConnection = objConn
  objCMD.Prepared = True
  objCMD.CommandType = adCmdText

  objCMD.CommandText = "ZTILog"
  objCMD.CommandType = adCmdTable

  objRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

  objRS.AddNew
  
  objRS("logData")           = Now
  objRS("logSession")        = Session.SessionID
  objRS("logAplicacao")      = Session("SessionAplication")
  objRS("logSiteCounter")    = Session("SiteCounter")
  objRS("logUser")           = Session("UserID")
  objRS("logCode")           = intCode
  objRS("logMessage")        = strMessage
  objRS("logIP")             = Request.ServerVariables("remote_addr")

  objRS("logBrowserName")    = strBrowser
  objRS("logBrowserVersion") = strVersion

  objRS.Update

  Server.ScriptTimeOut = Session("ScriptTimeOut")

  On Error Resume Next

  objRS.Close
  Set objRS   = nothing

  objCMD.Close
  Set objCMD = nothing

  objConn.Close
  Set objConn = nothing

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCounter

REM -------------------------------------------------------------------------
REM Fim do Edit.inc
%>
