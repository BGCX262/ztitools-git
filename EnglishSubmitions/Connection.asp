<%@ LANGUAGE="VBSCRIPT" %>
<%
REM -------------------------------------------------------------------------
REM Connection.ASP
REM -------------------------------------------------------------------------
REM Descricao   : Sample of ZTITools.Connecion
REM Criacao     : 1/30/99 2:23AM
REM Local       : Brasilia/DF/Brazil
REM Elaborado   : Ruben Zevallos Jr. - zevallos@zevallos.com.br
REM Versao      : 1.0.0
REM Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM             : http://www.zevallos.com.br
REM -------------------------------------------------------------------------
REM The ZTITools.Connection ScriptLet is part of ZTITools 2.5 a complet tool
REM that search to abstract the HTML to ASP programers.
REM You can use the following code as you want, but we will glad if you 
REM notice us and/or save some information about us.
REM The author is not responsible for any damage or problem related with 
REM part or the entire source of ZTItools.Connection.
REM -------------------------------------------------------------------------

  Set sobjConn      = Server.CreateObject("ADODB.Connection")
  
  Set sobjZTIConn = Server.CreateObject("ZTITools.Connection")

  sobjZTIConn.ConnectionType = conConnSQL
  sobjZTIConn.ServerAddress = "SQLServer"
  sobjZTIConn.DataBase = "ZTI"
  sobjZTIConn.UserId = "sa"
  sobjZTIConn.Password = ""
  
  Response.Write "<h3 align=center>ZTITools.Connection Properties</h3><HR>"

  Response.Write "<BR>ConnectionType=" & sobjZTIConn.ConnectionType
  Response.Write "<BR>DriverID=" & sobjZTIConn.DriverID      
  Response.Write "<BR>DBQ=" & sobjZTIConn.DBQ           
  Response.Write "<BR>ServerAddress=" & sobjZTIConn.ServerAddress 
  Response.Write "<BR>DataBase=" & sobjZTIConn.DataBase      
  Response.Write "<BR>UserID=" & sobjZTIConn.UserId        
  Response.Write "<BR>Password=" & sobjZTIConn.Password      
  Response.Write "<BR>DefaultDir" & sobjZTIConn.DefaultDir    
  Response.Write "<BR>ConnectionTimeout" & sobjZTIConn.ConnectionTimeout    
  Response.Write "<BR>CommandTimeout" & sobjZTIConn.CommandTimeout    

  Response.Write "<P>"

  Response.Write "<h3 align=center>ZTITools.Connection Methods</h3><HR>"

  Response.Write "<BR>Conn" & sobjZTIConn.Connection
  Response.Write "<BR>Conn" & sobjZTIConn.Clear
  
  Response.Write "<P>"

  Response.Write "<h3 align=center>ZTITools.Connection Properties - After Clear Method</h3><HR>"

  Response.Write "<BR>ConnectionType=" & sobjZTIConn.ConnectionType
  Response.Write "<BR>DriverID=" & sobjZTIConn.DriverID      
  Response.Write "<BR>DBQ=" & sobjZTIConn.DBQ           
  Response.Write "<BR>ServerAddress=" & sobjZTIConn.ServerAddress 
  Response.Write "<BR>DataBase=" & sobjZTIConn.DataBase      
  Response.Write "<BR>UserID=" & sobjZTIConn.UserId        
  Response.Write "<BR>Password=" & sobjZTIConn.Password      
  Response.Write "<BR>DefaultDir" & sobjZTIConn.DefaultDir    
  Response.Write "<BR>ConnectionTimeout" & sobjZTIConn.ConnectionTimeout    
  Response.Write "<BR>CommandTimeout" & sobjZTIConn.CommandTimeout    

  Response.Write "<P>"

  Response.Write "<h3 align=center>ZTITools.Connection Method Connection - After Clear Method</h3><HR>"

  Response.Write "<BR>Conn" & sobjZTIConn.Connection
  
  sobjConn.Open sobjZTIConn.Connection

  Set sobjZTIConn = Nothing
  
  On Error Resume Next

  sobjConn.Close
  Set sobjConn = nothing

%>