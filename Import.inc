<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Import.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de importacao
REM  Cria‡ao     : 1/30/99 12:31AM
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

Dim sstrConnectionString
Dim sstrImportSource, sstrsstrImportSourceTarget

Dim sImportSetSource
Dim sImportSetSourceDriverID
Dim sImportSetSourceDBQ
Dim sImportSetSourceServer
Dim sImportSetSourceDataBase
Dim sImportSetSourceUserId
Dim sImportSetSourcePassword
Dim sImportSetSourceDefaultDir

sImportSetSource = 0
sImportSetSourceDriverID = 0
sImportSetSourceDBQ = ""
sImportSetSourceServer = ""
sImportSetSourceDataBase = ""
sImportSetSourceUserId = ""
sImportSetSourcePassword = ""
sImportSetSourceDefaultDir = ""

Dim sImportSetTarget 
Dim sImportSetTargetDriverID
Dim sImportSetTargetDBQ
Dim sImportSetTargetServer 
Dim sImportSetTargetDataBase
Dim sImportSetTargetUserId
Dim sImportSetTargetPassword
Dim sImportSetTargetDefaultDir

sImportSetTarget = 0
sImportSetTargetDriverID = 0
sImportSetTargetDBQ = ""
sImportSetTargetServer  = ""
sImportSetTargetDataBase = ""
sImportSetTargetUserId = ""
sImportSetTargetPassword = ""
sImportSetTargetDefaultDir = ""

REM -------------------------------------------------------------------------
REM Inicio da Importacao
REM -------------------------------------------------------------------------
Public Sub ImportBegin(ByVal strSource, ByVal strSource)

  sstrImportSource = strSource
  sstrImportTarget = strTarget

End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportBegin

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetSource(ByVal strValue)

  sImportSetSource = strValue
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetSource

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetSourceDriver(ByVal strValue)

  sImportSetSourceDriver = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetSourceDriver

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetSourceServer(ByVal strValue)

  sImportSetSourceServer = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetSourceServer

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetSourceDataBase(ByVal strValue)

  sImportSetSourceDataBase = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetSourceDataBase

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetSourceUserId(ByVal strValue)

  sImportSetSourceUserId = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetSourceUserId

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetSourcePassword(ByVal strValue)

  sImportSetSourcePassword = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetSourcePassword

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetTarget(ByVal strValue)

  sImportSetTarget = strValue
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetTarget

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetTargetDriver(ByVal strValue)

  sImportSetTargetDriver = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetTargetDriver

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetTargetServer(ByVal strValue)

  sImportSetTargetServer = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetTargetServer

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetTargetDataBase(ByVal strValue)

  sImportSetTargetDataBase = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetTargetDataBase

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetTargetUserId(ByVal strValue)

  sImportSetTargetUserId = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetTargetUserId

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Sub ImportSetTargetPassword(ByVal strValue)

  sImportSetTargetPassword = strValue
    
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportSetTargetPassword

REM -------------------------------------------------------------------------
REM Inicio da Importacao
REM -------------------------------------------------------------------------
Public Sub ImportEnd
  Dim strSC

  sobjConn.Close

  DataSetConnectionType = sImportSetSource
  DataSetDriverID       = sImportSetSourceDriverID
  DataSetDBQ            = sImportSetSourceDBQ
  DataSetServer         = sImportSetSourceServer
  DataSetDataBase       = sImportSetSourceDataBase
  DataSetUserId         = sImportSetSourceUserId
  DataSetPassword       = sImportSetSourcePassword
  DataSetDefaultDir     = sImportSetSourceDefaultDir

  strCS = DataConnection

  sobjConn.Open strConnectionString, _
                sImportSetSourceUserId, _
                sImportSetSourcePassword

  ShowMessageError strConnectionString

  Set sobjCMD.ActiveConnection = sobjConn
  
  sobjCMD.Prepared = True
  sobjCMD.CommandType = adCmdText
  
  sql = "SELECT * FROM " & sstrImportSource

  sobjCMD.CommandText = sql

  ShowMessageError sql
      
  sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

  ShowMessageError "Tabela de origem " & sstrImportSource & " aberto!"

  DataSetConnectionType = sImportSetTarget
  DataSetDriverID       = sImportSetTargetDriverID
  DataSetDBQ            = sImportSetTargetDBQ
  DataSetServer         = sImportSetTargetServer
  DataSetDataBase       = sImportSetTargetDataBase
  DataSetUserId         = sImportSetTargetUserId
  DataSetPassword       = sImportSetTargetPassword
  DataSetDefaultDir     = sImportSetTargetDefaultDir

  strCS = DataConnection
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Function ImportBegin

REM -------------------------------------------------------------------------
REM Fim do Import.inc
%>
