<%
REM -------------------------------------------------------------------------
REM  /ZTITools/All.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Defaults
REM  Cria‡ao     : 11:23h 7/2/1998
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998, 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

Option Explicit 
Response.Expires = 0

%>
<!--#INCLUDE VIRTUAL="/HiperTools/AdoVbs.inc"--> 
<!--#INCLUDE VIRTUAL="/HiperTools/EditForm3.inc"--> 
<%

Const conDateTypeAAMMDD   = 0
Const conDateTypeAAAAMMDD = 1
Const conDateTypeMMDDAA   = 2
Const conDateTypeMMDDAAAA = 3
Const conDateTypeDDMMAA   = 4
Const conDateTypeDDMMAAAA = 5

Function CrLF
 CrLF = Chr(13) & Chr(10)

End Function

REM -------------------------------------------------------------------------
REM Fim do All.inc
%>