<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Test.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 10:50h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Testa se a opcao atual é do edit form
REM -------------------------------------------------------------------------
Private Function IsEditFormOption(ByVal strOption)
  IsEditFormOption = (Left(strOption, 1) = "h")

End Function
REM -------------------------------------------------------------------------
REM Final da Function IsEditFormOption

REM -------------------------------------------------------------------------
REM Testa se a opcao atual é do edit form
REM -------------------------------------------------------------------------
Private Function IsEditGridOption
  IsEditGridOption = (sparEditOption = conOptionGrid)

End Function
REM -------------------------------------------------------------------------
REM Final da Function IsEditFormOption

REM -------------------------------------------------------------------------
REM Rotina que testa se um campo e externo a uma tabela
REM -------------------------------------------------------------------------
Public Function IsExternalField

  IsExternalField = (Session(EditCurrentField & "ExternalAddress") > "")

End Function
REM -------------------------------------------------------------------------
REM Final da Function IsExternalField

REM -------------------------------------------------------------------------
REM Rotina que retorna a existencia de tabs na tela de edicao
REM -------------------------------------------------------------------------
Public Function EditHaveTabs
  EditHaveTabs = (InStr(Session(EditCurrentTable & "FieldUnit") ,"|") > 0)

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditHaveTab

REM -------------------------------------------------------------------------
REM Rotina que retorna a existencia de dados no texto de formato
REM -------------------------------------------------------------------------
Public Function EditIsThereOperation
  EditIsThereOperation = (sstrText > "")

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditIsThereOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna a existencia de dados no texto de formato
REM -------------------------------------------------------------------------
Public Function EditIsThereNumber
  EditIsThereNumber    = (sstrText2 > "")

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditIsThereOperation

%>