<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Error.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Errors
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
REM Manipula os erros de execucao
REM -------------------------------------------------------------------------
Public Function ErrorHandler

  If Err.Number > 0 Then
    ShowMessageError "Erro com o VBScript:"

    ShowItemListBegin "100%", 1
    ShowItemListError "Numero", Err.Number
    ShowItemListError "Descri&ccedil;&atilde;o.", Err.Description
    ShowItemListError "Help Context", Err.HelpContext
    ShowItemListError "Help Path", Err.HelpPath
    ShowItemListError "Native Error", Err.NativeError
    ShowItemListError "Source", Err.Source
    ShowItemListError "SQLState", Err.SQLState
    ShowItemListEnd

  End If
  
  ErrorHandler = Err.Number

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ErrorHandler

REM -------------------------------------------------------------------------
REM Manipula os erros de execucao do Objeto Conn
REM -------------------------------------------------------------------------
Public Function ErrorConnHandler(objConn)
  Dim i
  
  On Error Goto 0
  
  ErrorConnHandler = ErrorHandler
  
  IF objConn.Errors.Count > 0 then
    ShowMessageError objConn.Errors.Count & " - Erro(s) com a Base de dados:"

    ShowItemListBegin "100%", 1

    For i = 0 To objConn.Errors.Count
      ShowItemListError i + 1  &" - N&uacute;mero", objConn.Errors(i).Number
      ShowItemListError "Descri&ccedil;&atilde;o", objConn.Errors(i).Description
     
    Next
    
    On Error Resume Next

    ShowItemListError "Source", objConn.Source
    ShowItemListError "SQLState", objConn.SQLState
    
    ShowItemListEnd

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ErrorConnHandler

REM -------------------------------------------------------------------------
REM Fim do Error.inc
%>