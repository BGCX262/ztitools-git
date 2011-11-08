<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Counter.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Contador de intero de usuarios que acessam uma determinada
REM                pagina em um determinado diretorio.
REM  Cria‡ao     : 00:20h 08/03/97
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1997 by Zevallos(R) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------


REM -------------------------------------------------------------------------
REM  interic = PageCounter(cCounterFile as Char)
REM -------------------------------------------------------------------------
REM  Entradas    : cCounterFile = Diretorio e arquivo do contador
REM                ("Default.cnt")
REM                cDigit = Diretorio e inicial do arquivo .GIF
REM                ("Default")
REM  Saidas      : Resultado do contador   
REM -------------------------------------------------------------------------

Public Function PageCounter(ByVal strCounterFile)
  Dim objFS, objCounter
  Dim intCounterResult, intPageCounter
  
  REM -----------------------------------------------------------------------
  REM Le o arquivo
  REM -----------------------------------------------------------------------
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objCounter = objFS.OpenTextFile(strCounterFile, 1, FALSE)
  
  intCounterResult = objCounter.ReadLine
  objCounter.Close
  
  intCounterResult = FormatNumber(intCounterResult, 0) + 1
  
  PageCounter = intCounterResult
  
  REM -----------------------------------------------------------------------
  REM Grava no arquivo
  REM -----------------------------------------------------------------------
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objCounter = objFS.CreateTextFile(strCounterFile, True)
  
  objCounter.WriteLine intCounterResult
  objCounter.Close
  
  Set objFS = Nothing
  Set objCounter = Nothing

End Function

REM -------------------------------------------------------------------------
REM Fim do Function PageCounter

REM -------------------------------------------------------------------------
REM  interic = ShowCounter(cCounterFile as Char)
REM -------------------------------------------------------------------------
REM  Entradas    : cCounterFile = Arquivo do contador
REM                ("Default.cnt")
REM                cDigit = Tipo do digitos
REM                ("Green")
REM  Saidas      : Resultado do contador   
REM -------------------------------------------------------------------------
Function ShowCounter(strCounterFile, strDigit)
  Dim objFS, objCounter
  Dim intCounterResult

  REM On Error Resume Next

  REM -----------------------------------------------------------------------
  REM Le o arquivo
  REM -----------------------------------------------------------------------
  Set objFS = CreateObject("Scripting.FileSystemObject")
  Set objCounter = objFS.OpenTextFile(strCounterFile, 1, FALSE)

  intCounterResult = objCounter.ReadLine
  objCounter.Close

  ShowCounter = intCounterResult

  ShowNumbers intCounterResult, strDigit

End Function

REM -------------------------------------------------------------------------
REM Fim do Function PageCounter

REM ------------------------------------------------------------------------
REM Fim do Counter.inc
%>