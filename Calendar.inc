<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Calendar.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Calandarios
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
REM Monta o Calendario para a data
REM -------------------------------------------------------------------------
Private Sub ShowCalendar(ByVal dtmDate, ByVal blnAHRef, ByVal intStart, ByVal intEnd)
  Dim i
  Dim intMonth, intWeekDay, intDay, intYear
  Dim dtmBeginMonth
  Dim blnFirst


  intMonth   = Month(dtmDate)
  intDay     = Day(dtmDate)
  intYear    = Year(dtmDate)

  dtmBeginMonth = DateSerial(intYear, intMonth, 1)
  intWeekDay    = WeekDay(dtmBeginMonth)
  
  TableBegin GetMonthName(intMonth) & "/" & intYear, 7, ""

  TableBeginHeadRow 1
  TableHeadColumn "Dom"
  TableHeadColumn "Seg"
  TableHeadColumn "Ter"
  TableHeadColumn "Qua"
  TableHeadColumn "Qui"
  TableHeadColumn "Sex"
  TableHeadColumn "Sab"
  TableEndHeadRow

  TableBeginRow 1

  TableSetColumnAlign "RIGHT"

  If Not intWeekDay = vbSunday Then
    For i = vbSunday To intWeekDay - 1
      TableBeginColumn
      ShowHTMLCR "&nbsp;"
      TableEndColumn
      
    Next
  End If

  blnFirst = False

  For i=1 To GetMonthDays(intMonth)
    If WeekDay(dtmBeginMonth + i - 2) = vbSaturday And blnFirst Then
        TableEndRow
        TableBeginRow 1
        
    End If
    
    blnFirst = True

    If intDay = i Or (i >= intStart And i <= intEnd)  Then
      TableSetColumnBGColor conCalendarMarker

      TableBeginColumn
      
    Else
      TableBeginColumn
     
    End If

    If blnAHref Then
      ShowHTML "<A HREF=" & sstrThisScriptName & "?Date=" & Server.URLEncode(dtmBeginMonth + i - 1) & ">"
      
    End If

    ShowHTML i

    If blnAHref Then
      ShowHTML "</A>"
      
    End If

REM    ShowHTMLCR "</TD>"
      TableEndColumn
      TableSetColumnBGColor ""

  Next
  
  intWeekDay = vbSaturday - WeekDay(dtmBeginMonth + i - 2)

  If intWeekDay > 0 Then
    For i = 1 To intWeekDay
      TableBeginColumn
      ShowHTMLCR "&nbsp;"
      TableEndColumn
      
    Next
  End If

  TableEndRow

  TableEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ShowCalendar

REM -------------------------------------------------------------------------
REM Fim do Calendar.inc
%>