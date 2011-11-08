<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Table.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Tabelas
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

Dim sTableSetBGColor
Dim sTableSetHeadRowBGColor
Dim sTableSetRowBGColor
Dim sTableRowFontSize
Dim sTableSetBorder
Dim sTableSetColumnWidth
Dim sTableSetColumnHeight
Dim sTableSetColumnAlign
Dim sTableSetColumnColSpan
Dim sTableSetColumnRowSpan
Dim sTableSetColumnVAlign
Dim sTableSetColumnBGColor
Dim sTableSetColumnNoWrap
DIM sTableSetSpacing
DIM sTableSetPadding

Public sintGraphNum, sintGraphActual, sarrmax(5), sarrtotal(5)

sintGraphNum = 0
sintGraphActual = 1
REM TableSetBorder "0"
REM TableSetSpacing "0"
REM TableSetPadding "0"

REM -------------------------------------------------------------------------
REM Inicio da tabela
REM -------------------------------------------------------------------------
Private Sub TableBegin(ByVal strMessage, ByVal intColumns, ByVal strWidth)

  On Error Resume Next

  If Not sTableSetHeadRowBGColor > "" Then
    TableSetHeadRowBGColor "#CCCCFF"
    TableSetHeadRowBGColor conTableHeadBGColor

  End If

  If Not sTableSetRowBGColor > "" Then
    TableSetRowBGColor "#ECECFF"
    TableSetRowBGColor conTableBGColor

  End if

  On Error Goto 0
                             
  Response.Write "<TABLE BGCOLOR=Black CELLSPACING=0 CELLPADDING=0"
REM White CELLSPACING=0 CELLPADDING=0" 
REM Black CELLSPACING=0 CELLPADDING=0"

  If strWidth > "" Then
    Response.Write  " WIDTH=""" & strWidth & """"
  
  End If
  
  Response.Write " BORDER="

  If Session("TableBorder") > "" Then
    Response.Write Session("TableBorder")
    
  Else
    Response.Write "0"
    
  End IF  

  ShowHTMLCR ">"
  
  ShowHTMLCR "<TR><TD VALIGN=TOP>"
                  
  ShowHTMLCR "<TABLE BGCOLOR=Black WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3>"

  If strMessage > "" Then
    Response.Write "<TR BGCOLOR=White><TD ALIGN=CENTER"
  REM  " & sTableSetHeadRowBGColor & "><TD ALIGN=CENTER"
  REM  White><TD ALIGN=CENTER"
    
    If intColumns > 1 Then
      Response.Write " COLSPAN=" & intColumns
      
    End If
    
    ShowHTMLCR "><FONT SIZE=3><B>" & strMessage & "</B></FONT></TD></TR>"
  
  End If

  TableSetColumnAlign ""    
  TableSetColumnVAlign ""    
  TableSetColumnBGColor ""
  TableSetColumnColSpan ""
  TableSetColumnWidth ""
  TableSetColumnHeight ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableBegin

REM -------------------------------------------------------------------------
REM Fim da tabela
REM -------------------------------------------------------------------------
Private Sub TableEnd

  ShowHTMLCR "</TABLE></TD></TR></TABLE>"

  TableSetBorder "0"
  TableSetSpacing "0"
  TableSetPadding "0"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableEnd

REM -------------------------------------------------------------------------
REM Inicio da tabela
REM -------------------------------------------------------------------------
Private Sub TableNormalBegin(ByVal strWidth)

  On Error Resume Next

  If Not sTableSetHeadRowBGColor > "" Then
    TableSetHeadRowBGColor "#CCCCFF"
    TableSetHeadRowBGColor = conTableHeadBGColor
    
  End If

  If Not sTableSetRowBGColor > "" Then
    TableSetRowBGColor "#ECECFF"
    TableSetRowBGColor = conTableBGColor
    
  End if

  On Error Goto 0

  Response.Write "<TABLE"
  
  If strWidth > "" Then 
    Response.Write " WIDTH=" & strWidth
    
  End If

  If Session("TableBorder") > "" And Session("TableBorder") <> "0" Then
    Response.Write " BORDER=" & Session("TableBorder")
    
  Else
    If sTableSetBorder > "" Then
      Response.Write " BORDER=" & sTableSetBorder

    Else
      Response.Write " BORDER=0"

    End If
    
  End IF  

  If sTableSetSpacing > "" Then
    Response.Write " CELLSPACING=" & sTableSetSpacing
  
  Else
    Response.Write " CELLSPACING=0"
    
  End If

  If sTableSetPadding > "" Then
    Response.Write " CELLPADDING=" & sTableSetPadding
    
  Else
    Response.Write " CELLPADDING=0"
    
  End If

  If sTableSetBGColor > "" Then
    Response.Write " BGCOLOR=" & sTableSetBGColor
    
  End If
  
  Response.Write ">"
    
  TableSetColumnAlign ""    
  TableSetColumnVAlign ""    
  TableSetColumnBGColor ""
  TableSetColumnColSpan ""
  TableSetColumnWidth ""
  TableSetColumnHeight ""
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableBegin

REM -------------------------------------------------------------------------
REM Fim da tabela
REM -------------------------------------------------------------------------
Private Sub TableNormalEnd

  Response.Write "</TABLE>"

  TableSetBorder "0"
  TableSetSpacing "0"
  TableSetPadding "0"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableEnd

REM -------------------------------------------------------------------------
REM Cria o Table Border
REM -------------------------------------------------------------------------
Private Sub TableSetBorder(ByVal strValue)

  sTableSetBorder = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetBorder

REM -------------------------------------------------------------------------
REM Cria o Table Border
REM -------------------------------------------------------------------------
Private Sub TableSetSpacing(ByVal strValue)

  sTableSetSpacing = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetBorder

REM -------------------------------------------------------------------------
REM Cria o Table Border
REM -------------------------------------------------------------------------
Private Sub TableSetPadding(ByVal strValue)

  sTableSetPadding = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetPadding

REM -------------------------------------------------------------------------
REM Inicio de uma linha
REM -------------------------------------------------------------------------
REM conTableColumnBGColor

Private Sub TableBeginRow(ByVal intFontSize)

    sTableRowFontSize = intFontSize

    Response.Write "<TR"

    If sTableSetRowBGColor > "" Then
      Response.Write " BGCOLOR=" & sTableSetRowBGColor
      
    End If
    
     ShowHTMLCR ">"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableBeginRow

REM -------------------------------------------------------------------------
REM Fim de uma linha
REM -------------------------------------------------------------------------
Private Sub TableEndRow

    Response.Write "</TR>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableEndRow

REM -------------------------------------------------------------------------
REM Cria o Table Border
REM -------------------------------------------------------------------------
Private Sub TableSetRowBGColor(ByVal strValue)
  sTableSetRowBGColor = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetRowBGColor

REM -------------------------------------------------------------------------
REM Cria o Table BGColor
REM -------------------------------------------------------------------------
Private Sub TableSetBGColor(ByVal strValue)
  sTableSetBGColor = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetBGColor

REM -------------------------------------------------------------------------
REM Inicio de uma linha
REM -------------------------------------------------------------------------
REM conTableHeadBGColor

Private Sub TableBeginHeadRow(ByVal intFontSize)

    sTableRowFontSize = intFontSize

    Response.Write "<TR"

    If sTableSetHeadRowBGColor > "" Then
      Response.Write " BGCOLOR=" & sTableSetHeadRowBGColor
      
    End If
    
     ShowHTMLCR ">"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableBeginHeadRow

REM -------------------------------------------------------------------------
REM Cria o Table Border
REM -------------------------------------------------------------------------
Private Sub TableSetHeadRowBGColor(ByVal strValue)

  sTableSetHeadRowBGColor = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetHeadRowBGColor

REM -------------------------------------------------------------------------
REM Fim de uma linha
REM -------------------------------------------------------------------------
Private Sub TableEndHeadRow

    Response.Write "</TR>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableEndHeadRow

REM -------------------------------------------------------------------------
REM Fim da tabela
REM -------------------------------------------------------------------------
Private Sub TableHeadColumn(ByVal strText)

    Response.Write "<TH"

    If sTableSetColumnHeight > "" Then
      Response.Write " HEIGHT=" & sTableSetColumnHeight
      
    End If

    If sTableSetColumnWidth > "" Then
      Response.Write " WIDTH=" & sTableSetColumnWidth
      
    End If

    If sTableSetColumnAlign > "" Then
      Response.Write " ALIGN=" & sTableSetColumnAlign
      
    End If

    If sTableSetColumnColSpan > "" Then
      Response.Write " COLSPAN=" & sTableSetColumnColSpan
      
    End If

    If sTableSetColumnRowSpan > "" Then
      Response.Write " ROWSPAN=" & sTableSetColumnRowSpan
      
    End If

    If sTableSetColumnVAlign > "" Then
      Response.Write " VALIGN=" & sTableSetColumnVAlign
      
    End If

    If sTableSetColumnBGColor > "" Then
      Response.Write " BGCOLOR=" & sTableSetColumnBGColor
      
    End If

   
    Response.Write ">"

    If sTableRowFontSize > "" Then
      Response.Write "<FONT SIZE=" & sTableRowFontSize & ">"
    
    End If

    Response.Write strText
    
    If sTableRowFontSize > "" Then
REM      Response.Write "</FONT>"
    
    End If

    ShowHTMLCR "</TH>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableHeadColumn

REM -------------------------------------------------------------------------
REM Fim da tabela
REM -------------------------------------------------------------------------
Private Sub TableBeginHeadColumn

    Response.Write "<TH"

    If sTableSetColumnHeight > "" Then
      Response.Write " HEIGHT=" & sTableSetColumnHeight
      
    End If

    If sTableSetColumnWidth > "" Then
      Response.Write " WIDTH=" & sTableSetColumnWidth
      
    End If

    If sTableSetColumnAlign > "" Then
      Response.Write " ALIGN=" & sTableSetColumnAlign
      
    End If

    If sTableSetColumnColSpan > "" Then
      Response.Write " COLSPAN=" & sTableSetColumnColSpan
      
    End If

    If sTableSetColumnRowSpan > "" Then
      Response.Write " ROWSPAN=" & sTableSetColumnRowSpan
      
    End If

    If sTableSetColumnVAlign > "" Then
      Response.Write " VALIGN=" & sTableSetColumnVAlign
      
    End If

    If sTableSetColumnBGColor > "" Then
      Response.Write " BGCOLOR=" & sTableSetColumnBGColor
      
    End If

   
    Response.Write ">"

    If sTableRowFontSize > "" Then
      Response.Write "<FONT SIZE=" & sTableRowFontSize & ">"
    
    End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableHeadColumn

REM -------------------------------------------------------------------------
REM Fim da tabela
REM -------------------------------------------------------------------------
Private Sub TableEndHeadColumn

    If sTableRowFontSize > "" Then
REM      Response.Write "</FONT>"
    
    End If

    Response.Write "</TH>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableHeadColumn

REM -------------------------------------------------------------------------
REM Inicio de uma coluna
REM -------------------------------------------------------------------------
Private Sub TableBeginColumn

    Response.Write "<TD"

    If sTableSetColumnHeight > "" Then
      Response.Write " HEIGHT=" & sTableSetColumnHeight
      
    End If

    If sTableSetColumnWidth > "" Then
      Response.Write " WIDTH=" & sTableSetColumnWidth
      
    End If

    If sTableSetColumnAlign > "" Then
      Response.Write " ALIGN=" & sTableSetColumnAlign
      
    End If

    If sTableSetColumnColSpan > "" Then
      Response.Write " COLSPAN=" & sTableSetColumnColSpan
      
    End If

    If sTableSetColumnRowSpan > "" Then
      Response.Write " ROWSPAN=" & sTableSetColumnRowSpan
      
    End If

    If sTableSetColumnVAlign > "" Then
      Response.Write " VALIGN=" & sTableSetColumnVAlign
      
    End If

    If sTableSetColumnBGColor > "" Then
      Response.Write " BGCOLOR=" & sTableSetColumnBGColor
      
    End If

    If sTableSetColumnNoWrap Then
      Response.Write " NOWRAP=NOWRAP"
      
    End If

    Response.Write ">"
    
    If sTableRowFontSize > "" OR Session("TableSetColumnColor") > "" Then
      Response.Write "<FONT"
      
      If Session("TableSetColumnColor") > "" Then
        Response.Write " COLOR=" & Session("TableSetColumnColor")
      
      End If

      If sTableRowFontSize > "" Then
          Response.Write " SIZE=" & sTableRowFontSize
      
      End If
      
      ShowHTMLCR ">"
    
    End If


End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableBeginColumn

REM -------------------------------------------------------------------------
REM Cria o Height
REM -------------------------------------------------------------------------
Private Sub TableSetColumnHeight(ByVal strValue)

  sTableSetColumnHeight = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnHeight

REM -------------------------------------------------------------------------
REM Cria o Width
REM -------------------------------------------------------------------------
Private Sub TableSetColumnWidth(ByVal strValue)

  sTableSetColumnWidth = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnWidth

REM -------------------------------------------------------------------------
REM Cria o Align
REM -------------------------------------------------------------------------
Private Sub TableSetColumnColSpan(ByVal strValue)

  sTableSetColumnColSpan = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnColSpan

REM -------------------------------------------------------------------------
REM Cria o Align
REM -------------------------------------------------------------------------
Private Sub TableSetColumnRowSpan(ByVal strValue)

  sTableSetColumnRowSpan = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnColSpan

REM -------------------------------------------------------------------------
REM Cria o Align
REM -------------------------------------------------------------------------
Private Sub TableSetColumnAlign(ByVal strValue)

  sTableSetColumnAlign = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnAlign

REM -------------------------------------------------------------------------
REM Cria o Align
REM -------------------------------------------------------------------------
Private Sub TableSetColumnVAlign(ByVal strValue)

  sTableSetColumnVAlign = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnAlign

REM -------------------------------------------------------------------------
REM Define o NoWrap
REM -------------------------------------------------------------------------
Private Sub TableSetColumnNoWrap(ByVal blnState)

  sTableSetColumnNoWrap = blnState

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnAlign

REM -------------------------------------------------------------------------
REM Cria o Align
REM -------------------------------------------------------------------------
Private Sub TableSetColumnBGColor(ByVal strValue)

  sTableSetColumnBGColor = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnAlign

REM -------------------------------------------------------------------------
REM Cria o Align
REM -------------------------------------------------------------------------
Private Sub TableSetColumnColor(ByVal strValue)

  Session("TableSetColumnColor") = strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableSetColumnAlign

REM -------------------------------------------------------------------------
REM Fim da tabela
REM -------------------------------------------------------------------------
Private Sub TableEndColumn

    If sTableRowFontSize > "" Then
REM      Response.Write "</FONT>"
    
    End If

    Response.Write "</TD>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableEndColumn

REM -------------------------------------------------------------------------
REM Retorna cores sequenciais
REM -------------------------------------------------------------------------
Private Function TableGraphColor

 If (Session(sintGraphActual & "sintCor") \ 3) = 6 Then
   Session(sintGraphActual & "sintCor") = 0     
 
 End If    

 Select Case (Session(sintGraphActual & "sintCor") \ 3)
   Case 0
     Session(sintGraphActual & "sintG") = Session(sintGraphActual & "sintG") + 85

   Case 1
     Session(sintGraphActual & "sintR") = Session(sintGraphActual & "sintR") - 85

   Case 2
     Session(sintGraphActual & "sintB") = Session(sintGraphActual & "sintB") + 85

   Case 3
     Session(sintGraphActual & "sintG") = Session(sintGraphActual & "sintG") - 85
       
   Case 4
     Session(sintGraphActual & "sintR") = Session(sintGraphActual & "sintR") + 85
       
   Case 5
     Session(sintGraphActual & "sintB") = Session(sintGraphActual & "sintB") - 85

 End Select     
 Session(sintGraphActual & "sintCor") = Session(sintGraphActual & "sintCor") + 1

 TableGraphColor = "#" & LeadingZeroes(Hex(Session(sintGraphActual & "sintR")), 2) & LeadingZeroes(Hex(Session(sintGraphActual & "sintG")), 2) & LeadingZeroes(Hex(Session(sintGraphActual & "sintB")), 2)
  
End Function
REM -------------------------------------------------------------------------
REM Final da Sub TableGraphColor

REM -------------------------------------------------------------------------
REM Desenha as Barras do Graficos
REM -------------------------------------------------------------------------
Private Sub TableColumnGraph(intParc, strLink)
Dim strCor, intLargura

  TableBeginColumn
  
  If sarrTotal(sintGraphActual) > 0 Then
    TableNormalBegin 300
  
  Else
    TableNormalBegin 250
  
  End If

  TableSetColumnWidth 70

  TableSetColumnAlign "Right" 
  TableBeginRow 1    

  If strLink > "" Then
    TableCompleteColumn strLink & "&nbsp;"
REM    TableCompleteColumn "&nbsp;&nbsp;<a href=" & sstrThisScriptName & "?" & conPOption & "=" & conOptionSession & strLink & "><FONT SIZE=-1><B>" & ZTINumber(100 * intParc / sarrTotal(sintGraphActual)) & "%" & "</a>&nbsp;&nbsp;"
  
  Else
    TableCompleteColumn "&nbsp;" & intParc & "&nbsp;"
  
  End If                

  TableSetColumnWidth 50
  If CLng(sarrTotal(sintGraphActual)) > 0 Then
    TableCompleteColumn "&nbsp;" & ZTINumber(100 * intParc / sarrTotal(sintGraphActual)) & "%" & "&nbsp;"
  
  End If

  TableSetColumnAlign "Left" 

  TableSetBorder 1

  TableSetColumnWidth ""

  TableBeginColumn

  Randomize
  
  strCor = sTableSetRowBGColor
           
  TableSetRowBGColor TableGraphColor
  TableSetBorder ""
  intLargura = 180 * intParc \ sarrMax(sintGraphActual)
  If intLargura < 1 Then
    intLargura = 1
    
  End If
    
  TableNormalBegin intLargura

  TableBeginRow 1  
  TableSetColumnHeight "15"
  TableCompleteColumn ""
  TableEndRow
  
  TableNormalEnd

  TableEndRow

  TableNormalEnd
  
  TableEndColumn
  
  TableSetColumnAlign ""
  
  TableSetColumnHeight ""
  TableSetBorder ""
  TableSetRowBGColor strCor
  If sintGraphActual = sintGraphNum Then
    sintGraphActual = 1
    
  Else
    sintGraphActual = sintGraphActual + 1
  
  End If  

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableColumnGraph

REM -------------------------------------------------------------------------
REM Desenha a Legenda do Grafico de Barras
REM -------------------------------------------------------------------------
Private Sub TableGraphLegend(intMax, intTotal)
Dim strCor
Dim i

  sintGraphNum = sintGraphNum + 1
  sintGraphActual = sintGraphNum
  If Session(sintGraphActual & "sintCor") = "" Then
    For i = 1 to sintGraphNum
      Session(i & "sintCor") = -1
      Session(i & "sintR")   = 255
      Session(i & "sintG")   = -85
      Session(i & "sintB")   = 0

    Next

  End If

  strCor = sTableSetRowBGColor
  TableSetRowBGColor sTableSetHeadRowBGColor

  sarrTotal(sintGraphActual) = intTotal
  sarrMax(sintGraphActual) = intMax
  TableBeginColumn

  If sarrTotal(sintGraphActual) > 0 Then
    TableNormalBegin 300
    TableSetColumnWidth 120

  Else
    TableNormalBegin 250
    TableSetColumnWidth 70

  End If

  TableBeginRow 1
                                      
  TableSetColumnAlign "Left"
  If intTotal > 0 Then
    TableCompleteColumn "Total"

  Else
    TableCompleteColumn "&nbsp;"
    
  End If

  TableSetColumnAlign "Left"
  If intMax > 2 Then
    TableSetColumnWidth 60
    TableSetColumnAlign "Left"
    TableCompleteColumn "0"
    TableSetColumnAlign "Center"
    If intMax < 10 Then
      TableCompleteColumn Round(intMax / 2, 1)
      
    Else
      TableCompleteColumn intMax \ 2
    
    End If  
    TableSetColumnAlign "Right"
    TableCompleteColumn intMax
    TableEndRow

    TableBeginRow 1                                      
    
    TableSetColumnAlign "Center"
    TableSetColumnWidth 120
    If intTotal > 0 Then
       TableCompleteColumn "&nbsp;<Font Size=-1><B>" & sarrTotal(sintGraphActual) & "<B></Font>&nbsp;"

    Else
       TableCompleteColumn "&nbsp;"
    
    End If

    TableSetColumnWidth 60
    TableSetColumnAlign "Left"
    TableCompleteColumn "|"
    TableSetColumnAlign "Center"
    TableCompleteColumn "|"
    TableSetColumnAlign "Right"
    TableCompleteColumn "|"
  
  Else
    TableSetColumnWidth 180 \ (intMax + 1)
    TableSetColumnAlign "Left"
    For i = 0 to intMax - 1
      TableCompleteColumn i
      TableSetColumnAlign "Center"
      
    Next

    TableSetColumnAlign "Right"
    TableCompleteColumn intMax
    TableEndRow

    TableBeginRow 1

    TableSetColumnAlign "Center"
    TableSetColumnWidth 120
    If intTotal > 0 Then
       TableCompleteColumn "&nbsp;<Font Size=-1><B>" & sarrTotal(sintGraphActual) & "<B></Font>&nbsp;"

    Else
       TableCompleteColumn "&nbsp;"
    
    End If
  
    TableSetColumnAlign "Left"
    TableSetColumnWidth 180 \ (intMax + 1)
    For i = 0 to intMax - 1
      TableCompleteColumn "|"
      TableSetColumnAlign "Center"
      
    Next

    TableSetColumnAlign "Right"
    TableCompleteColumn "|"

  End If  
  TableEndRow
    
  TableNormalEnd
  
  TableEndColumn
  
  TableSetColumnAlign ""
  TableSetColumnWidth ""
  TableSetRowBGColor strCor

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableGraphLegend

REM -------------------------------------------------------------------------
REM Mostra a Lista de Acessos
REM -------------------------------------------------------------------------
Private Sub TableCompleteColumn(strTexto)
  TableBeginColumn
  ShowHTMLCR strTexto
  TableEndColumn

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableCompleteColumn

REM -------------------------------------------------------------------------
REM Fim do Table.inc
%>