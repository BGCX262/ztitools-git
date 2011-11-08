<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Summary.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 11:30h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Recolhe as informacoes de um resumo
REM -------------------------------------------------------------------------
Private Sub EditSummaryPars(strText)
Dim blnParameter, i, j

  blnParameter = False

  Session("FieldNumber") = 0
  Session("TotalNumber") = 0
  sblnHaveSum   = False
  sblnHaveGraph = False

  EditBeginReadOperation strText

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",",";","=","-"

      Case "("
        blnParameter = True

      Case ")"
        blnParameter = False

      Case "#"
        Session("FieldNumber") = Session("FieldNumber") + 1
        Session(Session("FieldNumber") & "IsCalcField") = "1"
        Session(Session("FieldNumber") & "ShowField")   = EditNextOperation
        Session(Session("FieldNumber") & "Field")       = EditLastOperation
        Session(Session("FieldNumber") & "Value")       = 0
        ClearSession Session("FieldNumber") & "Type"
        ClearSession Session("FieldNumber") & "IsGraph"
        ClearSession Session("FieldNumber") & "IsSum"
        ClearSession Session("FieldNumber") & "IsCount"
        ClearSession Session("FieldNumber") & "Operation"
        ClearSession Session("FieldNumber") & "Have1stValue"

      Case Else
        If blnParameter Then
          Select Case LCase(EditLastOperation)
            Case "s"
              Session(Session("FieldNumber") & "IsSum")    = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "SUM(" & Session(Session("FieldNumber") & "Field") & ")"
              sblnHaveSum                        = True

            Case "c"
              Session(Session("FieldNumber") & "IsCount")  = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "COUNT(*)"

            Case "g"
              Session(Session("FieldNumber") & "IsGraph")  = "1"
              sblnHaveGraph                      = True

            Case "gs"
              Session(Session("FieldNumber") & "IsGraph")  = "1"
              Session(Session("FieldNumber") & "IsSum")    = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "SUM(" & Session(Session("FieldNumber") & "Field") & ")"
              sblnHaveSum                        = True
              sblnHaveGraph                      = True

            Case "gc"
              Session(Session("FieldNumber") & "IsGraph")  = "1"
              Session(Session("FieldNumber") & "IsCount")  = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "COUNT(*)"
              sblnHaveSum                        = True
              sblnHaveGraph                      = True

            Case "a"
              sstrGroupField = Session("FieldNumber")

            Case "t"
              Session(Session("FieldNumber") & "IsTotal")  = "1"

            Case "f"
              EditNextOperation
              Session(Session("FieldNumber") & "Type")     = EditNextOperation

            Case "o"
              EditNextOperation
              i = InStr(sstrText,",")
              j = InStr(sstrText,")")
              If i > j Or i = 0 Then
                i = j

              End If
              If i > 0 Then
                Session(Session("FieldNumber") & "Operation") = Left(sstrText, i - 1)
                sstrText                                      = Right(sstrText, Len(sstrText) - i + 1)

              End If

            Case "v"
              EditNextOperation
              i = InStr(sstrText,",")
              j = InStr(sstrText,")")
              If i > j Or i = 0 Then
                i = j

              End If
              If i > 0 Then
                Session(Session("FieldNumber") & "Value")     = Left(sstrText, i - 1)
                sstrText                                      = Right(sstrText, Len(sstrText) - i + 1)
                Session(Session("FieldNumber") & "Have1stValue")= "1"

              End If

            Case "l"
              EditNextOperation
              Session(Session("FieldNumber") & "CaptionField") = EditNextOperation

          End Select
        Else
          Session("FieldNumber") = Session("FieldNumber") + 1
          If EditFindField(EditCurrentTable, EditLastOperation) Then
            Session(Session("FieldNumber") & "ShowField") = Session("CurrentField")
            If Session(EditCurrentField & "EditOption") = conLookupField Then
              Session(Session("FieldNumber") & "Field") = Session(EditCurrentField & "LookupFieldText")

            Else
              Session(Session("FieldNumber") & "Field") = Session(EditCurrentField & "Field")

            End If
            Session(Session("FieldNumber") & "CaptionField") = EditCaption(False)
            Session(Session("FieldNumber") & "OrderSQL")  = Session("CurrentField")
            Session(Session("FieldNumber") & "Type")      = Session(EditCurrentField & "Type")
            Session(Session("FieldNumber") & "Value")     = 0
            ClearSession Session("FieldNumber") & "IsGraph"
            ClearSession Session("FieldNumber") & "IsSum"
            ClearSession Session("FieldNumber") & "IsCount"
            ClearSession Session("FieldNumber") & "IsCalcField"
            ClearSession Session("FieldNumber") & "Operation"

          End If

        End If

    End Select

  Loop

  sstrFields  = ""
  sstrFields2 = ""
Dim blnFirst, blnFirst2

  blnFirst    = True
  blnFirst2   = True
  For i = 1 to Session("FieldNumber")
    If Not blnFirst And Not Session(i & "Operation") > "" Then
      sstrFields = sstrFields & ", "

    ElseIf Not Session(i & "Operation") > "" Then
      blnFirst    = False

    End If

    If Not blnFirst2 And Not Session(i & "IsCalcField") > "" Then
      sstrFields2 = sstrFields2 & ", "

    ElseIf Not Session(i & "IsCalcField") > "" Then
      blnFirst2    = False

    End If

    If Session(i & "IsSum") > "" Then
      sstrFields = sstrFields & "SUM(" & Session(i & "Field") & ") as s" & Session(i & "Field")
      If Not Session(i & "IsCalcField") > "" Then
        sstrFields = sstrFields & ", " & Session(i & "Field")
        sstrFields2 = sstrFields2 & Session(i & "Field")

      End if

    ElseIf Session(i & "IsCount") > "" Then
      sstrFields = sstrFields & "COUNT(*) as s" & Session(i & "Field")
      If Not Session(i & "IsCalcField") > "" Then
        sstrFields = sstrFields &  ", " & Session(i & "Field")

      End if

    ElseIf Not Session(i & "IsCalcField") > "" Then
      sstrFields = sstrFields & Session(i & "Field")
      sstrFields2 = sstrFields2 & Session(i & "Field")
      If Session(EditCurrentTable & Session(i & "ShowField") & "EditOption") = conLookupField Then
        sstrFields = sstrFields & ", " & Session(EditCurrentTable & Session(i & "ShowField") & "Field")
        sstrFields2 = sstrFields2 & ", " & Session(EditCurrentTable & Session(i & "ShowField") & "Field")
        blnFirst    = False
        blnFirst2   = True

      End if
 
    End If

  Next
REM  sobjShow.MessageError sstrFields
REM  sobjShow.MessageError sstrFields2

End Sub

REM -------------------------------------------------------------------------
REM Listagem com resumo
REM -------------------------------------------------------------------------
Public Sub EditShowSummary
Dim sql, i, strSelect, strFrom, strJoin, strWhere, strGroup, strOrder, strSummary, strFields, intNumCols, strTip2

  If EditFindList(sparEditList) Then
    EditSummaryPars Session(EditCurrentList & "List")
    intNumCols = Session("FieldNumber")

    strJoin = ""
    For i = 1 to Session("FieldNumber")
REM Terminar
      If Not Session(i & "IsSum") > "" And Not Session(i & "IsCount") > "" And Not Session(i & "IsCalcField") > "" Then
        Session("CurrentField") = Session(i & "ShowField")
        If Session(EditCurrentField & "EditOption") = conLookupField Then
           strJoin = strJoin & " LEFT OUTER JOIN " & Session(EditCurrentField & "LookupTable") &_
           " on " & Session(EditCurrentField & "LookupTable") & "." & Session(EditCurrentField & "LookupField") & " = " & _
           EditTableOf(EditCurrentTable, Session("CurrentField")) & "." & Session(EditCurrentField & "Field")

        End If
     End If
    Next
    
    strWhere = Filter

    strGroup = " GROUP BY " & sstrFields2

    If Not Session("OrderField") > "" Then
      EditBeginReadOperation sstrFields

      Do While EditIsThereOperation
        If EditFindField(EditCurrentTable, EditNextOperation) Then
          sstrText = ""
          Session("OrderField") = Session("CurrentField")
          strOrder = " ORDER BY " & Session("CurrentField")

        End If

      Loop

    End If

    Dim blnFirst
    If IsNumeric(Session("OrderField")) Then
      EditBeginReadOperation Session("OrderField")
      strOrder = " ORDER BY "

      blnFirst = True

      Do While EditIsThereOperation
        Select Case EditNextOperation
           Case ","

           Case Else
             Session("CurrentField") = EditLastOperation
             If Not blnFirst Then
               strOrder = strOrder & ", "

             Else
               strOrder = strOrder & " "

             End If

             Select Case Session(EditCurrentField & "EditOption")
               Case conLookUpField
                 strOrder = strOrder & Session(EditCurrentField & "LookupTable") & "." & Session(EditCurrentField & "LookupFieldText")

               Case Else
                 strOrder = strOrder & Session(EditCurrentField & "Field")

            End Select

            blnFirst = False
            If Session("OrderDesc") = "1" Then
              strOrder = strOrder & " DESC "

            End If

        End Select

      Loop

    Else
      If Session("OrderField") > "" Then
        strOrder = " ORDER BY " & Session("OrderField")
        If Session("OrderDesc") = "1" Then
          strOrder = strOrder & " DESC "

        End If
      End If
    End If

    Dim intWidth

    intWidth = 0

    For i = 1 to Session("FieldNumber")
      Session("CurrentField") = Session(i & "ShowField")
      intWidth = intWidth + EditSize + 6

    Next

    If intWidth < 70 Then
      intWidth = 70

    ElseIf intWidth > 95 Then
      intWidth = 95

    End If

    sobjBrowse.Actual = EditCurrentTable

    ClassAss

    Dim intTotal(30), intMax(30)

    strSelect = "SELECT "

    strFrom = " FROM " & EditCurrentTable

    blnFirst = True
    If sblnHaveSum Then
      For i = 1 to Session("FieldNumber")
        If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
          If Not blnFirst Then
            sql = sql & ", "

          End If
          blnFirst = False

          If Session(i & "IsSum") > "" Then
            sql = sql & " SUM(" & Session(i & "Field") & ") as s" & Session(i & "Field")

          Else
            sql = sql & " COUNT(*) as s" & Session(i & "Field")

          End If

        End If
        Session(i & "MaxValue") = 0
        Session(i & "TotalValue") = 0

      Next

      sql = strSelect & sql & strFrom & strWhere

      sobjCMD.CommandText = sql
      sobjCMD.CommandType = adCmdText

REM      sobjShow.JSAlert sql

      sobjRS2.Open sobjCMD, , adOpenKeySet, adLockReadOnly

      For i = 1 to Session("FieldNumber")
        If Session(i & "IsSum") > ""Or Session(i & "IsCount") > "" Then
          Session(i & "TotalValue") = sobjRS2("s" & Session(i & "Field"))

        End If

      Next
      sobjRS2.Close

    End If

    sql = strSelect & sstrFields & strFrom & strJoin & strWhere & strGroup & strOrder

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

REM  DebugMessage sql, 1

    sobjRS2.Open sobjCMD, , adOpenKeySet, adLockReadOnly

Dim strMax
    If Not sobjRS2.EOF And Not sobjRS2.BOF Then
    Dim strMessage
      If sblnHaveGraph Then
        Do While Not sobjRS2.EOF
          For i = 1 to Session("FieldNumber")
            If Session(i & "IsGraph") > "" Then
              If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
                If Session(i & "MaxValue") < sobjRS2("s" & Session(i & "Field")) Then
                  Session(i & "MaxValue") = sobjRS2("s" & Session(i & "Field"))
REM                  sobjShow.JSAlert "Teste 1"

                End If
              Else
                If Session(i & "MaxValue") < sobjRS2(Session(i & "Field")) Then
                  Session(i & "MaxValue") = sobjRS2(Session(i & "Field"))
REM                  sobjShow.JSAlert "Teste 2"

                End If
              End If
            End If
          Next
          sobjRS2.MoveNext

        Loop
        sobjRS2.MoveFirst

      End If
REM      sobjShow.JSAlert Session(i & "MaxValue")

      strMessage = Session(EditCurrentList & "ListCaption")

      sobjShow.Nbsp
      sobjShow.Center

      sobjBrowse.BeginBrowse sobjRS2
      sobjTabela.Padding = "3"
      sobjTabela.Spacing = "1"
      sobjTabela.BeginTable intWidth & "%", strMessage, intNumCols, True
      sobjTabela.Padding = "1"
      sobjTabela.Spacing = "0"

      strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
      conOptionAZEdit & EditWhatList & EditRemakeTarget & """>Listagem Completa</A><BR>"

      For i = 1 to Session(EditCurrentTable & "ListNumber")
        If Session(EditCurrentTable & i & "ListName") <> sparEditList Then
          strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
          conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditList & "=" & Session(EditCurrentTable & i & "ListName") & """>" & Session(EditCurrentTable & i & "ListCaption") & "</A><BR>"

        End If

      Next

      sobjTabela.ColumnColSpan = intNumCols

      sobjTabela.BeginRow 1
      sobjTabela.BeginColumn
      sobjTabela.BeginTable "100%"
      sobjTabela.ColumnAlign = "Right"

      sobjTabela.BeginRow 1
      sobjTabela.BeginColumn
      Response.Write "<FONT SIZE=2><DIV ID=F2 onMouseOver=""makeCool(F2);"" onMouseOut=""makeNormal(F2);"" onClick=""OpenDesc(T2);"" STYLE=""cursor:hand""><B>Outras listagens</B></DIV>"
      Response.Write "<table><tr><td align=justify>"
      Response.Write "<DIV ID=T2 class=dbody style=""display:'none';""><B>"
      Response.Write strTip2
      Response.Write "</B></DIV></td></tr></table>"
      sobjTabela.EndColumn
      sobjTabela.EndRow

      sobjTabela.EndTable
      sobjTabela.EndColumn
      sobjTabela.EndRow
      sobjTabela.ColumnAlign = ""
      sobjTabela.ColumnColSpan = ""

      sobjTabela.BeginRow 1, True
      For i = 1 to Session("FieldNumber")
REM        sobjShow.MessageError Session(i & "Field") & Session(i & "MaxValue") & "*" & Session(i & "TotalValue") & "$"
        If Session(i & "IsGraph") > "" Then
          Session("CurrentField") = Session(i & "ShowField")
          TableSummaryLegend Session(i & "MaxValue"), Session(i & "TotalValue"), Session(i & "CaptionField"), Session(i & "Type"), Session(i & "OrderSQL")

        ElseIf Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
          EditTableHeader Session(i & "CaptionField"), Session(i & "OrderSQL"), Session(EditCurrentField & "ListSize")

        Else
          Session("CurrentField") = Session(i & "ShowField")
          EditTableHeader Session(i & "CaptionField"), Session(i & "OrderSQL"), Session(EditCurrentField & "ListSize")

        End If

        If sstrGroupField > "" Then
          If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
            Session(i & "SubTotal") = 0
            Session(i & "Total")    = 0

          End If
        End If

      Next

      sobjTabela.EndColumn

      sobjTabela.EndRow

      TableMultiGraph

      sstrGroupFieldLastValue = "kapushinefrebow"
      blnFirst = True
      Do While Not sobjRS2.EOF And sobjBrowse.IsBrowsePage
        If sstrGroupField > "" Then
          If sstrGroupFieldLastValue <> "kapushinefrebow" Then
            If sstrGroupFieldLastValue <> sobjRS2(Session(sstrGroupField & "Field")) Then
              EditShowSummarySubTotal intNumCols

            End If
          End If

          sstrGroupFieldLastValue = sobjRS2(Session(sstrGroupField & "Field"))

        End If

        sobjTabela.BeginRow 1
        For i = 1 to Session("FieldNumber")
          If Not Session(i & "Operation") > "" Or IsNull(Session(i & "Operation")) Then
            If Session(i & "IsCount") > "" Or Session(i & "IsSum") > "" Then
              Session(i & "Value") = sobjRS2("s" & Session(i & "Field"))

            Else
              Session(i & "Value") = sobjRS2(Session(i & "Field"))

            End If
          End If
        Next

        For i = 1 to Session("FieldNumber")
          If Session(i & "Operation") > "" Then
            If Not blnFirst And Session(i & "Have1stValue") > "" Then
              Session(i & "Value") = SolveOperation(i)

            ElseIf Not Session(i & "Have1stValue") > "" Then
              Session(i & "Value") = SolveOperation(i)

            End If
            blnFirst = False

          End If
          If sstrGroupField > "" Then
            If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
              Session(i & "SubTotal") = Session(i & "SubTotal") + Session(i & "Value")

            End If

          End If

          If Session(i & "IsGraph") > "" Then
            sobjTabela.ColumnGraph Session(i & "Value"), "<a href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & "&" & conPEditFilter & "=" & EditMakeFilter(sstrFields2) & ">" & FormatVariable(Session(i & "Value"), Session(i & "Type")) & "</a>"

          Else
            If Session(i & "IsCalcField") > "" Or Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
              sobjTabela.ColumnAlign = "Right"

              sobjTabela.BeginColumn
REM              EditBeginMakeLink False
              Response.Write FormatVariable(Session(i & "Value"), Session(i & "Type"))
              EditEndMakeLink
              sobjTabela.EndColumn
              sobjTabela.ColumnAlign = ""

            Else
              Session("CurrentField") = Session(i & "ShowField")
              If TypeName(EditFieldValue(sobjRS2)) = "Long" Or TypeName(EditFieldValue(sobjRS2)) = "Integer" Or _
                TypeName(EditFieldValue(sobjRS2)) = "Currency" Or TypeName(EditFieldValue(sobjRS2)) = "Double" Or TypeName(EditFieldValue(sobjRS2)) = "Single" Then
                sobjTabela.ColumnAlign = "Right"

              End If

              sobjTabela.BeginColumn
              Select Case Session(EditCurrentField & "EditOption")
                Case conLookupField, conRadioField, conCheckField, conSelectField
                  Response.Write BreakString(EditFieldValue(sobjRS2), Session(EditCurrentField & "ListChars"))

                Case Else
                  Response.Write BreakString(PutMask(FormatVariable(EditFieldValue(sobjRS2), Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars"))

              End Select
              sobjTabela.ColumnAlign = ""
              sobjTabela.EndColumn

            End If
          End If

        Next
        sobjTabela.EndRow

        sobjRS2.MoveNext
        sobjBrowse.IncLine

      Loop

      If sstrGroupField > "" Then
      Dim strColor
        EditShowSummarySubTotal intNumCols

        strColor = sobjTabela.RowColor
        sobjTabela.RowColor = sobjTabela.HeadRowColor
        sobjTabela.BeginRow 1
        For i = 1 to Session("FieldNumber")
          If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
            If Session(i & "IsGraph") > "" Then
              sobjTabela.BeginColumn
              sobjTabela.BeginTable ""
              sobjTabela.BeginRow 1
              sobjTabela.ColumnWidth = "70"
              sobjTabela.ColumnAlign = "Right"
              sobjTabela.CompleteColumn FormatVariable(Session(i & "Total"), Session(i & "Type")) & " "
              sobjTabela.EndRow
              sobjTabela.EndTable
              sobjTabela.ColumnWidth = ""
              sobjTabela.ColumnAlign = ""

              sobjTabela.EndColumn

            Else
              sobjTabela.ColumnAlign = "Right"
              sobjTabela.CompleteColumn FormatVariable(Session(i & "Total"), Session(i & "Type"))
              sobjTabela.ColumnAlign = ""

            End If

          Else
            sobjTabela.CompleteColumn "&nbsp;"

          End If
        Next
        sobjTabela.EndRow
        sobjTabela.RowColor = strColor

      End If

      For i = 1 to Session("FieldNumber")
        ClearSession i & "Total"
        ClearSession i & "SubTotal"
        ClearSession i & "IsCalcField"
        ClearSession i & "IsSum"
        ClearSession i & "IsTotal"
        ClearSession i & "IsCount"
        ClearSession i & "Field"
        ClearSession i & "CaptionField"
        ClearSession i & "ShowField"
        ClearSession i & "OrderSQL"
        ClearSession i & "MaxValue"
        ClearSession i & "TotalValue"
        ClearSession i & "Operation"
        ClearSession i & "Type"
        ClearSession i & "Have1stValue"
        ClearSession i & "Value"

      Next
      sobjTabela.EndTable

      FormEnd

      sobjBrowse.Status = True
      sobjBrowse.EndBrowse
      sobjBrowse.Bar

      sobjShow.CenterEnd

    Else
      sobjShow.MessageError "<center>Não há detalhamento na Tabela " & Session(EditCurrentList &_
      "Caption") & " para as opções escolhidas !!!</center><P>"
      sobjShow.HTMLCR " <a href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd &_
                 "&" & EditWhatList &_
                 ">Incluir um registro</a><P>"

    End If

    sobjShow.BR
    sobjRS2.Close

    EditFPLPLink

  Else
    sobjShow.MessageError "Opção não existente - Listagem (" & sparEditList & ") !"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowEditor

REM -------------------------------------------------------------------------
REM Rotina que retorna um dado ou operacao
REM -------------------------------------------------------------------------
Public Function EditNextMathData
Dim strSubText, strChar, blnEnd

  strChar = Left(sstrText, 1)
  blnEnd = False

  Do
    Select Case strChar
      Case "-" , "+", "*", "/",""
        If strSubText = "" Then
          strSubText = strChar
          sstrText = Mid(sstrText, 2)

        End If
        blnEnd = True

      Case Else
        strSubText = strSubText & strChar
        sstrText = Mid(sstrText, 2)
        strChar = Left(sstrText, 1)

    End Select

   Loop While Not blnEnd

   EditNextMathData = strSubText
   sstrSubText = strSubText

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditNextOperation

Private Function SolveOperation(ByVal intField)
Dim dblTotal, strOperator

  dblTotal = 0
  EditBeginReadOperation Session(intField & "Operation")
  strOperator = "+"
  Do While EditIsThereOperation
    Select Case EditNextMathData
       Case "+","-","/","*"
         strOperator = EditLastOperation

       Case Else
         If IsNumeric(EditLastOperation) Then
           Select Case strOperator
             Case "+"
               dblTotal = dblTotal + CDbl(EditLastOperation)

             Case "-"
               dblTotal = dblTotal - CDbl(EditLastOperation)

             Case "*"
               dblTotal = dblTotal * CDbl(EditLastOperation)

             Case "/"
               dblTotal = dblTotal / CDbl(EditLastOperation)

           End Select

         ElseIf Len(EditLastOperation) = 1 Then
           Select Case strOperator
             Case "+"
               dblTotal = dblTotal + Session((Asc(UCase(EditLastOperation)) - 64) & "Value")

             Case "-"
               dblTotal = dblTotal - Session((Asc(UCase(EditLastOperation)) - 64) & "Value")

             Case "*"
               dblTotal = dblTotal * Session((Asc(UCase(EditLastOperation)) - 64) & "Value")

             Case "/"
               dblTotal = dblTotal / Session((Asc(UCase(EditLastOperation)) - 64) & "Value")

           End Select

        End If

    End Select

  Loop
  SolveOperation = dblTotal

End Function

Private Sub EditShowSummarySubTotal(ByVal intNumCols)
Dim strColor, i
  strColor = sobjTabela.RowColor
  sobjTabela.RowColor = "WHITE"
  sobjTabela.BeginRow 1
  For i = 1 to Session("FieldNumber")
    If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
      If Session(i & "IsGraph") > "" Then
        sobjTabela.BeginColumn
        sobjTabela.BeginTable ""
        sobjTabela.BeginRow 1
        sobjTabela.ColumnWidth = "70"
        sobjTabela.ColumnAlign = "Right"
        sobjTabela.CompleteColumn FormatVariable(Session(i & "SubTotal"), Session(i & "Type")) & "&nbsp;"
        If Session(i & "TotalValue") > 0 Then
          sobjTabela.ColumnWidth = "50"
          sobjTabela.CompleteColumn "&nbsp;" & ZTINumber(100 * Session(i & "SubTotal") / Session(i & "TotalValue")) & "%" & "&nbsp;"

        End If
        sobjTabela.ColumnWidth = ""
        sobjTabela.EndRow
        sobjTabela.EndTable
        sobjTabela.ColumnAlign = ""

        sobjTabela.EndColumn
        Session(i & "Total") = Session(i & "Total") + Session(i & "SubTotal")
        Session(i & "SubTotal") = 0

      Else
        sobjTabela.ColumnAlign = "Right"
        sobjTabela.CompleteColumn FormatVariable(Session(i & "SubTotal"), Session(i & "Type"))
        Session(i & "Total") = Session(i & "Total") + Session(i & "SubTotal")
        Session(i & "SubTotal") = 0
        sobjTabela.ColumnAlign = ""

      End If
REM     ElseIf sstrGroupField = Session(i & "ShowField") Then
REM       sobjTabela.CompleteColumn Session(i & "SubTotal")
    Else
      sobjTabela.CompleteColumn "&nbsp;"

    End If

  Next
  sobjTabela.EndRow

  sobjTabela.RowColor = strColor
  sobjTabela.ColumnColSpan = intNumCols
  sobjTabela.BeginRow 1
  sobjTabela.CompleteColumn "&nbsp;"
  sobjTabela.EndRow
  sobjTabela.ColumnColSpan = ""

End Sub

Private Sub TableMultiGraph
Dim i
  For i = 1 to sintGraphNum
    Session(i & "sintCor") = Session("1sintCor")
    Session(i & "sintR")   = Session("1sintR")
    Session(i & "sintG")   = Session("1sintG")
    Session(i & "sintB")   = Session("1sintB")

  Next

End Sub

REM -------------------------------------------------------------------------
REM Desenha a Legenda do Grafico de Barras
REM -------------------------------------------------------------------------
Private Sub TableSummaryLegend(intMax, intTotal, strText, strType, strOrder)
Dim strCor
Dim i

REM  sobjShow.MessageError strType
  sintGraphNum = sintGraphNum + 1
  sintGraphActual = sintGraphNum

  If Session("1sintCor") = "" Then
     Session("1sintCor") = -1
     Session("1sintR")   = 255
     Session("1sintG")   = -85
     Session("1sintB")   = 0

  End If
  strCor = sobjTabela.RowColor
  sobjTabela.RowColor = sobjTabela.HeadRowColor

  sarrTotal(sintGraphActual) = intTotal
  sarrMax(sintGraphActual) = intMax
  sobjTabela.BeginColumn

  If sarrTotal(sintGraphActual) > 0 Then
    sobjTabela.BeginTable 300
    sobjTabela.ColumnWidth = 120

  Else
    sobjTabela.BeginTable 250
    sobjTabela.ColumnWidth = 70

  End If

  sobjTabela.BeginRow ""

  sobjTabela.ColumnAlign = "Left"
  sobjTabela.BeginRow 1
  sobjTabela.BeginColumn
  If LCase(CStr(Session("OrderField"))) = LCase(CStr(strOrder)) Then
    If Session("OrderDesc") = "1" Then
      ShowImage "/ZTITools/OD.gif", "", "Ordenado descendentemente pelo campo " & EditCaption(False), sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter, "", 8, 7
      sobjShow.Nbsp
      Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter & """>" & strText & "</A>"

    Else
      ShowImage "/ZTITools/OA.gif", "", "Ordenado ascendentemente pelo campo " & EditCaption(False), sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & conPEditOrderDescStr & GetFilter, "", 8, 7
      sobjShow.Nbsp
      Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & conPEditOrderDescStr & GetFilter & """>" & strText & "</A>"

    End If

  Else
    Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter & """>" & strText & "</A>"

  End If

  sobjTabela.ColumnAlign = "Left"
  
REM  sobjShow.MessageError intMax & "&" & intTotal & "&" & strText & "&" & strType & "&" & strOrder
  
  If intMax > 2 Then
    sobjTabela.ColumnWidth = 60
    sobjTabela.ColumnAlign = "Left"
    sobjTabela.CompleteColumn "0"
    sobjTabela.ColumnAlign = "Center"
    If intMax < 10 Then
      sobjTabela.CompleteColumn FormatVariable(Round(intMax / 2, 1), strType)

    Else
      sobjTabela.CompleteColumn FormatVariable(intMax \ 2, strType)

    End If
    sobjTabela.ColumnAlign = "Right"
    sobjTabela.CompleteColumn FormatVariable(intMax, strType)
    sobjTabela.EndRow

    sobjTabela.BeginRow 1

    sobjTabela.ColumnWidth = 60
    If intTotal > 0 Then
       sobjTabela.CompleteColumn "&nbsp;<Font Size=-1><B>" & FormatVariable(sarrTotal(sintGraphActual), strType) & "<B></Font>&nbsp;"

    Else
       sobjTabela.CompleteColumn "&nbsp;"

    End If

    sobjTabela.ColumnWidth = 60
    sobjTabela.ColumnAlign = "Left"
    sobjTabela.CompleteColumn "|"
    sobjTabela.ColumnAlign = "Center"
    sobjTabela.CompleteColumn "|"
    sobjTabela.ColumnAlign = "Right"
    sobjTabela.CompleteColumn "|"

  Else
    sobjTabela.ColumnWidth = 180 \ (intMax + 1)
    sobjTabela.ColumnAlign = "Left"
    For i = 0 to intMax - 1
      sobjTabela.CompleteColumn FormatVariable(i, strType)
      sobjTabela.ColumnAlign = "Center"

    Next

    sobjTabela.ColumnAlign = "Right"
    sobjTabela.CompleteColumn FormatVariable(intMax, strType)
    sobjTabela.EndRow

    sobjTabela.BeginRow 1

    sobjTabela.ColumnWidth = 60
    If intTotal > 0 Then
      sobjTabela.CompleteColumn "&nbsp;<B>" & FormatVariable(sarrTotal(sintGraphActual), strType) & "<B>&nbsp;"

    Else
       sobjTabela.CompleteColumn "&nbsp;"

    End If

    sobjTabela.ColumnAlign = "Left"
    sobjTabela.ColumnWidth = 120 \ (intMax + 1)
    For i = 0 to intMax - 1
      sobjTabela.CompleteColumn "|"
      sobjTabela.ColumnAlign = "Center"

    Next

    sobjTabela.ColumnAlign = "Right"
    sobjTabela.CompleteColumn "|"

  End If
  sobjTabela.EndRow

  sobjTabela.EndTable

  sobjTabela.EndColumn

  sobjTabela.ColumnAlign = ""
  sobjTabela.ColumnWidth = ""
  sobjTabela.RowColor = strCor

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub TableGraphLegend

%>