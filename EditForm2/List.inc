<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/List.inc
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
REM Rotina que monta a listagem
REM -------------------------------------------------------------------------
Public Sub EditShowAZ
Dim sql, i, j, strAux, strAux2, strTip, strTip2, arrFields(30), intCount, strOrder, strJoin
REM  DebugMode

  Session("LastToFind") = conOptionAZEdit

  sobjBrowse.Actual = EditCurrentTable

  If Not Session("OrderField") > "" Then
    EditBeginReadOperation Session(EditCurrentTable & "FieldList")

    Do While EditIsThereOperation
      If EditFindField(EditCurrentTable, EditNextOperation) Then
        sstrText = ""
        Session("OrderField") = Session("CurrentField")

      End If

    Loop

  End If

  EditBeginReadOperation Session("OrderField")
  strOrder = " ORDER BY"

  blnFirst = True

  strJoin = ""

  Do While EditIsThereOperation
    Select Case EditNextOperation
       Case ",", ";", "-"

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
             strJoin  = strJoin & " LEFT OUTER JOIN " & Session(EditCurrentField & "LookupTable") &_
             " on " & Session(EditCurrentField & "LookupTable") & "." & Session(EditCurrentField & "LookupField") & " = " & _
             EditTableOf(EditCurrentTable, Session("CurrentTable")) & "." & Session(EditCurrentField & "Field")


           Case Else
             strOrder = strOrder & Session(EditCurrentField & "Field")

        End Select

        blnFirst = False
        If Session("OrderDesc") = "1" Then
          strOrder = strOrder & " DESC "

        End If

    End Select

  Loop

  sobjTabela.ColumnNoWrap = False

  Dim blnFirst
  blnFirst = True
  sql = "SELECT "

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Session(EditCurrentTable & i & "IsKey") Then
      If Not blnFirst Then
        sql = sql & ", "

      Else
        blnFirst = False

      End If

      sql = sql & Session(EditCurrentTable & i & "Field")

    End If
  Next

  sql = sql & " FROM " & EditCurrentTable & strJoin & Filter

  sobjCMD.CommandText = sql & strOrder
  sobjCMD.CommandType = adCmdText
  DebugMessage sql & strOrder, 1

  sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

  j = 0
  If Not sobjRS.EOF And Not sobjRS.BOF Then
    sobjBrowse.BeginBrowse sobjRS

    Do While Not sobjRS.EOF And sobjBrowse.IsBrowsePage
      j = j + 1
      For i = 1 to Session(EditCurrentTable & "FieldNumber")
     		DebugMessage Session(EditCurrentTable & i & "Field"), 2
        If Session(EditCurrentTable & i & "IsKey") Then
          Session("Listagem" & j & "," & i & "Value") = sobjRS(Session(EditCurrentTable & i & "Field"))

        End If

      Next
      sobjRS.MoveNext
      sobjBrowse.IncLine

    Loop

    sobjBrowse.Status = True
    sobjBrowse.EndBrowse
	
	  sql = "SELECT * FROM " & EditCurrentTable & Session(EditCurrentTable & "JoinFormUnit")
	
	  For i = 1 to Session(EditCurrentTable & "FieldNumber")
	    If Session(EditCurrentTable & i & "EditOption") = conLookupField Then
	       sql = sql & " LEFT OUTER JOIN " & Session(EditCurrentTable & i & "LookupTable") &_
	       " on " & Session(EditCurrentTable & i & "LookupTable") & "." & Session(EditCurrentTable & i & "LookupField") & " = " & _
	       EditTableOf(EditCurrentTable, i) & "." & Session(EditCurrentTable & i & "Field")
	
	    End If
	
	  Next
	
    SetPutOr False
	
	  Dim k
	  For k = 1 to j
	    For i = 1 to Session(EditCurrentTable & "FieldNumber")
	      If Session(EditCurrentTable & i & "IsKey") Then
	        sql = sql & PutOr & Session(EditCurrentTable & i & "Field") & "=" & FormatForSql(Session("Listagem" & k & "," & i & "Value"),Session(EditCurrentTable & i & "Type"))
	
	      End If
	
	    Next
	  Next

	  sobjShow.Nbsp

	  sobjCMD.CommandText = sql & strOrder
	  sobjCMD.CommandType = adCmdText
	  DebugMessage sql & strOrder, 1

	  sobjRS2.Open sobjCMD, , adOpenKeySet, adLockReadOnly
	
	  If Not sobjRS2.EOF And Not sobjRS2.BOF Then
	    Dim strMessage
	
	    strMessage = "Lista de " & Session(EditCurrentTable & "Caption")
	
	    sobjShow.Center
	
	    EditBeginReadOperation Session(EditCurrentTable & "FieldList")
	
	    Dim Width, strList
	
	    Width = 0
	    Do While EditIsThereOperation
	      Select Case EditNextOperation
	         Case ",", ";", "-"
	         Case "="
	           EditIsThereOperation
	
	         Case Else
	           If EditFindField(EditCurrentTable, EditLastOperation) Then
	             If Session(EditCurrentField & "Type") <> "varchar" Then
	               Width = Width + EditSize + 8
	
	             Else
	               Width = Width + EditSize
	
	             End If
	             If strList > "" Then
	               strList = strList & ","
	
	             End If
	             strList = strList & Session("CurrentField")
	
	           End If
	
	       End Select
	
	    Loop
	
	    If Width < 60 Then
	      Width = 60
	
	    ElseIf Width > 95 Then
	      Width = 95
	
	    End If               

	    sobjTabela.Padding = 3
	    sobjTabela.Spacing = 1
	    sobjTabela.BeginTable Width & "%", strMessage, Session(EditCurrentTable & "ListOperationNumber") + 1, True
	    sobjTabela.Padding = 2
	    sobjTabela.Spacing = 0
	
	    ClassAss
	
	    If Session(EditCurrentTable & "OrderNumber") > 0 Then
	      For i = 1 to Session(EditCurrentTable & "OrderNumber")
	        strAux = ""
	        strAux2 = ""
	        EditBeginReadOperation Session(EditCurrentTable & i & "Order")
	        Do While EditIsThereOperation
	          Select Case EditNextOperation
	            Case ",", ";", "-", "="
	              strAux = strAux & EditLastOperation & " "
	
	            Case Else
	              If EditFindField(EditCurrentTable, EditLastOperation) Then
	                strAux = strAux & EditCaption(False)
	                If strAux2 > "" Then
	                  strAux2 = strAux2 & ","
	
	                End If
	                strAux2 = strAux2 & Session("CurrentField")
	
	              End If
	
	          End Select
	
	        Loop
	        If Session("OrderField") = strAux2 Then
	          strAux2 = Server.URLEncode(strAux2)
	          If Session("OrderDesc") = "1" Then
	            strTip = strTip & "<IMG SRC=""/ZTITools/OD.gif"" BORDER=0>" & _
	            "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
	            EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & GetFilter & """>" & _
	            " " & strAux & "</A><BR>"
	
	          Else
	            strTip = strTip & "<IMG SRC=""/ZTITools/OA.gif"" BORDER=0>" & _
	            "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
	            EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & conPEditOrderDescStr & GetFilter & """>" & _
	             " " & strAux & "</A><BR>"
	
	          End If
	        Else
	          strAux2 = Server.URLEncode(strAux2)
	          strTip = strTip & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
	          conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & GetFilter & """>" & strAux & "</A><BR>"
	
	        End If
	
	      Next
	
	    End If
	
	    For i = 1 to Session(EditCurrentTable & "ListNumber")
	      strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
	      conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditList & "=" & Session(EditCurrentTable & i & "ListName") & GetFilter & """>" & Session(EditCurrentTable & i & "ListCaption") & "</A><BR>"
	
	    Next
	
	    sobjTabela.ColumnColSpan = Session(EditCurrentTable & "ListOperationNumber") + 1
	
	    sobjTabela.BeginRow 1
	    sobjTabela.BeginColumn
	
	    sobjTabela.BeginTable "100%"
	    sobjTabela.ColumnAlign = "Left"
	    sobjTabela.ColumnVAlign = "Top"
	    sobjTabela.BeginRow 1
	    sobjTabela.BeginColumn
	    Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd & _
	               EditWhatList & conPEditEditableStr & ">Incluir</a></B>"
	    sobjShow.Nbsp
	    sobjShow.Nbsp
	    Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
	               EditWhatList & conPEditEditableStr & GetFilter & ">Editar</a></B>"
	
	    If Session(EditCurrentTable & "FieldFind") > "" Then
	      sobjShow.Nbsp
	      sobjShow.Nbsp
	      Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindAdvanced & _
	               EditWhatList & ">Localizar</a></B>"
	    Else
	      sobjShow.Nbsp
	      sobjShow.Nbsp
	      Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindCommon & _
	                 EditWhatList & ">Localizar</a></B>"
	
	    End If
	    sobjTabela.EndColumn
	                                                    
	    If Session(EditCurrentTable & "ListNumber") > 0 Then
	      sobjTabela.ColumnAlign = "Right"
	      sobjTabela.ColumnVAlign = "Top"
	      sobjTabela.BeginColumn
	      Response.Write "<FONT SIZE=2><DIV ID=F2 class=SS onMouseOver=""makeCool(F2);"" onMouseOut=""makeNormal(F2);"" onClick=""OpenDesc(T2);"" STYLE=""cursor:hand""><B>Outras listagens</B></DIV>"
	      Response.Write "<table><tr><td align=justify>"
	      Response.Write "<DIV ID=T2 class=dbody style=""display:'none';""><B>"
	      Response.Write strTip2
	      Response.Write "</B></DIV></td></tr></table>"
	      sobjTabela.EndColumn
	
	    End If
	
	    If Session(EditCurrentTable & "OrderNumber") > 0 Then
	      sobjTabela.ColumnAlign = "Right"
	      sobjTabela.ColumnVAlign = "Top"
	      sobjTabela.BeginColumn
	      Response.Write "<FONT SIZE=2><DIV ID=F1 class=SS onMouseOver=""makeCool(F1);"" onMouseOut=""makeNormal(F1);"" onClick=""OpenDesc(T1);"" STYLE=""cursor:hand""><B>Outras ordenações</B></DIV>"
	      Response.Write "<table><tr><td align=justify>"
	      Response.Write "<DIV ID=T1 class=dbody style=""display:'none';""><B>"
	      Response.Write strTip
	      Response.Write "</B></DIV></td></tr></table>"
	      sobjTabela.EndColumn
	
	    End If
	    sobjTabela.EndRow
	    sobjTabela.EndTable
	    sobjTabela.EndColumn
	    sobjTabela.EndRow
	    sobjTabela.ColumnAlign = ""
	    sobjTabela.ColumnColSpan = ""
	    sobjTabela.BeginRow 2, True
	
	    EditBeginReadOperation strList
	
	Dim intNumHeader, intNumArrHeaders, intAux
	
	    intNumHeader     = 0
	    intCount         = 1
	    intNumArrHeaders = 0
	
	    If Session(EditCurrentTable & "HeaderNumber") > 0 Then
	      sobjTabela.ColumnRowSpan = 2
	      intAux                   = 0
	
	    End If
	
	    Do While EditIsThereOperation
	      Select Case EditNextOperation
	        Case ",", ";", "-", "="
	
	        Case Else
	          Session("CurrentField") = EditLastOperation
	          intNumHeader = intNumHeader + 1
	REM          DebugMessage Session(EditCurrentTable & "HeaderNumber"), 2
	          If Session(EditCurrentTable & "HeaderNumber") > 0 Then
	            If intNumHeader = Session(EditCurrentTable & intCount & "HeaderBegin") Then
	              sobjTabela.ColumnRowSpan = ""
	              sobjTabela.ColumnColSpan = Session(EditCurrentTable & intCount & "HeaderWidth")
	              sobjTabela.HeadColumn "<SPAN CLASS=SS>" & Session(EditCurrentTable & intCount & "HeaderText") & "</SPAN>"
	              sobjTabela.ColumnRowSpan = 2
	              sobjTabela.ColumnColSpan = ""
	              intAux           = 1
	              intNumArrHeaders = intNumArrHeaders + 1
	              arrFields(intNumArrHeaders) = EditLastOperation

	            Else
	              If intAux > 0 And intAux < Session(EditCurrentTable & intCount & "HeaderWidth") Then
	                intNumArrHeaders = intNumArrHeaders + 1
	                arrFields(intNumArrHeaders) = EditLastOperation
	                intAux = intAux + 1

  	              If intAux >= Session(EditCurrentTable & intCount & "HeaderWidth") Then
    	              If Session(EditCurrentTable & "HeaderNumber") > intCount Then
  	                  intCount = intCount + 1

  	                End If
	                End If

	              Else
	                intAux = 0
	                EditTableHeader EditCaption(False), Session("CurrentField"), Session(EditCurrentField & "ListSize")
	
	               End If
	            End If
	
	          Else
	            EditTableHeader EditCaption(False), Session("CurrentField"), Session(EditCurrentField & "ListSize")
	
	          End If
	
	       End Select
	
	    Loop
	
	    If intNumArrHeaders > 0 Then
	      sobjTabela.ColumnRowSpan = ""
	      sobjTabela.EndRow
	      sobjTabela.BeginRow 2, True
	      For i = 1 to intNumArrHeaders
	        Session("CurrentField") = arrFields(i)
	        EditTableHeader EditCaption(False), Session("CurrentField"), Session(EditCurrentField & "ListSize")

	      Next
	
	    End If
	
	    sobjTabela.EndRow
	
	    Do While Not sobjRS2.EOF
	      EditBeginReadOperation strList
	
	      sobjTabela.BeginRow 1
	
	REM      intCount = 0
	
	      Do While EditIsThereOperation
	        Select Case EditNextOperation
	          Case ",",";","-","="
	          Case Else
	REM            intCount = intCount + 1
	REM            If arrFields(intCount) <> Session(EditCurrentField & "Value") Then
	            Session("CurrentField") = EditLastOperation
	            If TypeName(EditFieldValue(sobjRS2)) = "Long" Or TypeName(EditFieldValue(sobjRS2)) = "Integer" Or _
	              TypeName(EditFieldValue(sobjRS2)) = "Currency" Or TypeName(EditFieldValue(sobjRS2)) = "Double" Or TypeName(EditFieldValue(sobjRS2)) = "Single" Then
	              sobjTabela.ColumnAlign = "Right"
	
	            End If
	
	            arrFields(intCount) = Session(EditCurrentField & "Value")
	
	            sobjTabela.ColumnWidth = Session(EditCurrentField & "ListSize")
	            sobjTabela.BeginColumn
	            EditBeginMakeLink False
	            Select Case Session(EditCurrentField & "EditOption")
	              Case conLookupField, conRadioField, conCheckField, conSelectField
	                Response.Write BreakString(EditFieldValue(sobjRS2), Session(EditCurrentField & "ListChars"))
	
	              Case Else
	                Response.Write BreakString(PutMask(FormatVariable(EditFieldValue(sobjRS2), Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars"))
	
	            End Select
	            sobjTabela.ColumnAlign = ""
	            EditEndMakeLink
	
	            sobjTabela.EndColumn
	
	        End Select
	
	      Loop
	      sobjTabela.EndRow
	
	      sobjRS2.MoveNext
	    Loop
	
	    sobjTabela.EndTable
	
	    sobjBrowse.Bar
	
	    sobjShow.CenterEnd
	  End If
    sobjRS2.Close

  Else
    sobjShow.MessageError "<center>Não há detalhamento na Tabela " & Session(EditCurrentTable &_
    "Caption") & " para as opções escolhidas !!!<P>"
    sobjShow.HTMLCR " <a href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd &_
               EditWhatList & conPEditEditableStr & _
               ">Incluir um registro</a><P></center>"

  End If

  sobjRS.Close

  sobjShow.BR
  EditFPLPLink

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowEditor

REM -------------------------------------------------------------------------
REM Monta um header de listagem
REM -------------------------------------------------------------------------
Private Sub EditTableHeader(strField, strOrderField, intSize)
  sobjTabela.ColumnVAlign = "Middle"
  sobjTabela.ColumnWidth = intSize
  sobjTabela.ColumnColor = sobjTabela.HeadRowColor
  sobjTabela.BeginColumn
  Response.Write "<CENTER><B>"
  If LCase(CStr(Session("OrderField"))) = LCase(CStr(strOrderField)) Then
    If Session("OrderDesc") = "1" Then
      ShowImage "/ZTITools/OD.gif", "", "Ordenado descendentemente pelo campo " & strField, sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & GetEditable & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetList & GetFilter, "", 8, 7
      Response.Write " <A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & GetList & GetFilter & """>" & strField & "</A>"

    Else
      ShowImage "/ZTITools/OA.gif", "", "Ordenado ascendentemente pelo campo " & strField, sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & GetEditable & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & conPEditOrderDescStr & GetList & GetFilter, "", 8, 7
      Response.Write " <A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & conPEditOrderDescStr & GetList & GetFilter & """>" & strField & "</A>"

    End If
  Else
    Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & GetList & GetFilter & """>" & strField & "</A>"

  End If
  sobjTabela.EndColumn
  sobjTabela.ColumnColor = ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Function EditTableHeader

Private Function EditListStr
 EditListStr = "&" & conPEditList & "=" & sparEditList

End Function

%>