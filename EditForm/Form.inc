<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Form.inc
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
REM Rotina altera o ColSpan da Tabela conforme
REM -------------------------------------------------------------------------
Public Sub EditColumnColSpan
Dim intColSpan

  If sintNumCol = 1 And EditIsThereNumber Then
    intColSpan = Session(EditCurrentTable & "UnitOperationNumber") \ EditNextNumber

  Else
    intColSpan = Session(EditCurrentTable & "UnitOperationNumber") \ sintLastNumber

  End If

  If sintNumCol <= (Session(EditCurrentTable & "UnitOperationNumber") mod sintLastNumber) Then
    intColSpan = intColSpan + 1

  End If

  If intColSpan > 1 Then
    sobjTabela.ColumnColSpan = intColSpan

  Else
    sobjTabela.ColumnColSpan = ""

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditColumnColSpan

REM -------------------------------------------------------------------------
REM Rotina cria os javascripts das mensagens de exclusao e salvamento
REM -------------------------------------------------------------------------
Public Sub JavaScriptDefaults
  sobjShow.HTMLCR "function ButtonsTest(valor){"
  sobjShow.HTMLCR "  switch (valor){"
  If Session("RecordAction") > "" Then
    sobjShow.HTMLCR "  case """ & constrCancel & """ :"
    sobjShow.HTMLCR "    if (blnChanged)"
    sobjShow.HTMLCR "      return confirm(""Você tem certeza que deseja cancelar a inclusão do registro ? \nOs dados digitados serão perdidos."");"
    sobjShow.HTMLCR "    else"
    sobjShow.HTMLCR "      return true;"
    sobjShow.HTMLCR "  Break;"

  Else
    sobjShow.HTMLCR "  case """ & constrDelete & """ :"
    sobjShow.HTMLCR "  return confirm(""Você tem certeza que deseja excluir este registro ?""); Break;"

  End If

  If Session(EditCurrentTable & "ClientValidation") And Session("RecordAction") > "" Then
    sobjShow.HTMLCR "  case """ & constrSave & """ :"
    sobjShow.HTMLCR "  case """ & constrSaveAdd & """ :"
    sobjShow.HTMLCR "  case """ & constrSaveCopy & """ :"
    sobjShow.HTMLCR "  return ValidateForm(); Break;"

  End If

  sobjShow.HTMLCR "  default : return true;"
  sobjShow.HTMLCR "  }}"

  If Session("RecordAction") > "" Then
    sobjShow.HTMLCR "var blnChanged = false;"
  Dim i, j, k, blnFirst
    blnFirst = true
    EditBeginReadOperation Session("LastTab")

    Dim strList

    Do While EditIsThereOperation
      Select Case EditNextOperation
        Case ",", ";", "-"
        Case "="
          EditIsThereOperation

        Case Else
          If EditFindField(EditCurrentTable, EditLastOperation) then
            If Session(EditCurrentField & "DisableNumber") > 0 Then
              If blnFirst Then
                sobjShow.HTMLCR "function DisableFields(strField)"
                sobjShow.HTMLCR "{"

              Else
                sobjShow.HTMLCR "  else{"

              End If

              sobjShow.HTMLCR "  if (strField == '" & Session(EditCurrentField & "Field") & "'){"
              For j = 1 to Session(EditCurrentField & "DisableNumber")
                Select Case Session(EditCurrentField & "EditOption")
                  Case conCheckField
                    If CStr(Session(EditCurrentField & ",1Value")) = CStr(Session(EditCurrentField & "," & j & ",1DisableValue")) Xor Session(EditCurrentField & "," & j & "DisableCondition") Then
                      sobjShow.HTMLCR "    if (document.frmEditForm." & Session(EditCurrentField & "Field") & ".checked) "

                    Else
                      sobjShow.HTMLCR "    if (! document.frmEditForm." & Session(EditCurrentField & "Field") & ".checked) "

                    End If

                  Case Else
                    Response.Write "    if ("
                    For k = 1 to Session(EditCurrentField & "," & j & "DisableValueNumber")
                      If k > 1 Then
                        If Session(EditCurrentField & "," & j & "DisableCondition") Then
                          Response.Write " || "

                        Else
                          Response.Write " && "

                        End If

                      End If
                      If Session(EditCurrentField & "EditOption") = conRadioField Then
                        If Session(EditCurrentField & "," & j & "DisableCondition") Then
                          Response.Write "(document.frmEditForm." & Session(EditCurrentField & "Field") & ".item(" & Session(EditCurrentField & "," & j & "," & k & "DisableValue") & ").checked)"

                        Else
                          Response.Write "(!document.frmEditForm." & Session(EditCurrentField & "Field") & ".item(" & Session(EditCurrentField & "," & j & "," & k & "DisableValue") & ").checked)"

                        End If

                      Else
                        Response.Write "(document.frmEditForm." & Session(EditCurrentField & "Field") & ".value "
                        If Session(EditCurrentField & "," & j & "DisableCondition") Then
                          Response.Write "="

                        Else
                          Response.Write "!"

                        End If
                        Response.Write "= """ & Session(EditCurrentField & "," & j & "," & k & "DisableValue") & """)"

                      End If

                    Next
                    sobjShow.HTMLCR ")"


                End Select
                sobjShow.HTMLCR "      document.frmEditForm." & Session(EditCurrentField & "," & j & "DisableField") & ".disabled = true;"
                sobjShow.HTMLCR "    else"
                sobjShow.HTMLCR "      document.frmEditForm." & Session(EditCurrentField & "," & j & "DisableField") & ".disabled = false;"

              Next

              If blnFirst Then
                blnFirst = false

              Else
                sobjShow.HTMLCR "  }"

              End If
              sobjShow.HTMLCR "  }"
              sobjShow.HTMLCR "}"

            End If

          End If

      End Select

    Loop


  End If

  sobjShow.HTMLCR "function FieldChanged(strField){"
  sobjShow.HTMLCR "blnChanged = true;"
  If Not blnFirst Then
    sobjShow.HTMLCR "DisableFields(strField);"

  End If
  sobjShow.HTMLCR "}"

  sobjShow.HTMLCR "function TabCtrlsClick(intTab){"
  sobjShow.HTMLCR "  document.frmEditForm.hidEditor.value = intTab;"
  sobjShow.HTMLCR "  document.frmEditForm.submit();"
REM  sobjShow.HTMLCR "  alert(document.frmEditForm.hidEditor.value);"
  sobjShow.HTMLCR "}"

  sobjShow.HTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DeleteSaveMessages

REM -------------------------------------------------------------------------
REM Zera o contador de colunas
REM -------------------------------------------------------------------------
Sub StartRow
	sintNumCol = 1

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub StartRow

REM -------------------------------------------------------------------------
REM adiciona 1 no contador de colunas
REM -------------------------------------------------------------------------
Sub IncNumCol
	sintNumCol = sintNumCol + 1

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub IncNumCol

REM -------------------------------------------------------------------------
REM Rotina que mostra a tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditShowEditor
Dim blnNotStarting, FontSize, strURL, i, blnFirstEditable, strLastOperation, blnEnable

  sobjShow.BR
  sobjShow.Center
  If Session(EditCurrentTable & sparEditTab & "Task") > "" Then
    Execute Session(EditCurrentTable & sparEditTab & "Task")

  End If

  sobjShow.HTMLCR "<STYLE> .BTM{font: 8pt Arial}</STYLE>"

  Session("LastToFind") = conOptionFPEdit
  blnFirstEditable = True

  sobjTabela.Spacing = "1"
  sobjTabela.Padding = "3"
  If Session("Message") > "" Then
    sobjTabela.BeginTable "65%", "", 1
    sobjTabela.BeginRow 2
    sobjTabela.BeginColumn
    Response.Write Session("Message")
    sobjTabela.EndColumn
    sobjTabela.EndRow
    sobjTabela.EndTable

    Session("Message") = ""

  End If

  strURL = sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionExeEdition & EditWhatWhereHow & GetTab

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Request.QueryString("F" & i) > "" Then
      strURL = strURL & "&" & "F" & i & "=" & Request.QueryString("F" & i)

    End If

  Next

  sobjForm.BeginForm strURL, "frmEditForm"

  If Session(EditCurrentTable & "EditOption") <> 0 Then
    FontSize = 1

  Else
    FontSize = 2

  End If

  Dim strMessage

  sobjForm.Hidden "hidEditor"

  sobjForm.Enable = True

  If Not Session("RecordAction") > "" Then
    strMessage = Session(EditCurrentTable & "Caption")
    sobjForm.Enable = False
    blnEnable = False

  ElseIf Session("RecordAction") = "1" Then
    strMessage = Session(EditCurrentTable & "Caption") & " - Inclusão"
    blnEnable = True

  ElseIf Session("RecordAction") = "2" Then
    strMessage = Session(EditCurrentTable & "Caption") & " - Alteração"
    blnEnable = True

  End If

  sobjTabela.BeginTable "65%", strMessage, 1, True
  sobjTabela.Spacing = "0"
  sobjTabela.Padding = "0"

  blnNotStarting = False

  sobjTabela.BeginRow 0, True
  sobjTabela.ColumnAlign = "LEFT"
  sobjTabela.BeginColumn
  Response.Write "<FONT SIZE=1>"
  If sparEditFilter > "" Then
    Response.Write "Tabela Filtrada<BR>&nbsp;&nbsp;"

  End If

  Response.Write "Nº de registros : " & Session("RecordCount")

  If sparEditFilter > "" Then
    Response.Write "<BR>"

  End If
  Response.Write "</FONT>"
  sobjTabela.EndColumn
  sobjTabela.ColumnAlign = ""
  sobjTabela.EndRow

Dim intPos
  Session("LastTab") = Session(EditCurrentTable & "FieldUnit")

  For i = 1 to sparEditTab
    intPos = InStr(Session("LastTab"), "|")
    If intPos > 0 Then
      Session("LastTab") = Right(Session("LastTab"), Len(Session("LastTab")) - intPos)

    End If

  Next

  intPos = InStr(Session("LastTab"), "|")
  If intPos > 0 Then
    Session("LastTab") = Left(Session("LastTab"), intPos - 1)

  End If

  EditTotalOperation Session("LastTab"), "UnitOperationNumber", True, "UnitOperations"

  If Session(EditCurrentTable & "ClientValidation") And Session("RecordAction") > "" Then
      ClientValidates

  Else
    sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"

  End If

  JavaScriptDefaults

  sobjForm.Class = "BTM"

  sobjTabela.BeginRow 1
  sobjTabela.BeginColumn

  If EditHaveTabs Then
Dim LastColor
    ClassAss

    EditBeginReadOperation Session(EditCurrentTable & "UnitTabs")

    sobjTabela.BeginTable "100%"
    sobjTabela.BeginRow 1

    LastColor = sobjTabela.RowColor

    i = 0
    Do While EditIsThereOperation
      Select Case EditNextOperation
        Case "|"

        Case Else
          sobjTabela.BeginColumn
          If CStr(i) = CStr(sparEditTab) Then
            sobjTabela.RowColor = LastColor

          Else
            sobjTabela.RowColor = sobjTabela.HeadRowColor

          End If

          sobjTabela.Padding = "1"
          sobjTabela.BeginTable "100%"
          sobjTabela.BeginRow 1
          sobjTabela.BeginColumn
          sobjShow.Center
          If CStr(i) = CStr(sparEditTab) Then
            Response.Write "<DIV Class=SS><B>" & EditLastOperation & "</B><DIV>"

          Else
            Response.Write "<DIV Class=SS ID=Editor" & i & " onMouseOver=""makeCool(Editor" & i & ");"" onMouseOut=""makeNormal(Editor" & i & ");"" onClick=""TabCtrlsClick(" & i & ");"" STYLE=""cursor:hand""><B>" & EditLastOperation & "</B></DIV>"

          End If

REM          <a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
REM                           GetFilter & EditWhatWhereHow & "&" & conPEditTab & "=" & i & ">" & EditLastOperation & "</a></B>"

          i = i + 1
          sobjTabela.EndColumn
          sobjTabela.EndRow

          sobjTabela.EndTable

          sobjTabela.EndColumn

      End Select

    Loop

    sobjTabela.EndRow
    sobjTabela.EndTable

    sobjTabela.RowColor = LastColor

  End If

  EditBeginReadNumbers Session(EditCurrentTable & "UnitOperations")

  If Left(sstrText, 1) <> "=" Then
    StartRow
    sobjTabela.Padding =  2
    sobjTabela.BeginTable "100%"
    blnNotStarting = True
    EditColumnColSpan
    sobjTabela.BeginRow FontSize
    sobjTabela.BeginColumn

  End If

  EditBeginReadOperation Session("LastTab")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ","
        IncNumCol
        EditColumnColSpan
        sobjTabela.EndColumn
        sobjTabela.BeginColumn

      Case ";"
        StartRow
        sobjTabela.EndColumn
        sobjTabela.EndRow
        EditColumnColSpan
        sobjTabela.BeginRow FontSize
        sobjTabela.BeginColumn

      Case "-"
        StartRow
        sobjTabela.EndColumn
        sobjTabela.EndRow
        sobjTabela.EndTable
        sobjTabela.EndColumn
        sobjTabela.EndRow
        sobjTabela.ColumnColSpan = ""
        sobjTabela.BeginRow 1
        sobjTabela.BeginColumn
        sobjTabela.BeginTable "100%"
        EditColumnColSpan
        sobjTabela.BeginRow FontSize
        sobjTabela.BeginColumn

      Case "="
        StartRow
        If blnNotStarting Then
          sobjTabela.EndColumn
          sobjTabela.EndRow
          sobjTabela.EndTable
          sobjTabela.EndColumn
          sobjTabela.EndRow

        End If

        blnNotStarting = True

        sobjTabela.BeginRow 2, True
        sobjTabela.ColumnAlign = "CENTER"
        sobjTabela.Column EditNextOperation
        sobjTabela.ColumnAlign = ""
        sobjTabela.EndRow
        sobjTabela.ColumnColSpan = ""
        sobjTabela.BeginRow 1
        sobjTabela.BeginColumn
        sobjTabela.BeginTable "100%"
        EditColumnColSpan
        sobjTabela.BeginRow FontSize
        sobjTabela.BeginColumn
        EditNextOperation

      Case Else
        If EditFindField(EditCurrentTable, EditLastOperation) Then
           If (Session(EditCurrentTable & "EditOption") > 0) And (sparEditEditable = 1) Then
              If sobjForm.Enable And Session("RecordAction") > "" Then
                 sobjForm.onChange = "FieldChanged('" & Session(EditCurrentField & "Field") & "')"

              End If
              blnFirstEditable = EditShowField(blnFirstEditable And Session("RecordAction") > "", False, blnEnable, False)

           Else
             EditShowFieldAsText

           End If

        Else
          If EditLastOperation = " " Then
            Response.Write "&nbsp;"

          ElseIf EditFindGrid(EditLastOperation) Then
            sobjShow.BR
            sobjShow.HTMLCR "<IFRAME FRAMEBORDER=NO HEIGHT=100 WIDTH=450 SRC=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionGrid & "&" & conPEditWhat & "=" & Session(EditCurrentGrid & "WhatList") & """></IFRAME>"

          Else
            Response.Write EditLastOperation

          End If

        End If

    End Select
  Loop

REM  If EditFindField("ageCodigo") Then
REM    sobjShow.JSAlert "Campo ageCodigo encontrado"

REM  Else
REM    sobjShow.JSAlert "Campo ageCodigo não encontrado"

REM  End If

REM  For i = 1 to Session(EditCurrentTable & "FieldNumber")
REM    sobjShow.JSAlert Session(EditCurrentTable & i & "Field")

REM  Next

  If Session("RecordAction") > "" Then
  Dim blnFirst
    blnFirst = True

    EditBeginReadOperation Session("LastTab")

    Do While EditIsThereOperation
      Select Case EditNextOperation
        Case ",", ";", "-", "="

        Case Else
          If EditFindField(EditCurrentTable, EditLastOperation) Then
            If Session(EditCurrentField & "DisableNumber") > 0 Then
              If blnFirst Then
                sobjShow.HTMLCR "<SCRIPT FOR=window EVENT=onload LANGUAGE=""JavaScript"">"
                blnFirst = False

              End If

              sobjShow.HTMLCR "DisableFields('" & Session(EditCurrentField & "Field") & "')"

            End If
          End If

      End Select

    Loop

    If Not blnFirst Then
      sobjShow.HTMLCR "</SCRIPT>"

    End If

  End If

  sobjForm.onChange = ""

  sobjTabela.EndColumn
  sobjTabela.EndRow
  sobjTabela.EndTable

  If Session(EditCurrentTable & "UnitList") > "" Then
    EditGrid

  End If

  sobjTabela.EndColumn
  sobjTabela.EndRow

  sobjTabela.BeginRow
  sobjTabela.BeginColumn
  sobjShow.Center

  sobjShow.CrLf = False

  If Not Session("RecordAction") > "" Then

    If Session("AbsolutePosition") > 1 Then
      sobjForm.Enable = True
      sobjForm.onClick = "return ButtonsTest('" & constrFirst & "')"

    Else
      sobjForm.Enable = False

    End If
    sobjForm.Title = "Primeiro Registro"
    sobjForm.Submit "btmEditor", constrFirst
REM sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdPrimeiroWin.gif", "Vai para o primeiro registro"

    If Session("AbsolutePosition") > 1 Then
      sobjForm.Enable = True
      sobjForm.onClick = "return ButtonsTest('" & constrPrevious & "')"

    Else
      sobjForm.Enable = False

    End If
    sobjForm.Title = "Registro Anterior"
    sobjForm.Submit "btmEditor", constrPrevious
REM    sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdAnteriorWin.gif", "Vai para o registro anterior"

  End If

  sobjForm.Enable = True

  If sparEditEditable = 1 Then
    Select Case Session(EditCurrentTable & "EditOption")
      Case "1"
        If Not Session("RecordAction") > "" Then
          sobjForm.AccessKey = "I"
          sobjForm.onClick = "return ButtonsTest('" & constrAdd & "')"

          sobjForm.Title = "inclui um registro"
          sobjForm.Submit "btmEditor", constrAdd
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdIncluirWin.gif", "Inclui um registro"
          sobjForm.AccessKey = ""
          sobjForm.onClick = "return ButtonsTest('" & constrCopy & "')"
          sobjForm.Title = "Inclui um registro copiando os dados do registro corrente"
          sobjForm.Submit "btmEditor", constrCopy
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdCopiarWin.gif", "Inclui um registro copiando os dados deste"
          sobjForm.AccessKey = "A"
          sobjForm.onClick = "return ButtonsTest('" & constrEdit & "')"
          sobjForm.Title = "Edita o registro corrente"
          sobjForm.Submit "btmEditor", constrEdit
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdAlterarWin.gif", "Entra em modo de alteração"
          sobjForm.AccessKey = "P"
          sobjForm.onClick = "return ButtonsTest('" & constrDelete & "')"
          sobjForm.Title = "Exclui o registro corrente"
          sobjForm.Submit "btmEditor", constrDelete
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdApagarWin.gif", "Exclui este registro"
          sobjForm.AccessKey = ""

        End If

        If Not Session("RecordAction") > "" Then
REM        sobjForm.Submit "btmEditor", constrFind

        Else
          sobjForm.AccessKey = "S"
          sobjForm.onClick = "return ButtonsTest('" & constrSave & "')"
          sobjForm.Title = "Salta as alterações efetuadas"
          sobjForm.Submit "btmEditor", constrSave
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarWin.gif", "Salva as alterações feitas neste registro"
          sobjForm.AccessKey = ""
          sobjForm.onClick = "return ButtonsTest('" & constrSaveAdd & "')"
          sobjForm.Title = "Salta as alterações efetuadas e entra em modo de inclusão"
          sobjForm.Submit "btmEditor", constrSaveAdd
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarIncluirWin.gif", "Salva os dados deste registro e inclui outro"
          sobjForm.onClick = "return ButtonsTest('" & constrSaveCopy & "')"
          sobjForm.Title = "Salta as alterações efetuadas e entra em modo de inclusão copiando os dados do registro corrente"
          sobjForm.Submit "btmEditor", constrSaveCopy
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarCopiarWin.gif", "Salva os dados deste registro e copiando os dados deste"
          sobjForm.AccessKey = "C"
          sobjForm.onClick = "return ButtonsTest('" & constrCancel & "')"
          sobjForm.Title = "Cancela todas as alterações efetuadas"
          sobjForm.Submit "btmEditor", constrCancel
REM          sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdCancelarWin.gif", "Cancela a inclusão deste registro"
          sobjForm.AccessKey = ""

        End If

    End Select
  End If

  If Not Session("RecordAction") > "" Then
    sobjForm.onClick = "return ButtonsTest('" & constrFind & "')"
    sobjForm.Title = "Entra em modo de localização de registros"
    sobjForm.Submit "btmEditor", constrFind
REM    sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdLocalizarWin.gif", "Entra em modo de localização de registros"
    sobjForm.onClick = "return ButtonsTest('" & constrList & "')"
    sobjForm.Title = "Entra no modo de visualização dos registros em lista"
    sobjForm.Submit "btmEditor", constrList
REM    sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdListagemWin.gif", "Vai para a forma de visualização de listagem"

    If Session("AbsolutePosition") = Session("RecordCount") Then
      sobjForm.Enable = False

    Else
      sobjForm.onClick = "return ButtonsTest('" & constrNext & "')"

    End If
    sobjForm.Title = "Próximo Registro"
    sobjForm.Submit "btmEditor", constrNext
REM    sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdProximoWin.gif", "Vai para o próximo registro"

    If Session("AbsolutePosition") = Session("RecordCount") Then
      sobjForm.Enable = False

    Else
      sobjForm.onClick = "return ButtonsTest('" & constrLast & "')"

    End If
    sobjForm.Title = "Último Registro"
    sobjForm.Submit "btmEditor", constrLast
REM    sobjForm.Image "btmEditor", "", "/ZTITools/Img/EditForm/cmdUltimoWin.gif", "Vai para o úlitmo registro"
    sobjForm.onClick = ""

  End If

  sobjForm.onClick = ""
  sobjForm.Enable = True
REM  sobjForm.Reset "Apaga"

  sobjTabela.EndColumn
  sobjTabela.EndRow

  sobjTabela.EndTable

  sobjForm.EndForm

  EditFPLPLink

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowEditor

REM -------------------------------------------------------------------------
REM Monta o link interno do campo corrente
REM -------------------------------------------------------------------------
Public Sub EditMakeGridLink(strTable)
  If Session(strTable & EditCurrentGrid & "GridOption") = 0 Then
    Response.Write "<A HREF=" & sstrThisScriptName & "?" & conPEditOption & "=" & conOptionEdit & "&" & _
                   conPEditWhat & "=" & Session(EditCurrentTable & "WhatList") & conPEditEditableStr & EditGridTarget & " Target=_Top>"

  ElseIf Session(strTable & EditCurrentGrid & "GridOption") = 1 Then
    Response.Write "<A HREF=" & sstrThisScriptName & "?" & conPEditOption & "=" & conOptionFPEdit & "&" & _
                   conPEditWhat & "=" & Session(strTable & "WhatList") & conPEditEditableStr & EditGridTarget & " Target=_Top>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditMakeInternalLink

REM -------------------------------------------------------------------------
REM Monta o link interno das chaves primarias
REM -------------------------------------------------------------------------
Public Function EditGridTarget
Dim strAux, i

  strAux = ""

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Session(EditCurrentTable & i & "IsKey") Then
      If strAux > "" Then
        strAux = strAux & " And "

      End If
      strAux = strAux & Session(EditCurrentTable & i & "Field") & "=" & FormatForSQL(Session(EditCurrentTable & i & "Value"), Session(EditCurrentTable & i & "Type"))

    End If

  Next

  If strAux > "" Then
    EditGridTarget = "&" & conPEditFind & "=" & Server.URLEncode(strAux)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditGridTarget

REM -------------------------------------------------------------------------
REM Monta um Grid na
REM -------------------------------------------------------------------------
Private Sub EditShowGrid
Dim sql, i, strPrincipalTable, intAction
REM  DebugMode

  ClassAss
  strPrincipalTable = EditCurrentTable
  Session("CurrentTable") = Session(EditCurrentTable & EditCurrentGrid & "GridTable")
  Session(strPrincipalTable & "GridWidth") = ""

  sobjShow.HTMLCR "<STYLE> .EF{font: 8pt Arial}</STYLE>"
  sobjForm.Class = "EF"
  If Request.Form("btmEdit") > "" Then
    intAction = 1

  ElseIf Request.Form("btmSave") > "" Then
    intAction = 0
    If sparEditFind > "" Then
      EditGetForm Session(strPrincipalTable & EditCurrentGrid & "GridFields")
    	EditBeginReadOperation Session(strPrincipalTable & EditCurrentGrid & "GridFields")
    	Do While EditIsThereOperation
    	  Select Case EditNextOperation
    	     Case ",", ";", "-"
    	     Case "="
    	       EditIsThereOperation

    	     Case Else
    	       If EditFindField(EditCurrentTable, EditLastOperation) Then
    	         Session(EditCurrentField & "OldValue") = Session(EditCurrentField & "Value")

    	       End If

    	   End Select

    	Loop

      sql = "SELECT * FROM " & EditCurrentTable & " WHERE " & sparEditFind

      sobjCMD.CommandText = sql
      sobjCMD.CommandType = adCmdText
      DebugMessage sql, 0

    	sobjRS2.Open sobjCMD, , adOpenDynamic, adLockOptimistic

      If sobjRS2.EOF Or sobjRS2.BOF Then
        sobjRS2.AddNew

      End If

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
         sobjRS2(Session(EditCurrentTable & i & "Field")) = EditFormatFieldtoSave(Session(EditCurrentTable & _
                                                                  i & "OldValue"), Session(EditCurrentTable & i & "Type"), _
                                                                  Session(EditCurrentTable & i & "Size"), _
                                                                  Session(EditCurrentTable & i & "CharCase"))

      Next
      sobjRS2.Update
      sobjRS2.Close
      sparEditFind = ""
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock

    End If

  ElseIf Request.Form("btmDelete") > "" Then
    intAction = 0
    If sparEditFind > "" Then
      sobjCMD.CommandText = "DELETE FROM " & EditCurrentTable & " WHERE " & sparEditFind
      sobjCMD.CommandType = adCmdText
      DebugMessage "DELETE FROM " & EditCurrentTable & " WHERE " & sparEditFind, 0

    	sobjRS2.Open sobjCMD, , adOpenKeySet, adLockReadOnly

    End If
    sparEditFind = ""

  ElseIf Request.Form("btmInsert") > "" Or sparEditGridOption = "1" Then
    intAction = 2
    sparEditFind = ""

  ElseIf Request.Form("btmCancel") > "" Then
    intAction = 0
    Application.Lock
    Application(EditCurrentTable & EditMakeTarget) = ""
    Application.Unlock

  Else
    intAction = 0

  End If

	If EditFindRelation(strPrincipalTable, EditCurrentTable) Or strPrincipalTable = EditCurrentTable Then
  	sql = "SELECT * FROM " & EditCurrentTable

  	EditBeginReadOperation Session(strPrincipalTable & EditCurrentGrid & "GridFields")

  	Dim Width, strList
  	Width = 0
  	Do While EditIsThereOperation
  	  Select Case EditNextOperation
  	     Case ",", ";", "-"
  	     Case "="
  	       EditIsThereOperation

  	     Case Else
  	       If EditFindField(EditCurrentTable, EditLastOperation) Then
  	         If Not Session(strPrincipalTable & "GridWidth") > "" Then
               If Session(EditCurrentField & "Type") <> "varchar" Then
                 Width = Width + EditSize * 5

               Else
                 Width = Width + EditSize * 6

  	           End If

  	         End If
             If strList > "" Then
  	           strList = strList & ","

  	         End If
  	         strList = strList & Session("CurrentField")
          	 If Session(EditCurrentTable & Session("CurrentField") & "EditOption") = conLookupField Then
          	    sql = sql & " LEFT OUTER JOIN " & Session(EditCurrentTable & Session("CurrentField") & "LookupTable") &_
          	    " on " & Session(EditCurrentTable & Session("CurrentField") & "LookupTable") & "." & Session(EditCurrentTable & Session("CurrentField") & "LookupField") & " = " & _
          	    EditTableOf(EditCurrentTable, i) & "." & Session(EditCurrentTable & Session("CurrentField") & "Field")

          	 End If

  	       End If

  	   End Select

  	Loop

    If Not Session(strPrincipalTable & "GridWidth") > "" Then
      If Width < 200 Then
        Width = 200

      ElseIf Width > 600 Then
        Width = 600

      End If
      Session(strPrincipalTable & "GridWidth") = Width

  	End If

  	SetPutAnd False

  	For i = 1 to Session(strPrincipalTable & sintCurrentRelation & "FieldtoFieldNumber")
  	  If EditFindField(strPrincipalTable, Session(strPrincipalTable & sintCurrentRelation & "," & i & "RelationField1")) Then
    	  sql = sql & PutAnd & Session(strPrincipalTable & sintCurrentRelation & "," & i & "RelationField2") & "=" & FormatForSQL(Session(strPrincipalTable & Session("CurrentField") & "Value"), Session(strPrincipalTable & Session("CurrentField") & "Type"))

  	  End If
    Next

    sobjConn.Close
    sobjConn.Provider = "SQLOLEDB.1"
    sobjConn.Open

    sobjCMD.ActiveConnection = sobjConn
    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText
  	DebugMessage sql, 1

    sobjRS2.CursorLocation = adUseClient

  	sobjRS2.Open sobjCMD, , adOpenKeySet, adLockReadOnly

  	If Not sobjRS2.EOF And Not sobjRS2.BOF Or sparEditGridOption = "1" Then
      sobjBrowse.BeginBrowse sobjRS2

      EditTotalOperation strList, "GridFieldNumber", False, ""

      sobjTabela.Padding = "1"
      sobjTabela.Spacing = "1"
  	  sobjTabela.BeginTable Session(strPrincipalTable & "GridWidth") + 20, , , True

  	  sobjTabela.BeginRow 2, True
  	  EditBeginReadOperation strList
  	  Do While EditIsThereOperation
  	    Select Case EditNextOperation
  	      Case ",", ";", "-", "="

  	      Case Else
        	  Session("CurrentField") = EditLastOperation
            sobjTabela.ColumnWidth = Session(EditCurrentField & "ListSize")
            sobjTabela.ColumnColor = sobjTabela.HeadRowColor
            sobjTabela.CompleteColumn "<CENTER><B><SPAN Class=SS>" & EditCaption(False) & "</SPAN>"
            sobjTabela.ColumnColor = ""

  	     End Select

  	  Loop
  	  If intAction = 0 Then
        strURL = sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionGrid & GetFilter

        sobjForm.BeginForm strURL, "POST", "frmEditGrid", "", ""
        sobjTabela.ColumnWidth = "40"
        sobjTabela.BeginColumn
        sobjForm.Submit "btmInsert", " I "
        sobjTabela.EndColumn
        sobjForm.EndForm

      Else
        sobjTabela.CompleteColumn "&nbsp;"

      End If
      sobjTabela.EndRow

    Dim intPosition, blnFirst
      If sparEditFind > "" And intAction <> 0 Then
REM        DebugMessage sparEditFind, 0
        sobjRS2.Find sparEditFind
        intPosition = sobjRS2.AbsolutePosition
        blnFirst = True

      ElseIf intAction = 2 Then
        blnFirst = True

      Else
        blnFirst = False

      End If

      Dim strField, strURL

  	  Do
  	    If intPosition <> sobjRS2.AbsolutePosition or blnFirst Then
    	    EditBeginReadOperation strList
          If (intAction <> 2 Or Not blnFirst) And Not sparEditGridOption > "" Then
            For i = 1 to Session(EditCurrentTable & "FieldNumber")
              Session(EditCurrentTable & i & "Value") = sobjRS2(Session(EditCurrentTable & i & "Field"))

            Next

          Else
            EditClear conOptionAdd
          	For i = 1 to Session(strPrincipalTable & sintCurrentRelation & "FieldtoFieldNumber")
          	  If EditFindField(EditCurrentTable, Session(strPrincipalTable & sintCurrentRelation & "," & i & "RelationField2")) Then
                Session(EditCurrentField & "OldValue") = FieldValue(strPrincipalTable, Session(strPrincipalTable & sintCurrentRelation & "," & i & "RelationField1"))

          	  End If
            Next

          End If
          strURL = sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionGrid & EditGridTarget & GetFilter

          sobjForm.BeginForm strURL, "POST", "frmEditGrid", "", ""

    	    sobjTabela.BeginRow 1

    	    Do While EditIsThereOperation
    	      Select Case EditNextOperation
    	        Case ",",";","-","="
    	        Case Else
    	          Session("CurrentField") = EditLastOperation
    	          strField = Session("CurrentField")
    	          If blnFirst Then
      	          sobjTabela.BeginColumn
    	            EditShowField False, True, False, False
      	          sobjTabela.EndColumn

     	          Else
      	          If TypeName(EditFieldValue(sobjRS2)) = "Long" Or TypeName(EditFieldValue(sobjRS2)) = "Integer" Or TypeName(EditFieldValue(sobjRS2)) = "Byte" Or _
      	            TypeName(EditFieldValue(sobjRS2)) = "Currency" Or TypeName(EditFieldValue(sobjRS2)) = "Double" Or TypeName(EditFieldValue(sobjRS2)) = "Single" Then
      	            sobjTabela.ColumnAlign = "Right"

      	          End If

      	          sobjTabela.ColumnWidth = Session(EditCurrentField & "ListSize")
      	          sobjTabela.BeginColumn
                  If intAction = 0 Then
      	            EditMakeGridLink strPrincipalTable

      	          End If
      	          Session("CurrentField") = strField
      	          Select Case Session(EditCurrentField & "EditOption")
      	            Case conLookupField, conRadioField, conCheckField, conSelectField
      	              Response.Write BreakString(EditFieldValue(sobjRS2), Session(EditCurrentField & "ListChars"))

      	            Case Else
      	              Response.Write BreakString(PutMask(FormatVariable(EditFieldValue(sobjRS2), Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars"))

      	          End Select
              	  EditEndMakeLink
      	          sobjTabela.ColumnAlign = ""
      	          sobjTabela.EndColumn

      	        End If
    	      End Select
    	    Loop
       	  sobjTabela.ColumnWidth = "40"
      	  sobjTabela.BeginColumn
          If intAction = 0 Then
        	  sobjForm.Submit "btmEdit", "A"
        	  sobjForm.Submit "btmDelete", "E"

          ElseIf blnFirst Then
        	  sobjForm.Submit "btmSave", "S"
        	  sobjForm.Submit "btmCancel", "C"

          Else
            Response.Write "&nbsp;"

        	End If
          sobjTabela.EndColumn
    	    sobjTabela.EndRow
          sobjForm.EndForm

        End If

        If  Not sparEditGridOption > "" Then
          If blnFirst Then
            blnFirst = False
            sobjRS2.MoveFirst

          Else
    	      sobjRS2.MoveNext

      	  End If
        End If

        sobjBrowse.IncLine

  	  Loop While Not sobjRS2.EOF And sobjBrowse.IsBrowsePage

 	    sobjTabela.EndTable

      sobjBrowse.Status = True
      sobjBrowse.EndBrowse
	    sobjBrowse.Bar

    Else
      sobjTabela.BeginTable "450"
      sobjTabela.ColumnAlign = "Center"
      sobjTabela.ColumnVAlign = "Middle"
      sobjTabela.ColumnHeight = "96"
      sobjTabela.BeginRow 2
      sobjTabela.CompleteColumn "<B><A Class=SS HREF=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionGrid & GetFilter & conPEditGridOptionStr & """>Incluir " & Session(EditCurrentTable & "Caption") & "</A>"
      sobjTabela.EndRow
      sobjTabela.EndTable

    End If

    sobjRS2.Close

  Else
    sobjShow.MessageError "Não foi definido um relacionamento entre a tabela folha e a tabela do grid"

  End If
  Session("LastCurrentTable") = strPrincipalTable

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditExeEdition

REM -------------------------------------------------------------------------
REM Rotina que retorna o tamanho do campo para visualizacao
REM -------------------------------------------------------------------------
Public Function EditSizeToShow
  If Session(EditCurrentField & "ShowSize") > "" Then
    EditSizeToShow = Session(EditCurrentField & "ShowSize")

  Else
    If EditSize <= 80 Then
      EditSizeToShow = EditSize

    Else
      EditSizeToShow = 80

    End If

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditSizeToShow

REM -------------------------------------------------------------------------
REM Rotina que retorna o tamanho do campo em caracteres
REM -------------------------------------------------------------------------
Public Function EditSize
  Select Case Session(EditCurrentField & "Type")
    Case conDataInt, conDataTinyInt, conDataSmallInt
      EditSize = Len(CStr(256^Session(EditCurrentField & "Size"))) - 1

    Case conDataFloat
      EditSize = Len(CStr(256^Session(EditCurrentField & "Size"))) + 1

    Case conDataMoney
      EditSize = Len(CStr(256^Session(EditCurrentField & "Size"))) + 10

    Case conDataDateTime
      Select Case Session(EditCurrentField & "EditOption")
        Case conValidTime
          EditSize = 5

        Case conValidDateTime
          EditSize = 16

        Case Else
          EditSize = 10

      End Select

    Case Else
      EditSize = CLng(Session(EditCurrentField & "Size"))

  End Select

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditSize

REM -------------------------------------------------------------------------
REM Rotina que retorna o tamanho do campo para a edicao
REM -------------------------------------------------------------------------
Public Function EditSizeToGet
  If Session(EditCurrentField & "GetSize") > "" Then
    EditSizeToGet = Session(EditCurrentField & "GetSize")

  Else
    EditSizeToGet = EditSize

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditSizeToGet

REM -------------------------------------------------------------------------
REM Rotina que coloca um indicador de campo requerido
REM -------------------------------------------------------------------------
Public Sub EditShowFieldImages
  If Session(EditCurrentField & "RequireMessage") > "" Or Not Session(EditCurrentField & "Null") Then
    Response.Write " "
    ShowImage "/ztitools/CR.gif", "", "Campo obrigatório", "", "", 5, 5

  End If

  If Session(EditCurrentField & "ValidationError") > "" Then
    Response.Write " "
    ShowImage "/ztitools/CE.gif", "", Session(EditCurrentField & "ValidationError"), "", "", 5, 5
    Session(EditCurrentField & "ValidationError") = ""

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Function EditShowFieldImages

REM -------------------------------------------------------------------------
REM Inclui um texto de caption tratado na tela
REM -------------------------------------------------------------------------
Private Sub EditShowCaption(blnForm)
Dim i

  i = InStr(Session(EditCurrentField & "Caption"), "&")

  If i > 0 Then
    If blnForm Then
      Response.Write Left(Session(EditCurrentField & "Caption"), i - 1) & "<U>" & Mid(Session(EditCurrentField & "Caption"), i + 1 , 1) & "</U>" & Mid(Session(EditCurrentField & "Caption"), i + 2)
      sobjForm.AccessKey = ChangeChar(Mid(Session(EditCurrentField & "Caption"), i + 1 , 1))

    Else
      Response.Write Left(Session(EditCurrentField & "Caption"), i - 1) & Mid(Session(EditCurrentField & "Caption"), i + 1)

    End If

  Else
    Response.Write Session(EditCurrentField & "Caption")

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowCaption

REM -------------------------------------------------------------------------
REM Altera os campos de check para se apresentarem como radio
REM -------------------------------------------------------------------------
Public Sub CheckAsRadio
  sstrCheckAsRadio = "1"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub CheckAsRadio

REM -------------------------------------------------------------------------
REM Retorna os campos de check para se apresentarem como check, e nao como radio
REM -------------------------------------------------------------------------
Public Sub NoCheckAsRadio
  sstrCheckAsRadio = ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub NoCheckAsRadio

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo do form
REM -------------------------------------------------------------------------
Public Function EditShowField(blnFocus, blnFieldOnly, blnEnable, blnEmpty)
Dim sql, i, Value, blnFirst

  If Session(EditCurrentField & "ShowOption") > -1 Then
    sobjForm.Title = Session(EditCurrentField & "Hint")
    If blnEmpty Then
      Value = ""

    Else
      If IsNull(Session(EditCurrentField & "Value")) Then
        Value = ""

      Else
       Value = FormatVariable(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))

      End If
    End If

    Select Case Session(EditCurrentField & "EditOption")
      Case conTextField, conHTTPField, conEMailField, conAtualizationField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          sobjShow.BR

        End If

        sobjForm.Text Session(EditCurrentField & "Field"), EditSizeToShow, _
        ToString(Value), EditSizeToGet

      Case conAtualizationField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          sobjShow.BR

        End If

        Response.Write Value

      Case conFileField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          Response.Write "&nbsp;&nbsp;<FONT COLOR=#000066>" & Value & "</FONT>"
          sobjShow.BR

        End If

        sobjForm.File Session(EditCurrentField & "Field")

      Case conImageField
          If Not blnFieldOnly Then
            EditShowCaption True
            Response.Write "&nbsp;&nbsp;<FONT COLOR=#000066>" & Session(EditCurrentField & "ImageRoot") & Value & "</FONT>"
            sobjShow.BR

          End If

          If Value > "" Then
            ShowImage Session(EditCurrentField & "ImageRoot") & Value, "", Session(EditCurrentField & "ImageRoot") & Value, Session(EditCurrentField & "ImageRoot") & Value & " TARGET=_BLANK", "", Session(EditCurrentField & "ImageWidth"), ""

          End If

      Case conPasswordField
        EditShowCaption True
        EditShowFieldImages
        sobjShow.BR
        sobjForm.Password Session(EditCurrentField & "Field"), EditSizeToShow

      Case conTextAreaField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          sobjShow.BR

        End If
        sobjForm.TextArea Session(EditCurrentField & "Field"), CInt(Session(EditCurrentField & "TextAreaHeight")), CInt(EditSizeToShow), ToString(Value)

      Case conUFField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          sobjShow.BR

        End If
        If blnEnable Then
          sobjForm.SelectBegin Session(EditCurrentField & "Field"), 1, False
          If Value > "" Then
            sobjForm.SelectOption Value, Value, True

          Else
            sobjForm.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

          End If
          sobjForm.SelectEnd

        Else
          sobjForm.SelectUF Session(EditCurrentField & "Field"), 1, False, ToString(Value)

        End If

      Case conLookupField
REM         DebugMode
        If (Session(EditCurrentField & "LookupTable") > "") And (Session(EditCurrentField & "LookupField") > "")_
        And (Session(EditCurrentField & "LookupFieldText") > "") Then

          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            sobjShow.BR

          End If

          If Not blnEnable Then
            sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")
            sql = sql & " WHERE " & Session(EditCurrentField & "LookupField") & " = " & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))

          Else
            sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")

          Dim intCurrentField
            If Session(EditCurrentField & "FilterValue") > "" Then 
              sql = sql & " WHERE " & Session(EditCurrentField & "DependentField") & " = "
              sql = sql & FormatForSQL(Session(EditCurrentField & "FilterValue"), Session(EditCurrentField & "FilterFieldType"))

            Else                                                    
              If Session(EditCurrentField & "FilterField") > "" Then 
                intCurrentField = Session("CurrentField")
                If EditFindField(EditCurrentTable, Session(EditCurrentField & "FilterField")) Then
                  sql = sql & " WHERE " & Session(EditCurrentTable & intCurrentField & "DependentField") & " = "
                  sql = sql & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))

                End If
                Session("CurrentField") = intCurrentField
              End If
            End If

            sql = sql & " ORDER BY " & Session(EditCurrentField & "LookupFieldText")

          End If

          DebugMessage sql, 1

          sobjCMD.CommandText = sql
          sobjCMD.CommandType = adCmdText

          sobjForm.SelectBegin Session(EditCurrentField & "Field"), 1, False
          If (Not Value > "" Or IsNull(Value)) Then
            sobjForm.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

          Else
            If blnEnable Then
            	sobjForm.SelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", False

            End If

          End If

          If (Value > "" And Not IsNull(Value)) Or blnEnable Then
            sobjRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly
            Do While Not sobjRS.EOF
              sobjForm.SelectOption sobjRS(Session(EditCurrentField & "LookupField")), _
                                           sobjRS(Session(EditCurrentField & "LookupFieldText")), _
                                           (ToString(FormatVariable(sobjRS(Session(EditCurrentField & "LookupField")), Session(EditCurrentField & "Type"))) _
                                           = ToString(Value))

              sobjRS.MoveNext

            Loop
            sobjRS.Close
            sobjForm.SelectEnd

          End If

        Else
          Err.Raise 300, "EditShowField", "O campo """ & Session(EditCurrentField & "Field") & """, da tabela """ & EditCurrentTable & _
          """, está definido como lookup, mas sua definição de lookup não foi feita ou não está completa"

        End If

      Case conCheckField
        If Not sstrCheckAsRadio > "" Then
          sobjForm.CheckBox Session(EditCurrentField & "Field"), Session(EditCurrentField & ",1Value"),_
          "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",1Value")))
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages

          End If

        Else
          Response.Write Session(EditCurrentField & ",1Caption")
          sobjForm.Radio Session(EditCurrentField & "Field"), Session(EditCurrentField & ",1Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",1Value")))
          Response.Write " "
          Response.Write Session(EditCurrentField & ",2Caption")
          sobjForm.Radio Session(EditCurrentField & "Field"), Session(EditCurrentField & ",2Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",2Value")))

        End If

      Case conRadioField
        If Not blnFieldOnly Then
          EditShowCaption True

        End If
        Dim j
        sobjTabela.BeginTable "100%", "", 1
        sobjTabela.ColumnVAlign = "Middle"
        blnFirst = True
        For i = 1 to Session(EditCurrentField & "ValueNumber")
          If (i - 1) Mod Session(EditCurrentField & "RadioColumns") = 0 Then
            If not blnFirst Then
              sobjTabela.EndColumn
              sobjTabela.EndRow
              sobjTabela.BeginRow 1

            End If
            blnFirst = False
REM        DebugMessage i, 2
          End If

          sobjTabela.ColumnWidth = 10
          sobjTabela.BeginColumn
          sobjForm.Radio Session(EditCurrentField & "Field"), Session(EditCurrentField & "," & i & "Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & "," & i & "Value")))
          sobjTabela.EndColumn
          sobjTabela.ColumnWidth = ""
          sobjTabela.BeginColumn
          Response.Write "&nbsp;" & Session(EditCurrentField & "," & i & "Caption")
          sobjTabela.EndColumn

        Next
        sobjTabela.EndRow
        sobjTabela.EndTable

      Case conSelectField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          sobjShow.BR

        End If

        sobjForm.SelectBegin Session(EditCurrentField & "Field"), 1, False
        If Not Value > "" Then
          sobjForm.SelectOption "", "", True

        End If
        For i = 1 to Session(EditCurrentField & "ValueNumber")
          sobjForm.SelectOption Session(EditCurrentField & "," & i & "Value"), Session(EditCurrentField & "," & i & "Caption"), _
                                           (ToString(Value) = ToString(Session(EditCurrentField & "," & i & "Value")))

        Next
        sobjForm.SelectEnd

      Case conSeparateDateField
        If Not blnFieldOnly Then
          EditShowCaption True
          EditShowFieldImages
          sobjShow.BR

        End If

        sobjForm.SelectBegin Session(EditCurrentField & "Field") & "DayValue", 1, False

        For i = 1 to 31
          sobjForm.SelectOption LeadingZeroes(i, 2), LeadingZeroes(i, 2), i = Day(Value)

        Next

        sobjForm.SelectEnd

        Response.Write "<B>/</B>"

        sobjForm.SelectBegin Session(EditCurrentField & "Field") & "MonthValue", 1, False

        For i = 1 to 12
          sobjForm.SelectOption LeadingZeroes(i, 2), LeadingZeroes(i, 2), i = Month(Value)

        Next

        sobjForm.SelectEnd

        Response.Write "<B>/</B>"

        Dim BeginYear, EndYear

        BeginYear = Session(EditCurrentField & "BeginYear")
        EndYear   = Session(EditCurrentField & "EndYear")

        If Not BeginYear > "" Then
          BeginYear = Year(Now) - 20

        End If

        If Not EndYear > "" Then
          EndYear   = Year(Now) + 20

        End If

        sobjForm.SelectBegin Session(EditCurrentField & "Field") & "YearValue", 1, False

        For i = BeginYear to EndYear
          sobjForm.SelectOption i, i, i = Year(Value)

        Next

        sobjForm.SelectEnd

    End Select

    If blnFocus Then
      blnFocus = False
      sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
      sobjShow.HTMLCR "  document.frmEditForm." & Session(EditCurrentField & "Field") & ".focus();"
      sobjShow.HTMLCR "</SCRIPT>"

    End If

  Else
    EditShowFieldAsText

  End If

  EditShowField = blnFocus

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditShowField

REM -------------------------------------------------------------------------
REM Retorna o valor 1 se o parametro passado for verdadeiro
REM -------------------------------------------------------------------------
Private Function BitBln(blnValor)

  If blnValor Then
    BitBln = 1

  Else
    BitBln = 0

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function BitBln

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo como texto
REM -------------------------------------------------------------------------
Public Sub EditShowFieldAsText
Dim sql, i

  If EditFindField(EditCurrentTable, EditLastOperation) Then
    Response.Write "<FONT SIZE=-1>"
    Select Case Session(EditCurrentField & "EditOption")
      Case conTextField, conUFField, conTextAreaField, conAtualizationField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write " : "

        If (Session(EditCurrentField & "ExternalLink") > "") Then
          EditBeginMakeLink True

        End If

        Response.Write PutMask(FormatVariable(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter"))
        EditEndMakeLink

      Case conHTTPField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write " : "
        Response.Write "<A HRef=http://" & Session(EditCurrentField & "Value") & ">" & Session(EditCurrentField & "Value") & "</A>"

      Case conEMailField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write " : "
        Response.Write "<A HRef=mailto:" & Session(EditCurrentField & "Value") & ">" & Session(EditCurrentField & "Value") & "</A>"

      Case conLookupField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write " : "
        If Session(EditCurrentField & "Value") > "" Then
          sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")
          sql = sql & " WHERE " & Session(EditCurrentField & "LookupField") & "="
          sql = sql & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))

          sobjCMD.CommandText = sql

REM        DebugMessage sql, 1

          sobjRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

          If Not sobjRS.EOF And Not sobjRS.BOF Then
            EditBeginMakeLink True
            Response.Write sobjRS(Session(EditCurrentField & "LookupFieldText"))
            EditEndMakeLink

          End If
          sobjRS.Close

        End If

      Case conCheckField
        Response.Write "<B>" & EditCaption(False) & "</B> "
        EditBeginMakeLink True
        If CStr(Session(EditCurrentField & "Value")) = CStr(Session(EditCurrentField & ",1Value")) Then
          Response.Write Session(EditCurrentField & ",1Caption")

        Else
          Response.Write Session(EditCurrentField & ",2Caption")

        End If
        EditEndMakeLink

      Case conRadioField, conSelectField
        Response.Write "<B>" & EditCaption(False) & "</B> "
        EditBeginMakeLink True
        For i = 1 to Session(EditCurrentField & "ValueNumber")
          If CStr(Session(EditCurrentField & "Value")) = CStr(Session(EditCurrentField & "," & i & "Value")) Then
            Response.Write Session(EditCurrentField & "," & i & "Caption")

          End If

        Next
        EditEndMakeLink

    End Select
REM    Response.Write "</FONT>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowFieldAsText

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo como texto
REM -------------------------------------------------------------------------
Public Function EditFieldValue(objRS2)
Dim sql, Value, i

  Select Case Session(EditCurrentField & "EditOption")
    Case conTextField, conUFField, conTextAreaField, conAtualizationField
REM  DebugMessage Session(EditCurrentField & "Field"), 2
      Value = objRS2(Session(EditCurrentField & "Field"))

    Case conHTTPField
      Value = "<A HRef=" & objRS2(Session(EditCurrentField & "Field")) & ">" & objRS2(Session(EditCurrentField & "Field")) & "</A>"

    Case conEMailField
      Value = "<A HRef=""mailto:" & objRS2(Session(EditCurrentField & "Field")) & """>" & objRS2(Session(EditCurrentField & "Field")) & "</A>"

    Case conLookupField
      Value = objRS2(Session(EditCurrentField & "LookupFieldText"))

    Case conCheckField
      If Not IsNull(objRS2(Session(EditCurrentField & "Field"))) Then
        If CStr(objRS2(Session(EditCurrentField & "Field"))) = CStr(Session(EditCurrentField & ",1Value")) Then
          Value = Session(EditCurrentField & ",1Caption")

        Else
          Value = Session(EditCurrentField & ",2Caption")

        End If
      Else
        Value = Session(EditCurrentField & ",2Caption")

      End If


    Case conRadioField, conSelectField
      For i = 1 to Session(EditCurrentField & "ValueNumber")
        If objRS2(Session(EditCurrentField & "Field")) = Session(EditCurrentField & "," & i & "Value") Then
          Value = Session(EditCurrentField & "," & i & "Caption")

        End If

      Next

  End Select

  EditFieldValue = Value

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditShowFieldAsText

%>