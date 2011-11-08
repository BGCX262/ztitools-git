<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/URL.inc
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
REM Monta o link do campo corrente
REM -------------------------------------------------------------------------
Public Sub EditBeginMakeLink(ByVal blnPrefforExtenal)
  If blnPrefforExtenal Then
    If Session(EditCurrentField & "ExternalLink") > "" Then
      Response.Write "<A HREF=" & Session(EditCurrentField & "ExternalLink") & ">"

    Else
      If Session(EditCurrentField & "InternalLink") > "" Then
        EditMakeInternalLink

      End If
    End If

  Else
    If Session(EditCurrentField & "InternalLink") > "" Then
      EditMakeInternalLink

    Else
      If Session(EditCurrentField & "ExternalLink") > "" Then
        Response.Write "<A HREF=" & Session(EditCurrentField & "ExternalLink") & ">"

      End If
    End If

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditBeginMakeLink

REM -------------------------------------------------------------------------
REM Monta o link interno do campo corrente
REM -------------------------------------------------------------------------
Public Sub EditMakeInternalLink
Dim strURL, i

  strURL = "<A HREF=" & sstrThisScriptName & "?" & conPEditOption & "=" & conOptionFPEdit & "&" & _
           conPEditWhat & "=" & Session(EditCurrentField & "InternalLink")

  If Session(EditCurrentTable & "WhatList") = Session(EditCurrentField & "InternalLink") Then
    strURL = strURL & conPEditEditableStr

  End If

  strURL = strURL & EditTargetOfLinks & ">"

  Response.Write strURL

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditMakeInternalLink

REM -------------------------------------------------------------------------
REM Monta o link interno do campo corrente
REM -------------------------------------------------------------------------
Public Function EditTargetOfLinks
Dim strAux, i, intCurrentField

  For i = 1 to Session(EditCurrentField & "ILTargetNumber")
    If strAux > "" Then
      strAux = strAux & " And "

    End If
    If sobjRS2.State = 0 Then
      strAux = strAux & Session(EditCurrentField & "," & i & "ILTarget") & "=" & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))

    Else
      intCurrentField = Session("CurrentField")
      If EditFindField(EditCurrentTable, Session(EditCurrentField & "," & i & "ILTarget")) Then
REM      sobjShow.MessageError Session(EditCurrentField & "," & i & "ILTarget")
        strAux = strAux & Session(EditCurrentField & "Field") & "=" & FormatForSQL(sobjRS2(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"))      Session("CurrentField") = intCurrentField
        Session("CurrentField") = intCurrentField
      
      Else
        Session("CurrentField") = intCurrentField
        strAux = strAux & Session(EditCurrentField & "," & i & "ILTarget") & "=" & FormatForSQL(sobjRS2(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"))      Session("CurrentField") = intCurrentField
      
      End If
    End If

  Next
  If strAux > "" Then
    EditTargetOfLinks = "&" & conPEditFind & "=" & Server.URLEncode(strAux)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditTargetOfLinks

REM -------------------------------------------------------------------------
REM Monta o link interno das chaves primarias
REM -------------------------------------------------------------------------
Public Function EditMakeTarget
Dim strAux, i

  strAux = ""

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Session(EditCurrentTable & i & "IsKey") Then
      If strAux > "" Then
        strAux = strAux & " And "

      End If
      strAux = strAux & Session(EditCurrentTable & i & "Field") & "=" & FormatForSQL(Session(EditCurrentTable & i & "OldValue"), Session(EditCurrentTable & i & "Type"))
REM      sobjShow.MessageError Session(EditCurrentTable & i & "OldValue") & "ksjhgfsk"
REM      sobjShow.MessageError Session(EditCurrentTable & i & "Field") & EditCurrentTable & i & "Field"

    End If

  Next

  If strAux > "" Then
    EditMakeTarget = "&" & conPEditFind & "=" & Server.URLEncode(strAux)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditMakeTarget

REM -------------------------------------------------------------------------
REM Remonta o ultimo link dos valores campos
REM -------------------------------------------------------------------------
Public Function EditRemakeTarget
Dim strAux, i

  If sparEditFind > "" then
    EditRemakeTarget = "&" & conPEditFind & "=" & sparEditFind

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditRemakeTarget

REM -------------------------------------------------------------------------
REM Monta o fim link do campo corrente
REM -------------------------------------------------------------------------
Public Sub EditEndMakeLink
  If Session(EditCurrentField & "ExternalLink") > "" Or Session(EditCurrentField & "InternalLink") > "" Then
    sobjShow.HTMLCR "</A>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditEndMakeLink

REM -------------------------------------------------------------------------
REM
REM -------------------------------------------------------------------------
Public Function GetTab
  If sparEditTab > "" Then
    GetTab = "&" & conPEditTab & "=" & Server.URLEncode(sparEditTab)

  Else
    GetTab = ""

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function GetEditable

REM -------------------------------------------------------------------------
REM
REM -------------------------------------------------------------------------
Public Function GetEditable
  If sparEditEditable > "" Then
    GetEditable = "&" & conPEditEditable & "=" & Server.URLEncode(sparEditEditable)

  Else
    GetEditable = ""

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function GetEditable

REM -------------------------------------------------------------------------
REM Retorna filtro se tiver filtro no instante
REM -------------------------------------------------------------------------
Public Function GetFilter
  If sparEditFilter > "" Then
    GetFilter = "&" & conPEditFilter & "=" & Server.URLEncode(sparEditFilter)

  Else
    GetFilter = ""

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function

REM -------------------------------------------------------------------------
REM Retorna os valores padrao dos campos se estes existirem
REM -------------------------------------------------------------------------
Public Function GetDefaults
  If sparEditDefaults > "" Then
    GetDefaults = "&" & conPEditDefaults & "=" & Server.URLEncode(sparEditDefaults)

  Else
    GetDefaults = ""

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da GetDefaults

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Function GetList
  If sparEditList > "" Then
    GetList = "&" & conPEditList & "=" & Server.URLEncode(sparEditList)

  Else
    GetList = ""

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Private Function EditMakeFilter(ByVal strFields)
Dim strAux

  EditBeginReadOperation strFields

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",",";","-","="

      Case Else
        EditFindField EditCurrentTable, EditLastOperation
        If strAux > "" Then
          strAux = strAux + " AND "

        End If
        strAux = strAux & Session(EditCurrentField & "Field") & "=" & FormatForSQL(sobjRS2(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"))

    End Select

  Loop

  EditMakeFilter = Server.URLEncode(strAux)

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditMakeFilter

REM -------------------------------------------------------------------------
REM Retorna a parte da QueryString referente a qual listagem esta sendo usada
REM -------------------------------------------------------------------------
Private Function EditWhatList

  If sparEditWhat > "" Then
    EditWhatList = "&" & conPEditWhat & "=" & sparEditWhat

  ElseIf Session(EditCurrentTable & "WhatList") > "" Then
    EditWhatList = "&" & conPEditWhat & "=" & Session(EditCurrentTable & "WhatList")

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditWhatList

REM -------------------------------------------------------------------------
REM Retorna a parte da QueryString referente a listagem, e modo de insercao
REM -------------------------------------------------------------------------
Private Function EditWhereHow

  If sparEditEditable = "1" Then
    EditWhereHow = EditWhatList & GetEditable

  Else
    EditWhereHow = EditWhatList

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditWhereHow

REM -------------------------------------------------------------------------
REM Retorna a parte da QueryString referente a registro, listagem, e modo de insercao
REM -------------------------------------------------------------------------
Private Function EditWhatWhereHow

  If sparEditEditable = "1" Then
    EditWhatWhereHow = EditMakeTarget & EditWhatList & GetEditable & GetFilter

  Else
    EditWhatWhereHow = EditMakeTarget & EditWhatList & GetFilter

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditWhatWhereHow

REM -------------------------------------------------------------------------
REM Definie o Home para a pagina inicial - Ruben 1/19/99 11:21PM
REM -------------------------------------------------------------------------
Public Sub EditFPLPLinkSet(ByVal strLocal)

  Session("EditFPLPLink") = strLocal

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFPLPLinkSet

REM -------------------------------------------------------------------------
REM Cria os links de rodape que manda para a primeira pagina e para a ultima pagina executada
REM -------------------------------------------------------------------------
Public Sub EditFPLPLink
  Paragraph
  sobjShow.Center

  Response.Write "<FONT SIZE=2>"

  If sparEditOption <> conOptionFirstPage Then
    If Session("EditFPLPLink") > "" Then
      Response.Write "<A HREF=" & Session("EditFPLPLink")

    Else
      Response.Write "<A HREF=" & sstrThisScriptName

    End If

    Response.Write ">Página Inicial</a> "

  End If

  If QSLast > "" Then
    Response.Write "<A HREF=" & QSLast & ">Última Opção</a>"

  End If

  Response.Write "</FONT>"

  sobjShow.CenterEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFPLPLink

%>