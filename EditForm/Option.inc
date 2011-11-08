<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Option.inc
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
REM Rotina que seleciona
REM -------------------------------------------------------------------------
Public Sub EditShowOptions(ByVal strShowOption)
  
  Select Case strShowOption
    Case conOptionFirstPage
      Session("RecordAction") = ""
      ShowFirstPage
      QSSave

    Case conOptionEditor
      EditShowEditor

    Case conOptionFPEdit
      Session("RecordAction") = ""
      EditReadFieldValues False
      EditShowEditor
      QSSave

    Case conOptionAZEdit
      Session("RecordAction") = ""
      EditShowAZ
      QSSave

    Case conOptionSummary
      Session("RecordAction") = ""
      EditShowSummary
      QSSave

    Case conOptionFindCommon, conOptionFindAdvanced
      FindFirstPage
      QSSave

    Case conOptionGrid
      EditShowGrid
      Session("CurrentTable") = Session("LastCurrentTable") 

  End Select

  If Session("EditOption") <> strShowOption Then
    Session("LastEditOption") = Session("EditOption")
    Session("EditOption")     = strShowOption
    QSSave

  End If  
  
  If Request.Form("Order") > "" Then
    Session("Descendente") = Request.Form("Descendente")
    Session("Order")       = Request.Form("Order")

  Else
    If Session(EditCurrentTable & "OrderNumber") > 0 Then
      Session("Order") = 1

    End If
  End If

REM  ClearAllSessions

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowOptions

REM -------------------------------------------------------------------------
REM Rotina que mostra a tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditExeOptions(ByVal strShowOption)
REM  If strShowOption <> conOptionFirstPage And Session("CurrentTable") = "" Then
REM    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFirstPage
    
REM  End If  

  Select Case strShowOption                                    
    Case conOptionExeEdition
      EditExeEdition

    Case conOptionAdd, conOptionCopy
      Session("RecordAction") = "1"
      EditClear strShowOption
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
      EditWhatWhereHow

    Case conOptionEdit
      Session("RecordAction") = "2"
      If EditReadFieldValues(True) Then
        Application.Lock
        Application(EditCurrentTable & EditMakeTarget) = Now()
        Application.Unlock
        Session("InEdition") = EditCurrentTable & EditMakeTarget
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
        EditWhatWhereHow & GetTab

      Else
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhatWhereHow & GetTab

      End If

    Case conOptionDelete
      Session("RecordAction") = ""
      EditDelete
      If Session("AbsolutePosition") = 1 Then
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhereHow & "&" & conPEditMove & "=" & conMoveFirst

      ElseIf Session("AbsolutePosition") = Session("RecordCount") Then     
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhereHow & "&" & conPEditMove & "=" & conMoveLast

      Else
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhatWhereHow & "&" & conPEditMove & "=" & Session("AbsolutePosition") - 2

      End If

    Case conOptionSave
      EditSave
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock
      Session("InEdition") = ""
      Session("RecordAction") = ""
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow & GetTab

    Case conOptionSaveAdd
      EditSave
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock
      Session("InEdition") = ""
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd & _
      EditWhatWhereHow

    Case conOptionSaveCopy
      EditSave
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock
      Session("InEdition") = ""
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionCopy & _
      EditWhatWhereHow

    Case conOptionExeFind
      ExeFindFirstPage

  End Select

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditExeOptions

REM -------------------------------------------------------------------------
REM Executa o comando dos botoes do Editor
REM -------------------------------------------------------------------------
Private Sub EditExeEdition
  
  Select Case Request.Form("btmEditor")
    Case constrAdd
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd & _
      EditWhatWhereHow & GetDefaults

    Case constrCopy
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionCopy & _
      EditWhatWhereHow

    Case constrEdit
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEdit & _
      EditWhatWhereHow & GetTab

    Case constrDelete
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionDelete & _
      EditWhatWhereHow

    Case constrSave
      EditGetForm Session("LastTab")
      MakeServerValidation conOptionSave

    Case constrSaveAdd
      EditGetForm Session("LastTab")
      MakeServerValidation conOptionSaveAdd

    Case constrSaveCopy
      EditGetForm Session("LastTab")
      MakeServerValidation conOptionSaveCopy

    Case constrList
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
      EditWhatList & GetFilter

    Case constrCancel
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock
      Session("InEdition") = ""
      Response.Redirect QSLast
      Session("RecordAction") = ""

    Case constrFirst
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhereHow & "&" & conPEditMove & "=" & conMoveFirst

    Case constrPrevious
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhereHow & "&" & conPEditMove & "=" & Session("AbsolutePosition") - 2

    Case constrNext
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhereHow & "&" & conPEditMove & "=" & Session("AbsolutePosition")

    Case constrLast
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhereHow & "&" & conPEditMove & "=" & conMoveLast

    Case constrFind
      If Session(EditCurrentTable & "FieldFind") > "" Then
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindAdvanced & _
        EditWhatWhereHow

      Else
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindCommon & _
        EditWhatWhereHow

      End If

  End Select

  If Request.Form("hidEditor") > "" Then
    EditGetForm Session("LastTab")
    If Session("IsEdition") > "" Or Session("RecordAction") > "" Then
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
      EditWhatWhereHow & "&" & conPEditTab & "=" & Request.Form("hidEditor")

    Else
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow & "&" & conPEditTab & "=" & Request.Form("hidEditor")

    End If
  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditExeEdition

REM -------------------------------------------------------------------------
REM Rotina que le as query strings
REM -------------------------------------------------------------------------
Public Sub EditQueryString
  
  SetShowComment False
  sparEditOption     = lCase(Request.QueryString(conPEditOption))
  sparEditWhat       = lCase(Request.QueryString(conPEditWhat))
  sparEditMove       = lCase(Request.QueryString(conPEditMove))
  sparEditEditable   = lCase(Request.QueryString(conPEditEditable))
  sparEditFind       = LCase(Request.QueryString(conPEditFind))
  sparEditFilter     = LCase(Request.QueryString(conPEditFilter))
  sparEditList       = LCase(Request.QueryString(conPEditList))
  sparEditDefaults   = Request.QueryString(conPEditDefaults)
  sparEditTab        = Request.QueryString(conPEditTab)
  sparEditGridOption = Request.QueryString(conPEditGridOption)

  If Not sparEditOption > "" Then
    sparEditOption   = conOptionFirstPage

  End If
  If Not sparEditEditable > "" Then
    sparEditEditable = 0

  End If

  If Not sparEditTab > "" Then
    sparEditTab = 0

  End If

  Session("OrderField") = lCase(Request.QueryString(conPEditOrderField))
  Session("OrderDesc")  = lCase(Request.QueryString(conPEditOrderDesc))

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditQueryString

%>