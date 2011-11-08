<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Find.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Localizacao
REM  Cria‡ao     : 11:00h 18/01/1999
REM  Local       : Brasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
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
REM Mostra a Primeira Pagina de localizacao
REM -------------------------------------------------------------------------
Private Sub FindFirstPage
Dim i

  ShowHTMLCR "<STYLE>"
  ShowHTMLCR "INPUT.BTM {font: 8pt Arial}"
  ShowHTMLCR "</STYLE>"
  
  CheckAsRadio
  BR
  Center

  If Session("Message") > "" Then
    TableBegin "", 1, "60%"
    TableBeginRow 2
    TableBeginColumn
    ShowHTMLCR Session("Message")
    TableEndColumn
    TableEndRow
    TableEnd

    Session("Message") = ""

  End If

  FormBegin sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionExeFind, "POST", "frmFind", "", ""

  TableBegin "Localização de Registros - " & Session(EditCurrentTable & "Caption"), 4, "85%"

  Session("NumberOfFinds") = 0
  Session("LastOption") = sparEditOption

  If sparEditOption = conOptionFindCommon Then
    TableSetColumnAlign "left"
    TableBeginHeadRow 1
    TableHeadColumn "Selecione um campo da tabela e digite o valor a ser localizado."
    TableEndHeadRow

    TableBeginRow 1
    TableSetColumnVAlign "Top"
    TableBeginColumn

    FormSelectBegin "txtCampo1", 1, False
    For i = 1 to Session(EditCurrentTable & "FieldNumber")
      Session("CurrentField") = i
      If (Session(EditCurrentField & "EditOption") = conTextField) And (Not IsExternalField) And (EditCaption(False) > "") Then
        FormSelectOption Session(EditCurrentField & "Field"), EditCaption(False), False

      End If

    Next
    FormSelectEnd

    FormInputTextMaxLength "txtValor", 40, "", 100
    Nbsp
    ShowHTMLCR "Igual"
    FormInputCheckBox "blnIgual", 1, "left", 0
    BR
    FormInputHidden "txtOperador1", " like "

    Session("NumberOfFinds") = 1

  Else
    EditBeginReadOperation Session(EditCurrentTable & "FieldFind")

    TableBeginHeadRow 1
    TableHeadColumn "Campo"
    TableHeadColumn "Operador"
    TableHeadColumn "Valor"
    TableEndHeadRow
    
    Do While EditIsThereOperation
      Select Case EditNextOperation
         Case ","

         Case Else
           If EditFindField(EditCurrentTable, EditLastOperation) Then
             Session("NumberOfFinds") = Session("NumberOfFinds") + 1
             TableBeginRow 1

             TableBeginColumn
             EditShowCaption False
             TableEndColumn

             TableBeginColumn

             Dim strEditOption
               strEditOption = Session(EditCurrentField & "EditOption")

             If (strEditOption = conTextField) Or (strEditOption = conAtualizationField) Or (strEditOption = conSeparateDateField) Or (strEditOption = conTextAreaField) Then
               FormSelectBegin "txtOperador" & Session("NumberOfFinds"), 1, False
               FormSelectOption "=", "Igual a", True
               If (Session(EditCurrentField & "Type") = "varchar") Then
                 FormSelectOption " like ", "Contendo o texto", False

               Else
                 FormSelectOption ">", "Maior que", False
                 FormSelectOption "<", "Menor que", False
                 FormSelectOption ">=", "Maior ou igual a", False
                 FormSelectOption "<=", "Menor ou igual a", False

               End If  
               FormSelectEnd

             Else
               ShowHTMLCR "Igual a"
               FormInputHidden "txtOperador" & Session("NumberOfFinds"), "="

             End If

             TableEndColumn

             TableBeginColumn
             EditShowField False, True, True, True
             TableEndColumn

             TableEndRow
             FormInputHidden "txtCampo" & Session("NumberOfFinds"), Session(EditCurrentField & "Field")

          End If

      End Select

    Loop
    TableSetColumnColSpan 4

  End If

  TableBeginRow 1
  TableBeginColumn
  Center

  FormSetClass "BTM"
  FormInputSubmit "btmLocaliza", "Localizar"
  FormInputReset "Limpar"
  If Session("LastToFind") = conOptionAZEdit Then
    FormInputSubmit "btmLocaliza", "Listagem"
  
  Else
    FormInputSubmit "btmLocaliza", "Edição"
    
  End If  
  TableEndColumn
  TableEndRow

  TableEnd

  FormEnd

  TableSetColumnNoWrap False

  CenterEnd

  EditFPLPLink

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FistPage

REM -------------------------------------------------------------------------
REM Executa
REM -------------------------------------------------------------------------
Private Sub ExeFindFirstPage
Dim strCampo, strValor, strOperador, strFind, i, blnSomeToFind, blnIgual
  
  If Request.Form("btmLocaliza") = "Localizar" Then
    blnIgual = Request.Form("blnIgual")
    blnSomeToFind = False
    For i = 1 to Session("NumberOfFinds")
      strCampo = Request.Form("txtCampo" & i)
      EditFindField EditCurrentTable, strCampo
      strOperador = Request.Form("txtOperador" & i)

      If Session("LastOption") = conOptionFindCommon Then
        strValor = Request.Form("txtValor")
        If (blnIgual = "1") Or (Session(EditCurrentField & "Type") <> "varchar") Then
          strOperador = "="
 
        End If
      
      Else
        strValor = Request.Form(Session(EditCurrentField & "Field"))
        
      End If

      If strValor > "" Then
        blnSomeToFind = True   
        Select Case Session(EditCurrentField & "Type")
          Case "date", "datetime"
            strValor = ImportDate(strValor, conDateTypeDDMMAAAA)
            If Not IsDate(strValor) Then
          		Session("Message")= ErrorImage & "Data inválida"
          	  strValor          = ""

          	End If
        	  strValor            = SQLString(strValor)

          Case "int"
            strValor = NormalizeString(strValor)
          	If Not IsNumeric(strValor) Then
          		Session("Message")= ErrorImage & "Número inteiro inválido"
          	  strValor          = ""

          	End If

          Case "float"
          	If Not IsNumeric(strValor) Then
          		Session("Message")= ErrorImage & "Número real inválido"
          	  strValor          = ""

          	End If

          Case "money"
            If InStr(strValor, "$") Then
              strValor = Mid(strValor, InStr(strValor, "$") + 1)

            End If
          	If Not IsNumeric(strValor) Then
          		Session("Message")= ErrorImage & "Valor inválido"
          	  strValor          = ""

          	End If

          Case "varchar"
            If strOperador = "=" Then
           	  strValor            = SQLString(strValor)

            Else
           	  strValor            = SQLString("%" & strValor & "%")

            End If

        End Select

        If Not Session("Message") > "" Then
          If strFind > "" Then
            strFind = strFind & " And " & strCampo & strOperador & strValor

          Else
            strFind = strFind & strCampo & strOperador & strValor

          End If
        End If
      End If

    Next

    If Not blnSomeToFind Then
      Session("Message") = ErrorImage & "Digite o valor a ser localizado"

    End If

    If Session("Message") > "" Then
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & Session("LastOption") & EditWhatList

    Else
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & Session("LastToFind") & conPEditEditableStr & EditWhatList & "&" & conPEditFilter & "=" & Server.URLEncode(strFind)

    End If

  Else
    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & Session("LastToFind") & conPEditEditableStr & EditWhatList

  End if

  Session("NumberOfFinds") = ""
  Session("LastOption") = ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FistPage

%>