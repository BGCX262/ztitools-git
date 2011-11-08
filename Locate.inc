<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Locate.inc
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
Private Sub LocateFirstPage
  
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

  FormBegin sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionExeLocate, "POST", "frmLocate", "", ""

  TableSetSpacing 2
  TableSetPadding 2
  TableSetColumnNoWrap True

  TableNormalBegin "60%"
  
  TableSetColumnColSpan ""
  
  TableBeginHeadRow 3
  
  TableSetColumnAlign "CENTER"
  
  TableHeadColumn "Localização - " & Session(EditCurrentTable & "Caption")
  
  TableSetColumnAlign ""
  
  TableEndHeadRow
  
  TableBeginRow 2
  TableSetColumnVAlign "Top"
  
  REM -----------------------------------------------------------------------
  REM Coluna do Mes Corrente
  REM -----------------------------------------------------------------------
  TableBeginColumn
                             
  ShowHTMLCR "Selecione um campo da tabela e digite o valor a ser localizado."
  Paragraph
                                                               
Dim i  
  FormSelectBegin "txtCampo", 1, False
  For i = 1 to Session(EditCurrentTable & "FieldNumber")                   
    Session("CurrentField") = i
    If (Session(EditCurrentField & "EditOption") = conTextField) And (Not IsExternalField) Then
      FormSelectOption i, EditCaption(False), False
      
    End If  
  
  Next
  FormSelectEnd

  FormInputTextMaxLength "txtValor", 40, "", 100
              
  ShowHTMLCR "&nbsp;Igual"            
  FormInputCheckBox "blnIgual", 1, "left", 0
  BR
  
  TableEndColumn
  TableEndRow

  TableBeginRow 1
  TableBeginColumn
  Center
  FormInputSubmit "btmLocaliza", "Localizar"
  FormInputSubmit "btmLocaliza", "Listagem"
  TableEndColumn
  TableEndRow
  
  FormEnd

  TableNormalEnd
    
  TableSetColumnNoWrap False
  
  CenterEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FistPage

REM -------------------------------------------------------------------------
REM Executa                                                             
REM -------------------------------------------------------------------------
Private Sub ExeLocateFirstPage
Dim intCampo, strValor
  
  If Request.Form("btmLocaliza") = "Localizar" Then
    intCampo = CInt(Request.Form("txtCampo"))
    strValor = Request.Form("txtValor")
                                                                                         
    If strValor > "" Then
      Select Case Session(EditCurrentTable & intCampo & "Type")
        Case "date"   
          strValor = NormalizeString(strValor)
          strValor = Mid(strValor, 1, 2) & "/" & Mid(strValor, 3, 2) & "/" & Mid(strValor, 5)
        	If Not IsDate(strValor) Then
        		Session("Message")= ErrorImage & "Data inválida"
        	  strValor          = ""
    
        	End If
        
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
        
      End Select
    
    Else
      Session("Message") = ErrorImage & "Digite o valor a ser localizado"
      strValor           = ""
    
    End If
    
    If strValor > "" Then
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & "&" & "F" & intCampo & "=" & EditFormatField(strValor, Session(EditCurrentTable & intCampo & "Type"), Session(EditCurrentTable & intCampo & "Size")) & "&" & conPEditOrderField & "=" & intCampo
    
    Else
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionLocate & EditWhatList
    
    End If
  
  Else
    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList
 
  End if

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FistPage

%>