<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Validate.inc
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
REM Redireciona o cliente para o form ou para a opcao do parametro
REM -------------------------------------------------------------------------
Private Sub MakeServerValidation(strRedirectOption)
  If ServerValidates Then
    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & strRedirectOption & _
    EditWhatWhereHow & GetTab

  Else
    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
    EditWhatWhereHow & GetTab

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MakeServerValidation

REM -------------------------------------------------------------------------
REM Retorna o HTML que monta a imagem de erro
REM -------------------------------------------------------------------------
Private Function ErrorImage

  ErrorImage = "<IMG SRC=/ztitools/CE.gif width=5 Height=5> "

End Function
REM -------------------------------------------------------------------------
REM Final da Function ErrorImage

REM -------------------------------------------------------------------------
REM Rotina que faz as validacoes no servidor
REM -------------------------------------------------------------------------
Private Function ServerValidates
Dim i, blnIsValid

  blnIsValid = True

  EditBeginReadOperation Session("LastTab")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditCurrentTable, EditLastOperation) Then
          If (Session(EditCurrentField & "RequireMessage") > "") Then
            If Not Session(EditCurrentField & "Value") > "" Then
              Session("Message") = Session("Message") & ErrorImage & Session(EditCurrentField & "RequireMessage") & "<BR>"
              Session(EditCurrentField & "ValidationError") = Session(EditCurrentField & "RequireMessage")
              blnIsValid = False

            End If

          ElseIf Not Session(EditCurrentField & "Null") Then
            If Not Session(EditCurrentField & "Value") > "" Then
              Session("Message") = Session("Message") & ErrorImage & "O campo <B>" &  EditCaption(False) & "</B> deve ser preenchido</B><BR>"
              Session(EditCurrentField & "ValidationError") = "O campo " &  EditCaption(False) & " deve ser preenchido"
              blnIsValid = False

            End If
          End If

        End If
  REM    Response.Write EditLastOperation

    End Select

  Loop

  EditBeginReadOperation Session("LastTab")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditCurrentTable, EditLastOperation) Then
          If Not ValidField(EditCurrentField) Then
            blnIsValid = False

          End If
        End If

    End Select

  Loop

  ServerValidates = blnIsValid

End Function
REM -------------------------------------------------------------------------
REM Final da Function ServerValidates

REM -------------------------------------------------------------------------
REM Rotina que faz a validacao de um campo no servidor
REM -------------------------------------------------------------------------
Private Function ValidField(ByVal strField)
Dim blnIsValid
  blnIsValid = True

  Select Case Session(strField & "ValidationOption")
    Case conValidCGC
      If Not IsCGC(Session(strField & "Value")) Then
        blnIsValid = False

      End If

    Case conValidCPF
      If Not IsCPF(Session(strField & "Value")) Then
        blnIsValid = False

      End If

    Case conValidDate
      If TypeName(Session(strField & "Value")) = "string" Then
        If Not IsValidDate(Session(strField & "Value")) Then
          blnIsValid = False
          
        End If
      End If
        
    Case conValidCompareDates                  
      EditFindField EditCurrentTable, Session(strField & "DoubleValidateField")
      Select Case CompareDate(Session(strField & "Value"), Session(EditCurrentField & "Value"))
        Case 1
          blnIsValid = False
          EditFindField EditCurrentTable, Session(strField & "Field")
          Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
          Session(EditCurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
          
        Case 2
          blnIsValid = False
          Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
          Session(EditCurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
  
        Case 3
          blnIsValid = False
          EditFindField EditCurrentTable, Session(strField & "Field")
          Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
          Session(EditCurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
          EditFindField EditCurrentTable, Session(EditCurrentField & "DoubleValidateField")
          Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
          Session(EditCurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"
         
        Case 4
          blnIsValid = False
          Session("Message") = Session("Message") & ErrorImage & Session(strField & "DoubleValidateMessage")
          Session(EditCurrentField & "ValidationError") = Session(strField & "DoubleValidateMessage")
         
      End Select
    
    Case conValidTime
      If Not IsValidTime(Session(strField & "Value")) Then
        blnIsValid = False

      Else
        Session(strField & "Value") = FormatTime(Session(strField & "Value"))

      End If

    Case conValidSepDate
REM      FormJavaValidSepDate Session(EditCurrentTable & i & "Field") & "DayValue", Session(EditCurrentTable & i & "Field") & "MonthValue", Session(EditCurrentTable & i & "Field") & "YearValue", "V" & Session(EditCurrentTable & i & "Field"), False, ""

    Case conValidDateMToday
      If IsValidDate(Session(strField & "Value")) Then
        If CDate(Session(strField & "Value")) > Now + 1 Then
          blnIsValid = False

        End If

      Else
        blnIsValid = False

      End If

    Case conValidSepDateMToday
REM      FormJavaValidSepDate Session(EditCurrentTable & i & "Field") & "DayValue", Session(EditCurrentTable & i & "Field") & "MonthValue", Session(EditCurrentTable & i & "Field") & "YearValue", "V" & Session(EditCurrentTable & i & "Field"), True, Session(EditCurrentTable & i & "Caption") & " deve ser maior que hoje"

    Case conValidEmail
      If Not IsEmail(Session(strField & "Value")) Then
        blnIsValid = False

      End If

  End Select

  If Not blnIsValid And Session(strField & "ValidationOption") <> conValidCompareDates Then
    EditFindField EditCurrentTable, strField
    Session("Message") = Session("Message") & ErrorImage & "O campo <B>" & EditCaption(False) & "</B> foi preenchido de forma inválida</B><BR>"
    Session(EditCurrentField & "ValidationError") = "O campo " & EditCaption(False) & " foi preenchido de forma inválida"

  End If

  ValidField = blnIsValid

End Function
REM -------------------------------------------------------------------------
REM Final da Function ValidField

REM -------------------------------------------------------------------------
REM Entra com as validacoes dos campos
REM -------------------------------------------------------------------------
Private Sub ClientValidates
Dim i, blnStrExt
  sstrLastValidate = ""  

  sobjShow.HTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"  
  SetPutJSStartEnd False
  blnStrExt = True
  
  EditBeginReadOperation Session("LastTab")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditCurrentTable, EditLastOperation) Then
          If blnStrExt Then
            StrExtBegin
            blnStrExt = False

          End If
          Select Case Session(EditCurrentField & "ValidationOption")
            Case conValidCGC
              FormJavaValidCGC Session(EditCurrentField & "Field"), "frmEditForm", "V" & Session(EditCurrentField & "Field")

            Case conValidCPF
              FormJavaValidCPF Session(EditCurrentField & "Field"), "frmEditForm", "V" & Session(EditCurrentField & "Field")

            Case conValidDate
              FormJavaValidDate Session(EditCurrentField & "Field"), "frmEditForm", "V" & Session(EditCurrentField & "Field"), False, ""

            Case conValidCompareDates
              FormJavaCompareDate Session(EditCurrentField & "Field"), Session(EditCurrentField & "DoubleValidateField"), "frmEditForm", "V" & Session(EditCurrentField & "Field"), Session(EditCurrentField & "DoubleValidateMessage")

            Case conValidTime
REM              FormJavaValidTime Session(EditCurrentField & "Field"), "V" & Session(EditCurrentField & "Field"), False, ""

            Case conValidSepDate
              FormJavaValidSepDate Session(EditCurrentField & "Field") & "DayValue", Session(EditCurrentField & "Field") & "MonthValue", Session(EditCurrentField & "Field") & "YearValue", "frmEditForm", "V" & Session(EditCurrentField & "Field"), False, ""

            Case conValidDateMToday
              FormJavaValidDate Session(EditCurrentField & "Field"), "frmEditForm", "frmEditForm", "V" & Session(EditCurrentField & "Field"), True, EditCaption(False) & " deve ser maior que hoje"

            Case conValidSepDateMToday
              FormJavaValidSepDate Session(EditCurrentField & "Field") & "DayValue", Session(EditCurrentField & "Field") & "MonthValue", Session(EditCurrentField & "Field") & "YearValue", "frmEditForm", "V" & Session(EditCurrentField & "Field"), True, EditCaption(False) & " deve ser maior que hoje"

            Case conValidEmail
              FormJavaValidEmail Session(EditCurrentField & "Field"), "frmEditForm", "V" & Session(EditCurrentField & "Field")

          End Select
          If Session(EditCurrentField & "ValidationOption") <> conValidNone Then
            If sstrLastValidate = "" Then
              sstrLastValidate = sstrLastValidate & "V" & Session(EditCurrentField & "Field") & "()"

            else
              sstrLastValidate = sstrLastValidate & " && V" & Session(EditCurrentField & "Field") & "()"

            End If
          End If
        End If

    End Select

  Loop

  sobjShow.HTMLCR "function ValidateForm() {"
  Session("FormJavaForm") = "frmEditForm"

  EditBeginReadOperation Session("LastTab")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditCurrentTable, EditLastOperation) Then
          If Session(EditCurrentField & "RequireMessage") > "" Then
            FormJavaRequired Session(EditCurrentField & "Field"), Session(EditCurrentField & "RequireMessage")

          ElseIf Not Session(EditCurrentField & "Null") Then
            FormJavaRequired Session(EditCurrentField & "Field"), "O campo " &  EditCaption(False) & " deve ser preenchido"

          End If
        End If

    End Select

  Loop
  If sstrLastValidate > "" Then
    sobjShow.HTMLCR "return (" & sstrLastValidate & ");"

  Else
    sobjShow.HTMLCR "return true;"

  End If

  Response.Write "}"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ClientValidates

%>