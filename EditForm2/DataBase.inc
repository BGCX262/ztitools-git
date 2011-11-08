<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/DataBase.inc
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
REM Rotina que le os dados e tamanhos dos campos ja adicionados na tabela
REM -------------------------------------------------------------------------
Public Function EditReadFieldValues(blnEdit)
Dim sql, i, blnFirst
REM   DebugMode

  Application.Lock
  Application(EditCurrentTable & EditMakeTarget) = ""
  Application.Unlock
  If (Not Application(EditCurrentTable & EditMakeTarget) > "") Or (Not blnEdit) Then
    blnFirst = True
    sql = "SELECT "

    For i = 1 to Session(EditCurrentTable & "FieldNumber")
    	DebugMessage Session(EditCurrentTable & i & "Field"), 2
      If Session(EditCurrentTable & i & "IsKey") Or InStr(LCase(sparEditFind), LCase(Session(EditCurrentTable & i & "Field"))) > 0 Then
        If Not blnFirst Then
          sql = sql & ", "

        Else
          blnFirst = False

        End If
        sql = sql & Session(EditCurrentTable & i & "Field")

      End If
    Next

    sql = sql & " FROM " & EditCurrentTable & Filter

    sobjConn.Close
  Dim strProvider
    strProvider = sobjConn.Provider
    sobjConn.Provider = "SQLOLEDB.1"
    sobjConn.Open

    sobjCMD.ActiveConnection = sobjConn
    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    sobjRS.CursorLocation = adUseClient

    sobjRS.Open sobjCMD, , adOpenStatic, adLockReadOnly  

    If Not sobjRS.EOF And Not sobjRS.BOF Then
      sobjRS.MoveLast
      sobjRS.MoveFirst
      If sparEditFind > "" Then
        sobjRS.Find sparEditFind
        If sobjRS.EOF Or sobjRS.BOF Then
          sobjRS.MoveFirst

        End If  

      Else
        If (sparEditMove <> conMoveFirst) And (sparEditMove > "") Then
          If (sparEditMove = conMoveLast) Then
            sobjRS.MoveLast

          Else                            
            sobjRS.Move CInt(sparEditMove)

          End If
        End If
      End If

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "IsKey") Then
          Session(EditCurrentTable & i & "Value")    = sobjRS(Session(EditCurrentTable & i & "Field"))

        End If
      Next

      Session("AbsolutePosition") = sobjRS.AbsolutePosition
      Session("RecordCount")		          = sobjRS.RecordCount
      sobjRS.Close

      SetPutAnd False

      sql = "SELECT "

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If i > 1 Then
          sql = sql & ", "

        End If

        sql = sql & EditTableOf(EditCurrentTable, i) & "." & Session(EditCurrentTable & i & "Field")

      Next

      sql = sql & " FROM " & EditCurrentTable & Session(EditCurrentTable & "JoinFormUnit")

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "IsKey") Then
          sql = sql & PutAnd & EditTableOf(EditCurrentTable, i) & "." & Session(EditCurrentTable & i & "Field") & "=" & FormatForSql(Session(EditCurrentTable & i & "Value"), Session(EditCurrentTable & i & "Type"))

        End If
      Next

      sobjConn.Close
      sobjConn.Provider = strProvider
      sobjConn.Open

      sobjCMD.ActiveConnection = sobjConn
      sobjCMD.CommandText = sql
      sobjCMD.CommandType = adCmdText

      sobjRS.CursorLocation = adUseServer

      sobjRS.Open sobjCMD, , adOpenStatic, adLockReadOnly

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        DebugMessage Session(EditCurrentTable & i & "Field"), 2
        If Session(EditCurrentTable & i & "Type") = "text" Then
          Session(EditCurrentTable & i & "OldValue") = ToString(sobjRS(Session(EditCurrentTable & i & "Field")))
          Session(EditCurrentTable & i & "Value")    = ToString(sobjRS(Session(EditCurrentTable & i & "Field")))

        Else
          Session(EditCurrentTable & i & "OldValue") = sobjRS(Session(EditCurrentTable & i & "Field"))
          Session(EditCurrentTable & i & "Value")    = sobjRS(Session(EditCurrentTable & i & "Field"))

        End If
      Next
      sobjRS.Close

    Else
      sobjRS.Close
      If Not blnEdit Then
        If Not sparEditFilter > "" Then
          Session("RecordAction") = "1"
          sparEditOption = conOptionAdd
          EditClear sparEditOption

        Else
          If Session(EditCurrentTable & "ClientValidation") Then
            Session("Message") = "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""Registro não localizado"");</SCRIPT>"

          End If
          Session("Message") = Session("Message") & "<FONT COLOR=RED>" & "<B>Registro não localizado</B></FONT>"
          sparEditFilter     = ""
          EditReadFieldValues = EditReadFieldValues(blnEdit)

        End If

      Else
        Session("Message") = "Este registro foi excluído, não podendo ser mais alterado"
        EditReadFieldValues = False

      End If

    End If
    EditReadFieldValues = True

  Else
    Session("Message") = "Este registro está sendo alterado por outro usuário neste momento"
    EditReadFieldValues = False

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditReadFieldValues

REM -------------------------------------------------------------------------
REM Limpa as variaveis da tabela corrente
REM -------------------------------------------------------------------------
Private Sub EditClear(intClearOption)
Dim sql, i, intPos, intPos2, strAux
  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    Select Case Left(Session(EditCurrentTable & i & "DefaultValue"), 8)
      Case constrProximo
        Dim strTable, strField

        If Len(Session(EditCurrentTable & i & "DefaultValue")) > 8 Then
          strTable = Mid(Session(EditCurrentTable & i & "DefaultValue"), 9, InStr(Session(EditCurrentTable & i & "DefaultValue"), ".") - 9)
          strField = Mid(Session(EditCurrentTable & i & "DefaultValue"), InStr(Session(EditCurrentTable & i & "DefaultValue"), ".") + 1)

        ElseIf Session(EditCurrentTable & i & "ExternalAddress") > "" Then
          strField = Session(EditCurrentTable & i & "Field")
          strTable = Session(EditCurrentTable & i & "OwnerTable")

        Else
          strField = Session(EditCurrentTable & i & "Field")
          strTable = EditCurrentTable

        End If

        sql = "SELECT " & strField
        sql = sql & " FROM " & strTable
        sql = sql & " ORDER BY " & strField

        sobjCMD.CommandText = sql
        sobjCMD.CommandType = adCmdText

        sobjRS.Open sobjCMD, , adOpenStatic, adLockReadOnly

        If Not sobjRS.EOF And Not sobjRS.BOF Then
          sobjRS.MoveLast
          Dim intCodigo
          Session(EditCurrentTable & i & "OldValue") = CLng(sobjRS(strField)) + 1
          Do While Application(EditCurrentTable & EditMakeTarget) > ""
            Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "OldValue") + 1
          
          Loop
          Session(EditCurrentTable & i & "Value")    = Session(EditCurrentTable & i & "OldValue")
          Application.Lock
          Application(EditCurrentTable & EditMakeTarget) = Now()
          Application.Unlock

        Else
          Session(EditCurrentTable & i & "Value") = 1
          Session(EditCurrentTable & i & "OldValue") = 1

        End If
        sobjRS.Close

      Case Else
        If intClearOption = conOptionAdd Then
          If sparEditDefaults > "" Then                   
            intPos = InStr(sparEditDefaults, LCase(Session(EditCurrentTable & i & "Field")) & "=")
            If intPos > 0 Then
REM              sobjShow.MessageError Session(EditCurrentTable & i & "Field")
              intPos = intPos + Len(Session(EditCurrentTable & i & "Field"))
              strAux = Right(sparEditDefaults, Len(sparEditDefaults) - intPos)
              intPos2 = InStr(strAux, ";")
              If intPos2 > 0 Then
                Session(EditCurrentTable & i & "Value") = EditFormatField(Left(strAux, intPos2 - 1), Session(EditCurrentTable & i & "Type"), Session(EditCurrentTable & i & "Size"), Session(EditCurrentTable & i & "CharCase"))

              Else
                Session(EditCurrentTable & i & "Value") = EditFormatField(strAux, Session(EditCurrentTable & i & "Type"), Session(EditCurrentTable & i & "Size"), Session(EditCurrentTable & i & "CharCase"))

              End If

            Else
              Session(EditCurrentTable & i & "Value") = Session(EditCurrentTable & _
                                                                  i & "DefaultValue")
            End If

          Else
            If Session(EditCurrentTable & i & "EditOption") = conAtualizationField Then
              Session(EditCurrentTable & i & "Value") = Now

            Else
              Session(EditCurrentTable & i & "Value") = Session(EditCurrentTable & _
                                                                i & "DefaultValue")

            End If
          End If
        End If

    End Select
  Next

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditClear

REM -------------------------------------------------------------------------
REM Armazena os campos da tabela do form em variaveis
REM -------------------------------------------------------------------------
Private Sub EditGetForm(strFields)
Dim i

  EditBeginReadOperation strFields
  Dim Width

  Width = 0
  Do While EditIsThereOperation
    Select Case EditNextOperation
       Case ",", ";", "-", ""
       Case "="
         EditIsThereOperation

       Case Else
         If EditFindField(EditCurrentTable, EditLastOperation) Then
           Select Case Session(EditCurrentField & "EditOption")
             Case conAtualizationField
               Session(EditCurrentField & "Value") = ZTILongDate(Now)

             Case conSeparateDateField
                Session(EditCurrentField & "Value") = ImportDate(Request.Form(Session(EditCurrentField & "Field") & _
                                                "DayValue") & Request.Form(Session(EditCurrentField & "Field") & _
                                                "MonthValue") & Request.Form(Session(EditCurrentField & "Field") &_
                                                "YearValue"), conDateTypeDDMMAAAA)

              Case conPasswordField, conFileField
                If Request.Form(Session(EditCurrentField & "Field")) > "" Then
                  If LCase(Request.Form(Session(EditCurrentField & "Field"))) = "nada" Then
                    Session(EditCurrentField & "Value") = ""

                  Else
                    Session(EditCurrentField & "Value") = EditFormatField(Request.Form(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"), Session(EditCurrentField & "Size"), Session(EditCurrentField & "CharCase"))

                  End If
                End If

              Case conImageField
REM                SendImage
REM                sobjShow.JSAlert "Teste"

              Case conCheckField
                If LCase(Request.Form(Session(EditCurrentField & "Field"))) = LCase(CStr(Session(EditCurrentField & ",1Value"))) Then
                  Session(EditCurrentField & "Value") = Session(EditCurrentField & ",1Value")

                Else
                  Session(EditCurrentField & "Value") = Session(EditCurrentField & ",2Value")

                End If

              Case Else
                Select Case Session(EditCurrentField & "Type")
                  Case conDataFloat, conDataReal, conDataMoney, conDataDateTime
REM                    sobjShow.JSAlert Request.Form(Session(EditCurrentField & "Field")) & " = " & EditFormatField(Request.Form(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"), Session(EditCurrentField & "Size"), Session(EditCurrentField & "CharCase")) & "" & TypeName(Session(EditCurrentField & "Value"))
                    Session(EditCurrentField & "Value") = EditFormatField(Request.Form(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"), Session(EditCurrentField & "Size"), Session(EditCurrentField & "CharCase"))
REM                    sobjShow.JSAlert TypeName(Session(EditCurrentField & "Value"))

                  Case Else
                    Session(EditCurrentField & "Value") = Request.Form(Session(EditCurrentField & "Field"))

                End Select
           End Select
         End If

     End Select

  Loop

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ClienteGetForm

REM -------------------------------------------------------------------------
REM Grava as variaveis de Cliente no RS
REM -------------------------------------------------------------------------
Private Sub EditSave
Dim i, j, strAnotherTable, strSaved, strCurrentTable, strTable

REM  DebugMode

  If Session("RecordAction") = "1" Then
    sobjCMD.CommandText = EditCurrentTable
    sobjCMD.CommandType = adCmdTable

    sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

    sobjRS.AddNew

    strAnotherTable = ""
    For i = 1 to Session(EditCurrentTable & "FieldNumber")
      If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
        If Not Session(EditCurrentTable & i & "ExternalAddress") > "" Then
          DebugMessage Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "Value"), 0
          sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatFieldtoSave(Session(EditCurrentTable & _
                                                                   i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                   Session(EditCurrentTable & i & "Size"),  _
                                                                   Session(EditCurrentTable & i & "CharCase"))
          Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")
REM          sobjShow.JSAlert Session(EditCurrentTable & i & "Field")

        Else
          If Not strAnotherTable > "" Then
            strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

          End If
        End If
      End If

    Next
    sobjRS.Update
    sobjRS.Close

    strSaved = EditCurrentTable & "&"

    Do While strAnotherTable <> ""
      strCurrentTable = strAnotherTable
      strSaved = strSaved & "(" & strCurrentTable & ")"
      strAnotherTable = ""

      sobjCMD.CommandText = strCurrentTable
      sobjCMD.CommandType = adCmdTable

      sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

      sobjRS.AddNew

      strTable = ""
      If EditFindRelation(EditCurrentTable, strCurrentTable) Then
         strTable = EditCurrentTable & sintCurrentRelation

      ElseIf EditFindRelation(strCurrentTable, EditCurrentTable) Then
         strTable = strCurrentTable & sintCurrentRelation

      End If

      If strTable > "" Then
        For j = 1 to Session(strTable & "FieldtoFieldNumber")
          If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
            If EditFindField(EditCurrentTable, Session(strTable & "," & j & "RelationField1")) Then
              sobjRS(Session(strTable & "," & j & "RelationField2")) = EditFormatFieldtoSave(Session(EditCurrentField & "Value"), _
                                                                       Session(EditCurrentField & "Type"), EditSize, _
                                                                       Session(EditCurrentTable & i & "CharCase"))
              Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

            End If
          End If
        Next
      End If

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
          If Session(EditCurrentTable & i & "OwnerTable") = strCurrentTable Then
  REM           sobjShow.JSAlert Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "Value")
            sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatFieldtoSave(Session(EditCurrentTable & _
                                                                     i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                     Session(EditCurrentTable & i & "Size"),  _
                                                                     Session(EditCurrentTable & i & "CharCase"))
            Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

          Else
            If (Not strAnotherTable > "") And (InStr(strSaved, Session(EditCurrentTable & i & "OwnerTable")) = 0) Then
              strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

            End If
          End If
        End If
      Next

      sobjRS.Update
      sobjRS.Close

    Loop

  Else
    Dim sql

    sql = "SELECT * FROM " & EditCurrentTable

    If sparEditFind > "" Then
      sql = sql & " WHERE " & sparEditFind

    End If

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

REM    DebugMessage sql, 1

    sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

    strAnotherTable = ""
    For i = 1 to Session(EditCurrentTable & "FieldNumber")
      If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
        If Not Session(EditCurrentTable & i & "ExternalAddress") > "" Then
REM          sobjShow.JSAlert Session(EditCurrentTable & i & "Field") & " = " & Session(EditCurrentTable & i & "Value") & " = " & EditFormatFieldtoSave(Session(EditCurrentTable & i & "Value"), Session(EditCurrentTable & i & "Type"), Session(EditCurrentTable & i & "Size"), Session(EditCurrentTable & i & "CharCase")) & " " & TypeName(Session(EditCurrentTable & i & "Value")) & " " & Session(EditCurrentTable & i & "ShowOption")
          sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatFieldtoSave(Session(EditCurrentTable & _
                                                                   i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                   Session(EditCurrentTable & i & "Size"), _
                                                                   Session(EditCurrentTable & i & "CharCase"))
          Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

        Else
          If Not strAnotherTable > "" Then
            strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

          End If
        End If
      End If
    Next
    sobjRS.Update
    sobjRS.Close

    strSaved = EditCurrentTable & "&"

    Do While strAnotherTable <> ""
      strCurrentTable = strAnotherTable
      strSaved = strSaved & "(" & strCurrentTable & ")"
      strAnotherTable = ""

      strTable = ""
      If EditFindRelation(EditCurrentTable, strCurrentTable) Then
         strTable = EditCurrentTable & sintCurrentRelation

      End If

      sql = "SELECT * FROM " & strCurrentTable

      SetPutAnd False

      If strTable > "" Then
        For j = 1 to Session(strTable & "FieldtoFieldNumber")
          If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
            If EditFindField(EditCurrentTable, Session(strTable & "," & j & "RelationField1")) Then
              sql = sql & PutAnd & Session(strTable & "," & j & "RelationField2") & " = " & _
                        FormatForSQL(Session(EditCurrentField & "OldValue"), Session(EditCurrentField & "Type"))

            End If
          End If
        Next

      End If

REM     DebugMessage sql, 1
      sobjCMD.CommandText = sql
      sobjCMD.CommandType = adCmdText

      sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

      If sobjRS.EOF Or sobjRS.BOF Then
        sobjCMD.CommandText = strCurrentTable
        sobjCMD.CommandType = adCmdTable

        sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

        sobjRS.AddNew

      End If

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
          If Session(EditCurrentTable & i & "OwnerTable") = strCurrentTable Then
              sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatFieldtoSave(Session(EditCurrentTable & _
                                                                       i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                       Session(EditCurrentTable & i & "Size"), _
                                                                       Session(EditCurrentTable & i & "CharCase"))
              Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

          Else
            If (Not strAnotherTable > "") And (InStr(strSaved, Session(EditCurrentTable & i & "OwnerTable")) = 0) Then
              strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

            End If
          End If
        End If
      Next

      If strTable > "" Then
        For j = 1 to Session(strTable & "FieldtoFieldNumber")
          If Session(EditCurrentTable & i & "ShowOption") <> 1 Then
            If EditFindField(EditCurrentTable, Session(strTable & "," & j & "RelationField1")) Then
              sobjRS(Session(strTable & "," & j & "RelationField2")) = EditFormatFieldtoSave(Session(EditCurrentField & "Value"), _
                                                                       Session(EditCurrentField & "Type"), EditSize, _
                                                                       Session(EditCurrentTable & i & "CharCase"))
              Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

            End If
          End If
        Next

      End If

      sobjRS.Update
      sobjRS.Close

    Loop

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ClienteSaveRS

REM -------------------------------------------------------------------------
REM Rotina que adiciona um target no link interno de um campo
REM -------------------------------------------------------------------------
Public Sub EditDelete
  DeleteFrom EditCurrentTable, "", "", "", ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditDelete

REM -------------------------------------------------------------------------
REM Rotina que adiciona um target no link interno de um campo
REM -------------------------------------------------------------------------
Public Sub DeleteFrom(ByVal strTable, ByVal strLastTable, ByVal strSelect, ByVal strLinkField, ByVal strLinkField2)
Dim strSelectAux, i, j, sql
REM  DebugMode

  For i = 1 to Session(strTable & "RelationNumber")
    strSelectAux = strSelect

    If Not strSelectAux > "" Then
        strSelectAux = strTable & " WHERE "

    End if

    For j = 1 to Session(strTable & i & "FieldtoFieldNumber")
      If strSelectAux > (strTable & " WHERE ") Then
        strSelectAux = strSelectAux & " AND "

      End if
      strSelectAux = strSelectAux & Session(strTable & i & "RelationTable") & "." & Session(strTable & i & "," & j & "RelationField2") & " = " & strTable & "." & Session(strTable & i & "," & j & "RelationField1")

    Next

    If Not strSelect > "" Then
      If sparEditFind > "" Then
        strSelectAux = strSelectAux & " AND " & sparEditFind

      End If

    End If

    strSelectAux = Session(strTable & i & "RelationTable") & ", " & strSelectAux

    DeleteFrom Session(strTable & i & "RelationTable"), strTable, strSelectAux, Session(strTable & i & ",1RelationField1"), Session(strTable & i & ",1RelationField2")

  Next

  If strSelect > "" Then
    sql = "DELETE FROM " & strTable & " WHERE " & strTable & "." & strLinkField2
    sql = sql & " IN (SELECT " & strLastTable & "." & strLinkField & " FROM " & strSelect & ")"

REM    DebugMessage sql, 1

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    sobjRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

  Else
    sql = "DELETE FROM " & strTable

    If sparEditFind > "" Then
      sql = sql & " WHERE " & sparEditFind

    End If

REM    DebugMessage sql & "<BR>", 1

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    sobjRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditDelete

%>