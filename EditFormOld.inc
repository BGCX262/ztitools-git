<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 10:00h 16/12/1998
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

  Const conDataFloat    = "float"
  Const conDataVarChar  = "varchar"
  Const conDataInt      = "int"
  Const conDataDateTime = "datetime"
  Const conDataMoney    = "money"

  Const conDataNull     = True
  Const conDataNotNull  = False

  Const conOptionCascade      = 0
  Const conOptionRestricted   = 1
  Const conOptionNullifies    = 2

  Const conClientValidation   = True
  Const conServerValidation   = False

  Const conMoveFirst          = "0"
  Const conMovePrevious       = "1"
  Const conMoveNext           = "2"
  Const conMoveLast           = "3"

  Const conTextField          = 0
  Const conUFField            = 1
  Const conLookupField        = 2
  Const conSeparateDateField  = 3
  Const conCheckField         = 4
  Const conTextAreaField      = 5
  Const conRadioField         = 6
  Const conSelectField        = 7
  Const conPasswordField      = 8
  Const conHTTPField          = 9
  Const conEMailField         = 10
  Const conImageField         = 11
  Const conImageFileField     = 12
  Const conFileField          = 13

  Const conValidNone          = 0
  Const conValidCGC           = 1
  Const conValidCPF           = 2
  Const conValidDate          = 3
  Const conValidSepDate       = 4
  Const conValidDateMToday    = 5
  Const conValidSepDateMToday = 6
  Const conValidEmail         = 7

  Const conPEditOption        = "EO"
  Const conPEditWhat          = "EW"
  Const conPEditMove          = "EM"
  Const conPEditEditable      = "EE"
  Const conPEditEditableStr   = "&EE=1"
  Const conPEditOrderField    = "EOF"
  Const conPEditOrderDesc     = "EOD"
  Const conPEditOrderDescStr  = "&EOD=1"
  Const conPEditFind          = "EF"
  Const conPEditFilter        = "ET"
  Const conPEditList          = "EL"
  Const conPEditDefaults      = "ED"

  Const conOptionFirstPage    = "h00"
  Const conOptionExeFirstPage = "h01"
  Const conOptionAZEdit       = "h02"
  Const conOptionFPEdit       = "h03"
  Const conOptionEditor       = "h04"
  Const conOptionExeEdition   = "h05"
  Const conOptionAdd          = "h06"
  Const conOptionCopy         = "h07"
  Const conOptionEdit         = "h08"
  Const conOptionDelete       = "h09"
  Const conOptionSave         = "h10"
  Const conOptionSaveAdd      = "h11"
  Const conOptionSaveCopy     = "h12"

  Const conOptionShowGrid     = "h13"

  Const conOptionFindCommon   = "h14"
  Const conOptionFindAdvanced = "h15"
  Const conOptionExeFind      = "h16"
  Const conOptionSummary      = "h17"

  Const constrProximo = "proximo "

  Const constrAdd     = "Incluir"
  Const constrEdit    = "Alterar"
  Const constrDelete  = "Excluir"
  Const constrSave    = "Salvar"
  Const constrFind    = "Localizar"
  Const constrList    = "Listagem"
  Const constrRefresh = "Atualizar"
  Const constrCopy    = "Copiar"
  Const constrSaveAdd = "Salvar e Incluir"
  Const constrSaveCopy= "Salvar e Copiar"
  Const constrFirst   = " |<  "
  Const constrPrevious= "  <  "
  Const constrNext    = "  >  "
  Const constrLast    = "  >| "

  Const constrOK     = "   OK   "
  Const constrCancel = "Cancelar"
  Const constrYes    = "   Sim   "
  Const constrNo     = "   Não   "

  Public sparEditOption
  Public sparEditWhat
  Public sparEditMove
  Public sparEditEditable
  Public sparEditFind
  Public sparEditFilter
  Public sparEditList
  Public sparEditDefaults

  Public sintCurrentTableNum
  Public sintCurrentRelation
  Public sstrSubText
  Public sstrGroupFieldLastValue
  Public sstrGroupField
  Public sstrFields
  Public sstrFields2
  Public sblnHaveSum
  Public sblnHaveGraph
  Public sintAbsolutePosition
  Public sintRecordCount
  Public sintLastNumber

Dim sstrText, sstrText2, sstrText3, sstrLastValidate, sintSize, sintNumCol, sstrCheckAsRadio

  Public sstrCreateURL, sstrCreateURLMode, sstrCreateURLTarget

  Session("TableNumber") = 0
  sintSize = 0

REM -------------------------------------------------------------------------
REM Entra Campo File para download de arquivos
REM -------------------------------------------------------------------------
Sub FormInputFile(ByVal strName, ByVal intSize, ByVal strValue, ByVal intMaxLength)
  
  FormInput "File", strName, intSize, strValue, intMaxLength, "", 0, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputText

REM -------------------------------------------------------------------------
REM Testa se a opcao atual é do edit form
REM -------------------------------------------------------------------------
Private Function IsEditFormOption(ByVal strOption)
  IsEditFormOption = (Left(strOption, 1) = "h")

End Function
REM -------------------------------------------------------------------------
REM Final da Function IsEditFormOption

REM -------------------------------------------------------------------------
REM Rotina que le as query strings
REM -------------------------------------------------------------------------
Public Sub EditQueryString

  sparEditOption     = lCase(Request.QueryString(conPEditOption))
  sparEditWhat       = lCase(Request.QueryString(conPEditWhat))
  sparEditMove       = lCase(Request.QueryString(conPEditMove))
  sparEditEditable   = lCase(Request.QueryString(conPEditEditable))
  sparEditFind       = LCase(Request.QueryString(conPEditFind))
  sparEditFilter     = LCase(Request.QueryString(conPEditFilter))
  sparEditList       = LCase(Request.QueryString(conPEditList))
  sparEditDefaults   = Request.QueryString(conPEditDefaults)

  Session("LastEditOption") = Session("EditOption")
  Session("EditOption")     = sparEditOption

  If Not sparEditOption > "" Then
    sparEditOption   = conOptionFirstPage

  End If
  If Not sparEditEditable > "" Then
    sparEditEditable = 0

  End If

  Session("OrderField") = lCase(Request.QueryString(conPEditOrderField))
  Session("OrderDesc")  = lCase(Request.QueryString(conPEditOrderDesc))

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditQueryString

Private Sub ClearSession(ByVal strSession)
  If Session(strSession) > "" Then
    Session(strSession) = Null

  End If

End Sub

Private Sub ClearAllSessions
Dim strTable, i ,j, k

  If Session("TableNumber") > "" Then
    For i = 1 to Session("TableNumber")
      strTable = Session("Table" & i)
      If Session(strTable & "FieldNumber") > "" Then
        For j = 1 to Session(strTable & "FieldNumber")
          ClearSession strTable & j & "Field"
          ClearSession strTable & j & "Size"
          ClearSession strTable & j & "Type"
          ClearSession strTable & j & "Null"
          ClearSession strTable & j & "Caption"
          ClearSession strTable & j & "EditOption"
          ClearSession strTable & j & "ValidationOption"
          ClearSession strTable & j & "ShowOption"
          ClearSession strTable & j & "DefaultValue"
          ClearSession strTable & j & "RequireMessage"
          ClearSession strTable & j & "Hint"
          ClearSession strTable & j & "ExternalLink"
          ClearSession strTable & j & "InternalLink"
          ClearSession strTable & j & "ShowSize"
          ClearSession strTable & j & "BeginYear"
          ClearSession strTable & j & "EndYear"
          ClearSession strTable & j & "LookupTable"
          ClearSession strTable & j & "LookupField"
          ClearSession strTable & j & "LookupFieldText"
          ClearSession strTable & j & "ListSize"
          ClearSession strTable & j & "ListChars"
          ClearSession strTable & j & "RadioColumns"
          ClearSession strTable & j & "ExternalAddress"
          ClearSession strTable & j & "OwnerTable"
          ClearSession strTable & j & "IsKey"
          ClearSession strTable & j & "Value"
          ClearSession strTable & j & "OldValue"
          ClearSession strTable & j & "FieldtoFieldNumber"
          If Session(strTable & j & "ValueNumber") > "" Then
            For k = 1 to Session(strTable & j & "ValueNumber")
              ClearSession strTable & j & "," & k & "Value"
              ClearSession strTable & j & "," & k & "Caption"

            Next
            ClearSession strTable & j & "ValueNumber"
          End If
          If Session(strTable & j & "ILTargetNumber") > "" Then
            For k = 1 to Session(strTable & j & "ILTargetNumber")
              ClearSession strTable & j & "," & k & "ILTarget"

            Next
            ClearSession strTable & j & "ILTargetNumber"
          End If

        Next
        ClearSession strTable & "FieldNumber"

      End If

      If Session(strTable & "RelationNumber") > "" Then
        For j = 1 to Session(strTable & "RelationNumber")
          ClearSession strTable & j & "RelationTable"
          ClearSession strTable & j & "DeleteOption"
          ClearSession strTable & j & "UpdateOption"
          If Session(strTable & "FieldtoFieldNumber") > "" Then
            For k = 1 to Session(strTable & "FieldtoFieldNumber")
              ClearSession strTable & j & "," & k & "RelationField1"
              ClearSession strTable & j & "," & k & "RelationField2"

            Next
            ClearSession strTable & j & "FieldtoFieldNumber"
            
          End If

        Next
        ClearSession strTable & "RelationNumber"
      End If

      If Session(strTable & "IndexNumber") > "" Then
        For j = 1 to Session(strTable & "IndexNumber")
          ClearSession strTable & j & "IndexName"
          ClearSession strTable & j & "IndexContent"

        Next
        ClearSession strTable & "IndexNumber"
      End If

      If Session(strTable & "HeaderNumber") > "" Then
        For j = 1 to Session(strTable & "HeaderNumber")
          ClearSession strTable & j & "HeaderText"
          ClearSession strTable & j & "HeaderBegin"
          ClearSession strTable & j & "HeaderWidth"

        Next
        ClearSession strTable & "HeaderNumber"
      End If

      If Session(strTable & "ListNumber") > "" Then
        For j = 1 to Session(strTable & "ListNumber")
          ClearSession strTable & j & "ListName"
          ClearSession strTable & j & "List"
          ClearSession strTable & j & "ListCaption"

        Next
        ClearSession strTable & "ListNumber"
      End If

      If Session(strTable & "OrderNumber") > "" Then
        For j = 1 to Session(strTable & "OrderNumber")
          ClearSession strTable & j & "Order"

        Next
        ClearSession strTable & "OrderNumber"
      End If

      ClearSession "Table" & i

      ClearSession strTable & "UnitOperationNumber"
      ClearSession strTable & "ListOperationNumber"
      ClearSession strTable & "IndexClusteredName"
      ClearSession strTable & "Caption"
      ClearSession strTable & "EditOption"
      ClearSession strTable & "ClientValidation"
      ClearSession strTable & "UnitList"
      ClearSession strTable & "FieldList"
      ClearSession strTable & "FieldUnit"
      ClearSession strTable & "UnitOperations"
      ClearSession strTable & "FieldFind"
      ClearSession strTable & "JoinFormUnit"

    Next
    ClearSession "TableNumber"

  End If

REM  Dim sessitem
REM  For Each sessitem in Session.Contents
REM    Response.write(sessitem & " : " & Session.Contents(sessitem) & "<BR>")
REM  Next

End Sub

REM -------------------------------------------------------------------------
REM Retorna o nome da tabela atual
REM -------------------------------------------------------------------------
Public Function EditCurrentTable

  EditCurrentTable = Session("CurrentTable")

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCurrentTable

REM -------------------------------------------------------------------------
REM Rotina que cria uma nova secao de tabela com seus parametros principais
REM -------------------------------------------------------------------------
Public Sub DataBegin(ByVal strTable)
  Session("TableNumber")                    = Session("TableNumber") + 1

  Session("Table" & Session("TableNumber")) = strTable

  Session("CurrentTable")                          = strTable
  sintCurrentTableNum                = Session("TableNumber")
  Session(strTable & "FieldNumber")         = 0
  Session(strTable & "RelationNumber")      = 0
  Session(strTable & "IndexNumber")         = 0
  Session(strTable & "IndexClusteredName") = ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataBegin

REM -------------------------------------------------------------------------
REM Rotina que adiciona um campo da tabela no Dic. de Dados
REM -------------------------------------------------------------------------
Public Sub DataAddField(ByVal strField, ByVal strType, ByVal intSize, ByVal blnNull)
Dim i

  Session(EditCurrentTable & "FieldNumber") = Session(EditCurrentTable & "FieldNumber") + 1

  i = Session(EditCurrentTable & "FieldNumber")

  Session(EditCurrentTable & i & "Field")  = strField
  Session(EditCurrentTable & i & "Size")   = intSize
  Session(EditCurrentTable & i & "Type")   = strType
  Session(EditCurrentTable & i & "Null")   = blnNull

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddField

REM -------------------------------------------------------------------------
REM Rotina que cria uma nova chave primaria
REM -------------------------------------------------------------------------
Public Sub DataAddPrimaryKey(ByVal strField)
  If EditFindField(strField) Then
    Session(EditCurrentField & "IsKey") = True
REM      ShowMessageError Session(EditCurrentField & "Field")

  Else
    Err.Raise 300, "EditDataAddPrimaryKey", "O campo """ & strField & """, da tabela """ & EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditDataAddPrimaryKey

REM -------------------------------------------------------------------------
REM Rotina que adiciona um relacionamento entre tabelas
REM -------------------------------------------------------------------------
Public Sub DataAddRelation(ByVal strTable, ByVal intDeleteOption, ByVal intUpdateOption)
Dim i

  Session(EditCurrentTable & "RelationNumber")     = Session(EditCurrentTable & "RelationNumber") + 1

  i = Session(EditCurrentTable & "RelationNumber")

  Session(EditCurrentTable & i & "RelationTable")  = strTable
  Session(EditCurrentTable & i & "DeleteOption")   = intDeleteOption
  Session(EditCurrentTable & i & "UpdateOption")   = intUpdateOption
  Session(EditCurrentTable & i & "FieldtoFieldNumber") = 0


End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddRelation

REM -------------------------------------------------------------------------
REM Rotina que adiciona um campo da tabela no Dic. de Dados
REM -------------------------------------------------------------------------
Public Sub DataAddRelationFields(ByVal strField, ByVal strField2)
Dim i, j

  i = Session(EditCurrentTable & "RelationNumber")

  j = Session(EditCurrentTable & i & "FieldtoFieldNumber")

  Session(EditCurrentTable & i & "FieldtoFieldNumber") = j + 1

  j = j + 1

  Session(EditCurrentTable & i & "," & j & "RelationField1") = strField
  Session(EditCurrentTable & i & "," & j & "RelationField2") = strField2

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddRelationFields

REM -------------------------------------------------------------------------
REM Rotina que localiza um campo em qualquer tabela
REM -------------------------------------------------------------------------
Public Function DataFindField(ByVal strField)
Dim i, strCurrentTable

  i = 1

  strCurrentTable = EditCurrentTable

  Do While (Not EditFindField(strField)) And (i < Session("TableNumber"))
    i = i + 1
    DataNextTable

  Loop

REM  Response.Write Session(EditCurrentField & "Field") & " " & Session(EditCurrentField & "Caption")

  DataFindField = (i = Session("TableNumber"))

End Function
REM -------------------------------------------------------------------------
REM Final da Function DataFindField

REM -------------------------------------------------------------------------
REM Rotina que move para a proxima tabela
REM -------------------------------------------------------------------------
Public Sub DataNextTable(ByVal strField)

  If sintCurrentTableNum = Session("TableNumber") Then
    sintCurrentTableNum = 1
    Session("CurrentTable") = Session("Table1")

  Else
    sintCurrentTableNum = sintCurrentTableNum + 1
    Session("CurrentTable") = Session("Table" & sintCurrentTableNum)

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataNextTable

REM -------------------------------------------------------------------------
REM Rotina que adiciona um indice a tabela
REM -------------------------------------------------------------------------
Public Sub DataAddIndex(ByVal strName, ByVal strContent)
Dim i

  Session(EditCurrentTable & "IndexNumber") = Session(EditCurrentTable & "IndexNumber") + 1

  i = Session(EditCurrentTable & "IndexNumber")

  Session(EditCurrentTable & i & "IndexName")    = strName
  Session(EditCurrentTable & i & "IndexContent") = strContent

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddIndex

REM -------------------------------------------------------------------------
REM Rotina que adiciona um indice a tabela Clustered
REM -------------------------------------------------------------------------
Public Sub DataIndexClustered(ByVal strName, ByVal strContent)
  Session(EditCurrentTable & "IndexClusteredName")    = strName
  Session(EditCurrentTable & "IndexClusteredContent") = strContent

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataIndexClustered

REM -------------------------------------------------------------------------
REM Rotina que cria uma nova secao de tabela com seus parametros principais
REM -------------------------------------------------------------------------
Public Sub EditFormBegin(ByVal strTable, ByVal strCaption, ByVal intEditOption, ByVal strWhatList, _
                         ByVal blnClientValidation)
  Session("CurrentTable")                       = strTable
  Session(strTable & "Caption")          = strCaption
  Session(strTable & "EditOption")       = intEditOption
  Session(strTable & "WhatList")         = strWhatList
  Session(strTable & "ClientValidation") = blnClientValidation
  Session(strTable & "ListNumber")       = 0
  Session(strTable & "OrderNumber")      = 0
  Session(strTable & "HeaderNumber")     = 0
  Session(strTable & "UnitList")         = ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormBegin

REM -------------------------------------------------------------------------
REM Rotina que adiciona um campo da tabela na edicao
REM -------------------------------------------------------------------------
Public Sub EditAddFormField(ByVal strField, ByVal strCaption, ByVal intEditOption, ByVal intValidationOption, _
                            ByVal intShowOption, ByVal DefaultValue, ByVal strRequireMessage)
Dim i
  If EditFindField(strField) Or (InStr(strField, ".") <> 0) Then
    If (InStr(strField, ".") <> 0) Then
      Session(EditCurrentTable & "FieldNumber") = Session(EditCurrentTable & "FieldNumber") + 1
      Session("CurrentField")                   = Session(EditCurrentTable & "FieldNumber")
      i                                         = Session("CurrentField")

    End If

    Session(EditCurrentField & "Field")            = strField
    Session(EditCurrentField & "Caption")          = strCaption
    Session(EditCurrentField & "EditOption")       = intEditOption
    Session(EditCurrentField & "ValidationOption") = intValidationOption
    Session(EditCurrentField & "ShowOption")       = intShowOption
    Session(EditCurrentField & "DefaultValue")     = DefaultValue
    Session(EditCurrentField & "RequireMessage")   = strRequireMessage
    Session(EditCurrentField & "ValueNumber")      = 0
    Session(EditCurrentField & "Hint")             = ""
    Session(EditCurrentField & "ExternalLink")     = ""
    Session(EditCurrentField & "InternalLink")     = ""
    Session(EditCurrentField & "ShowSize")         = ""
    Session(EditCurrentField & "BeginYear")        = ""
    Session(EditCurrentField & "EndYear")          = ""
    Session(EditCurrentField & "LookupTable")      = ""
    Session(EditCurrentField & "LookupField")      = ""
    Session(EditCurrentField & "LookupFieldText")  = ""
    Session(EditCurrentField & "ListSize")         = ""
    Session(EditCurrentField & "ListChars")        = ""
    Session(EditCurrentField & "RadioColumns")     = 1
    If intEditOption = conTextAreaField Then
      Session(EditCurrentField & "TextAreaHeight")   = 4

    End If
    If (InStr(strField, ".") <> 0) Then
      Dim strCurrentField, strCurrentTable

      strCurrentField  = EditCurrentField
      strCurrentTable  = EditCurrentTable

      Session("CurrentTable") = Left(strField, InStr(strField, ".") - 1)

      If EditFindField(Mid(strField, InStr(strField, ".") + 1)) Then
        Session(strCurrentTable & i & "ExternalAddress") = EditCurrentField
        Session(strCurrentTable & i & "OwnerTable")     = Left(strField, InStr(strField, ".") - 1)
        Session(strCurrentTable & i & "Field")          = Mid(strField, InStr(strField, ".") + 1)
        REM ShowMessageError strCurrentTable & i & "Field" & "=" & Mid(strField, InStr(strField, ".") + 1)
        Session(strCurrentTable & i & "Size")           = Session(EditCurrentField & "Size")
        Session(strCurrentTable & i & "Null")           = Session(EditCurrentField & "Null")
        Session(strCurrentTable & i & "Type")           = Session(EditCurrentField & "Type")

      Else
        Err.Raise 300, "EditAddFormField", "O campo """ & Mid(strField, InStr(strField, ".") + 1) & """, da tabela """ & _
              Left(strField, InStr(strField, ".") - 1) & """, não existe"

      End If
      Session("CurrentField") = i
      Session("CurrentTable") = strCurrentTable
      REM ShowMessageError strCurrentTable & i & "=" & EditCurrentField & "=>" & Session(EditCurrentField & "Field")

    End If

  Else
    Err.Raise 300, "EditAddFormField", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddFormField

REM -------------------------------------------------------------------------
REM Rotina que testa se um campo e externo a uma tabela
REM -------------------------------------------------------------------------
Public Function IsExternalField

  IsExternalField = (Session(EditCurrentField & "ExternalAddress") > "")

End Function
REM -------------------------------------------------------------------------
REM Final da Function IsExternalField

REM -------------------------------------------------------------------------
REM Rotina que altera a altura de um text area
REM -------------------------------------------------------------------------
Public Function EditFieldTextAreaHeight(strField, intHeight)
  If EditFindField(strField) Then
    Session(EditCurrentField & "TextAreaHeight") = intHeight

  Else
    Err.Raise 300, "EditFormTextAreaHeight", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormTextAreaHeight

REM -------------------------------------------------------------------------
REM Rotina que altera o numero de colunas de um campo radio
REM -------------------------------------------------------------------------
Public Function EditFieldRadioColumns(strField, intColumns)
  If EditFindField(strField) Then
    Session(EditCurrentField & "RadioColumns") = intColumns

  Else
    Err.Raise 300, "EditFormRadioColumns", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormRadioColumns

REM -------------------------------------------------------------------------
REM Rotina que altera o tamanho do campo na apresentacao de listagens
REM -------------------------------------------------------------------------
Public Function EditFieldListSize(strField, intSize)
  If EditFindField(strField) Then
    Session(EditCurrentField & "ListSize") = intSize

  Else
    Err.Raise 300, "EditFieldListSize", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldListSize

REM -------------------------------------------------------------------------
REM Rotina que altera o numero de caracteres do campo na apresentacao de listagens
REM -------------------------------------------------------------------------
Public Function EditFieldListChars(strField, intSize)
  If EditFindField(strField) Then
    Session(EditCurrentField & "ListChars") = intSize

  Else
    Err.Raise 300, "EditFieldListChars", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldListChars

REM -------------------------------------------------------------------------
REM Rotina que altera o tamanho do campo na apresentacao de listagens
REM -------------------------------------------------------------------------
Public Function BreakString(strValue, intSize)
  If IntSize > "" Then
    If Len(strValue) > intSize Then
      BreakString = Left(strValue, intSize) & "..."

    Else
      BreakString = strValue

    End If

  Else
    BreakString = strValue

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldListSize

REM -------------------------------------------------------------------------
REM Rotina que altera o numero de colunas de um campo radio
REM -------------------------------------------------------------------------
Public Function EditFormFieldMask(strField, strMask, strCompleter)
  If EditFindField(strField) Then
    Session(EditCurrentField & "Mask") = strMask
    Session(EditCurrentField & "MaskCompleter") = strCompleter

  Else
    Err.Raise 300, "EditFormFieldMask", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormRadioColumns

REM -------------------------------------------------------------------------
REM Rotina que finaliza a criacao da secao de tabela
REM -------------------------------------------------------------------------
REM Public Sub EditAddSameFieldValue(strField, strField2)
REM Dim i
REM   If EditFindField(strField) Then
REM    Session(EditCurrentField & "EqualNumber") = Session(EditCurrentField & "EqualNumber") + 1
REM    i = Session(EditCurrentField & "EqualNumber")
REM    Session(EditCurrentField & "," & i & "EqualField") = strField2
REM
REM  Else
REM    Err.Raise 300, "EditAddSameFieldValue", "O campo """ & strField & """, da tabela """ & _
REM          EditCurrentTable & """, não existe"

REM  End If

REM End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddFormField

REM -------------------------------------------------------------------------
REM Rotina que finaliza a criacao da secao de tabela
REM -------------------------------------------------------------------------
Public Sub EditFormEnd
Dim strJoinFormUnit, i, j, strTable, strFirstTable, strSecondTable

  strJoinFormUnit = ""

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    Session("CurrentField") = i
    If IsExternalField Then
REM      Response.Write EditCurrentTable & "*" & Session(EditCurrentField & "OwnerTable") & EditFindRelation(EditCurrentTable, Session(EditCurrentField & "OwnerTable"))
REM      ShowMessageError Session(EditCurrentTable & "RelationNumber")
REM      ShowMessageError Session(EditCurrentField & "Field") & "->" & Session(EditCurrentField & "OwnerTable") & " "
      If InStr(strJoinFormUnit, Session(EditCurrentField & "OwnerTable")) = 0 Then
        strTable = ""
        If EditFindRelation(EditCurrentTable, Session(EditCurrentField & "OwnerTable")) Then
           strTable       = EditCurrentTable & sintCurrentRelation
           strFirstTable  = EditCurrentTable
           strSecondTable = Session(EditCurrentField & "OwnerTable")

        Else
          If EditFindRelation(Session(EditCurrentField & "OwnerTable"), EditCurrentTable) Then
            strTable = Session(EditCurrentField & "OwnerTable") & sintCurrentRelation
            strFirstTable = Session(EditCurrentField & "OwnerTable")
            strSecondTable = EditCurrentTable

          End If
        End If

        If strTable > "" Then
          If Session(strTable & "FieldtoFieldNumber") <> 0 Then
             strJoinFormUnit = strJoinFormUnit & " LEFT OUTER JOIN " & Session(EditCurrentField & "OwnerTable")

          End If

          For j = 1 to Session(strTable & "FieldtoFieldNumber")
            If j = 1 Then
              strJoinFormUnit = strJoinFormUnit & " ON "

            Else
              strJoinFormUnit = strJoinFormUnit & " AND "

            End If

            strJoinFormUnit = strJoinFormUnit & strSecondTable & "." & Session(strTable & "," & j & "RelationField2") & " = " & strFirstTable & "." & Session(strTable & "," & j & "RelationField1")

          Next

        End If
      End If
    End If

  Next

  Session(EditCurrentTable & "JoinFormUnit") = strJoinFormUnit
REM  ShowMessageError Session(EditCurrentTable & "JoinFormUnit")

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormEnd

Private Function Font(ByVal intSize, ByVal strColor, ByVal strFace)
Dim strFont
 
  strFont = "<FONT"
  If CStr(intSize) > "" Then
    strFont = strFont & " SIZE=" & intSize  
    
  End If

  If strColor > "" Then
    strFont = strFont & " COLOR=" & strColor
    
  End If

  If strFace > "" Then
    strFont = strFont & " FACE=" & strFace
    
  End If

  strFont = strFont & ">"
  
  Font = strFont
    
End Function

Private Function FontEnd
  FontEnd = "</FONT>"
  
End Function

REM -------------------------------------------------------------------------
REM Rotina que cria o formato da tela de listagem da tabela
REM -------------------------------------------------------------------------
Public Sub EditFormFind(ByVal strFieldFind)
  Session(EditCurrentTable & "FieldFind") = strFieldFind

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormList

REM -------------------------------------------------------------------------
REM Rotina que cria o formato da listagem do form
REM -------------------------------------------------------------------------
Public Sub EditFormUnitList(ByVal strUnitlist)

  Session(EditCurrentTable & "UnitList") = strUnitlist

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormUnitList

REM -------------------------------------------------------------------------
REM Rotina que cria o formato da tela de listagem da tabela
REM -------------------------------------------------------------------------
Public Sub EditFormList(ByVal strFieldList)
  Session(EditCurrentTable & "FieldList") = strFieldList
  EditTotalOperation strFieldList, "ListOperationNumber", False, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormList

REM -------------------------------------------------------------------------
REM Rotina que cria o formato da tela de edicao da tabela
REM -------------------------------------------------------------------------
Public Sub EditFormUnit(ByVal strFieldUnit)
  Session(EditCurrentTable & "FieldUnit") = strFieldUnit
  EditTotalOperation strFieldUnit, "UnitOperationNumber", True, "UnitOperations"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormUnit

REM -------------------------------------------------------------------------
REM Rotina que encontra um campo em uma tabela
REM -------------------------------------------------------------------------
Public Function EditFindRelation(ByVal strTable, ByVal strTable2)
Dim i

  i = 1
  Do While (LCase(Session(strTable & i & "RelationTable")) <> LCase(strTable2)) And _
           (i < Session(strTable & "RelationNumber"))
    i = i + 1

  Loop

  If LCase(Session(strTable & i & "RelationTable")) = LCase(strTable2) Then
    EditFindRelation = True
    sintCurrentRelation = i

  Else
    EditFindRelation = False

  End If


End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFindField

REM -------------------------------------------------------------------------
REM Rotina que adiciona uma forma de ordenacao na tabela em edicao
REM -------------------------------------------------------------------------
Public Sub EditAddOrder(ByVal strOrder)

Dim i

  Session(EditCurrentTable & "OrderNumber") = Session(EditCurrentTable & "OrderNumber") + 1
  i = Session(EditCurrentTable & "OrderNumber")

  Session(EditCurrentTable & i & "Order")   = strOrder

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddOrder

REM -------------------------------------------------------------------------
REM Rotina que muda o tamanho padrao de apresentacao de um campo
REM -------------------------------------------------------------------------
Public Sub EditFieldShowSize(ByVal strField, ByVal intSize)
  If EditFindField(strField) Then
    Session(EditCurrentField & "ShowSize") = intSize

  Else
    Err.Raise 300, "EditFieldShowSize", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFieldShowSize

REM -------------------------------------------------------------------------
REM Rotina que muda os anos de inicio e fim do select de ano de um SeparateDateField
REM -------------------------------------------------------------------------
Public Sub EditFieldSepDate(ByVal intBeginYear, ByVal intEndYear)
  If EditFindField(strField) Then
     Session(EditCurrentField & "BeginYear") = intBeginYear
     Session(EditCurrentField & "EndYear")   = intEndYear

  Else
    Err.Raise 300, "EditFieldSepDate", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFieldSepDate

REM -------------------------------------------------------------------------
REM Rotina que cria um link externo para o campo
REM -------------------------------------------------------------------------
Public Sub EditFieldExternalLink(ByVal strField, ByVal strLink)
  If EditFindField(strField) Then
    Session(EditCurrentField & "ExternalLink") = strLink

  Else
    Err.Raise 300, "EditFieldExternalLink", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFieldExternalLink

REM -------------------------------------------------------------------------
REM Rotina que cria um link interno para o campo
REM -------------------------------------------------------------------------
Public Sub EditFieldInternalLink(ByVal strField, ByVal strLink, ByVal strILTarget)
  If EditFindField(strField) Then
    Session(EditCurrentField & "InternalLink") = strLink
    Session(EditCurrentField & "ILTargetNumber") = 1
    Session(EditCurrentField & ",1ILTarget") = strILTarget

  Else
    Err.Raise 300, "EditFieldInternalLink", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFieldInternalLink

REM -------------------------------------------------------------------------
REM Rotina que adiciona um target no link interno de um campo
REM -------------------------------------------------------------------------
Public Sub EditFieldAddILTarget(ByVal strField, ByVal strILTarget)
  If EditFindField(strField) Then
    Session(EditCurrentField & "ILTargetNumber") = Session(EditCurrentField & _
      "ILTargetNumber") + 1
    Session(EditCurrentField & "," & Session(EditCurrentField & "ILTargetNumber") & "ILTarget") = strILTarget

  Else
    Err.Raise 300, "EditFieldAddILTarget", "O campo """ & strField & """, da tabela """ & _
        EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFieldAddILTarget

REM -------------------------------------------------------------------------
REM Rotina que adiciona um valor possivel para um campo de selecao
REM -------------------------------------------------------------------------
Public Sub EditAddFormFieldValue(ByVal strField, ByVal Value, ByVal strCaption)
  If EditFindField(strField) Then
    Session(EditCurrentField & "ValueNumber") = Session(EditCurrentField & "ValueNumber") + 1

    Session(EditCurrentField & "," & Session(EditCurrentField & "ValueNumber") & "Value") = Value
    Session(EditCurrentField & "," & Session(EditCurrentField & "ValueNumber") & "Caption") = strCaption

  Else
    Err.Raise 300, "EditAddFormFieldValue", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddFormFieldValue

REM -------------------------------------------------------------------------
REM Guarda uma dica de edicao de um campo
REM -------------------------------------------------------------------------
Public Sub EditFormFieldHint(ByVal strField, ByVal strHint)
  If EditFindField(strField) Then
    Session(EditCurrentField & "Hint") = strHint

  Else
    Err.Raise 300, "EditFormFieldHint", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormFieldHint

REM -------------------------------------------------------------------------
REM Rotina que altera os dados da ligacao para um campo de ligacao com outra tabela
REM -------------------------------------------------------------------------
Public Sub EditFormFieldLookup(ByVal strField, ByVal strLookupTable, ByVal strLookupField, _
                              ByVal strLookupFieldText)
  If EditFindField(strField) Then
    Session(EditCurrentField & "LookupTable") = strLookupTable
    Session(EditCurrentField & "LookupField") = strLookupField
    Session(EditCurrentField & "LookupFieldText") = strLookupFieldText

  Else
    Err.Raise 300, "EditFormFieldLookup", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormFieldLookup

REM -------------------------------------------------------------------------
REM Rotina que adiciona um header a listagem
REM -------------------------------------------------------------------------
Public Sub EditAddHeader(strHeader, intBegin, intWidth)
Dim i

  Session(EditCurrentTable & "HeaderNumber") = Session(EditCurrentTable & "HeaderNumber") + 1
  i = Session(EditCurrentTable & "HeaderNumber")

  Session(EditCurrentTable & i & "HeaderText") = strHeader
  Session(EditCurrentTable & i & "HeaderBegin") = intBegin
  Session(EditCurrentTable & i & "HeaderWidth") = intWidth

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddHeader

REM -------------------------------------------------------------------------
REM Rotina que encontra um campo em uma tabela
REM -------------------------------------------------------------------------
Public Function EditFindField(ByVal strField)
Dim i

  i = 1
  Do While (LCase(Session(EditCurrentTable & i & "Field")) <> LCase(strField)) And _
           (i < Session(EditCurrentTable & "FieldNumber"))

    i = i + 1

  Loop

  If LCase(Session(EditCurrentTable & i & "Field")) = LCase(strField) Then
    EditFindField = True

    Session("CurrentField") = i

  Else
    EditFindField = False

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFindField

REM -------------------------------------------------------------------------
REM Rotina que encontra um campo em uma tabela
REM -------------------------------------------------------------------------
Public Function EditFindList(ByVal strList)
Dim i

  i = 1
  Do While (LCase(Session(EditCurrentTable & i & "ListName")) <> LCase(strList)) And _
           (i < Session(EditCurrentTable & "ListNumber"))

    i = i + 1

  Loop

  If LCase(Session(EditCurrentTable & i & "ListName")) = LCase(strList) Then
    EditFindList = True

    Session("CurrentList") = i

  Else
    EditFindList = False

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFindField

REM -------------------------------------------------------------------------
REM Rotina retorna o nome do campo encontrado pela function EditFindField
REM -------------------------------------------------------------------------
Public Function EditCurrentField
  EditCurrentField = EditCurrentTable & Session("CurrentField")

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCurrentField

REM -------------------------------------------------------------------------
REM Rotina retorna o nome do campo encontrado pela function EditFindField
REM -------------------------------------------------------------------------
Public Function EditCurrentList
  EditCurrentList = EditCurrentTable & Session("CurrentList")

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCurrentField

REM -------------------------------------------------------------------------
REM Rotina retorna o nome do campo encontrado pela function EditFindField
REM -------------------------------------------------------------------------
Public Function EditFormAddList(strList, strCaption, strName)
Dim i
  Session(EditCurrentTable & "ListNumber") = Session(EditCurrentTable & "ListNumber") + 1
  i = Session(EditCurrentTable & "ListNumber")

  Session(EditCurrentTable & i & "ListName")          = strName
  Session(EditCurrentTable & i & "List")              = strList
  Session(EditCurrentTable & i & "ListCaption")       = strCaption

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCurrentField

REM -------------------------------------------------------------------------
REM Rotina que inicia a leitura de um campo de formato
REM -------------------------------------------------------------------------
Public Sub EditBeginReadOperation(ByVal strData)
  If strData > "" Then
    sstrText = strData

  Else
    sstrText = ""

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditBeginReadOperation

REM -------------------------------------------------------------------------
REM Rotina que inicia a leitura de um campo de formato
REM -------------------------------------------------------------------------
Public Sub EditBeginReadNumbers(ByVal strData)
  If strData > "" Then
    sstrText2 = strData

  Else
    sstrText2 = ""

  End If
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditBeginReadOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna a existencia de dados no texto de formato
REM -------------------------------------------------------------------------
Public Function EditIsThereOperation
  If sstrText > "" Then
    EditIsThereOperation = True

  Else
    EditIsThereOperation = False

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditIsThereOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna a existencia de dados no texto de formato
REM -------------------------------------------------------------------------
Public Function EditIsThereNumber
  If sstrText2 > "" Then
    EditIsThereNumber = True

  Else
    EditIsThereNumber = False

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditIsThereOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna um dado ou operacao do texto de formato
REM -------------------------------------------------------------------------
Public Function EditNextOperation
Dim strSubText, strChar, blnEnd

  strChar = Left(sstrText, 1)
  blnEnd = False

  Do
    Select Case strChar
      Case "."
        If strSubText = "" Then
          Session("CurrentTable") = EditLastOperation

        End If
        sstrText = Mid(sstrText, 2)
        strChar = Left(sstrText, 1)

      Case "," , ";", "-", "=", "(", ")", "#", ""
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

   EditNextOperation = strSubText
   sstrSubText = strSubText

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditNextOperation

Private Function EditLastOperation

  EditLastOperation = sstrSubText

End Function

REM -------------------------------------------------------------------------
REM Rotina que retorna um dado ou operacao do texto de formato mas nao o exclui
REM -------------------------------------------------------------------------
Public Function EditSeeNextOperation

  EditSeeNextOperation = EditNextOperation
  sstrText             = EditLastOperation & sstrText

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditSeeNextOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna um dado ou operacao do texto de formato
REM -------------------------------------------------------------------------
Public Function EditNextNumber
Dim strSubText, strChar, blnEnd, i

  i = 0
  strChar = Left(sstrText2, 1)
  blnEnd = False

  Do
    Select Case strChar
      Case ",", ""
        sstrText2 = Mid(sstrText2, 2)
        If strSubText <> "" Then
          blnEnd = True

        End If

      Case Else
        strSubText = strSubText & strChar
        sstrText2 = Mid(sstrText2, 2)
        strChar = Left(sstrText2, 1)

    End Select
    i = i + 1

  Loop While Not blnEnd And i < 20
  If strSubText > "" Then
    EditNextNumber = CInt(strSubText)
    sintLastNumber = CInt(strSubText)

  Else
    EditNextNumber = 1
    sintLastNumber = 1

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditNextOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna o numero de campos por linha de um formato
REM -------------------------------------------------------------------------
Public Sub EditTotalOperation(strText, strTotal, blnKeepParts, strParts)
Dim intGreat, i

  EditBeginReadOperation strText

  Session(EditCurrentTable & strParts) = ""

  i = 0
  intGreat = 0

  Do While EditIsThereOperation
    Select Case EditNextOperation

      Case ","
        If intGreat < i Then
          intGreat = i

        End If

      Case "="
        If intGreat < i Then
          intGreat = i

        End If
        EditNextOperation

      Case ";", "-", ""
        If intGreat < i Then
          intGreat = i

        End If

        If blnKeepParts Then
          If Session(EditCurrentTable & strParts) > "" Then
            Session(EditCurrentTable & strParts) = Session(EditCurrentTable & _
             strParts) + "," + CStr(i)

          Else
            Session(EditCurrentTable & strParts) = CStr(i)

          End If
        End If
        i = 0

      Case Else
        i = i + 1

    End Select

    IF EditLastOperation = "=" Then
      EditNextOperation

    End If
   Loop

   If blnKeepParts Then
     If Session(EditCurrentTable & strParts) > "" Then
       Session(EditCurrentTable & strParts) = Session(EditCurrentTable & _
        strParts) + "," + CStr(i)

     Else
       Session(EditCurrentTable & strParts) = CStr(i)

     End If
   End If

   Session(EditCurrentTable & strTotal) = intGreat

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTotalOperation

REM -------------------------------------------------------------------------
REM Rotina que le os dados e tamanhos dos campos ja adicionados na tabela
REM -------------------------------------------------------------------------
Private Function EditTableOf(intField)
  If Session(EditCurrentTable & intField & "OwnerTable") > "" Then
    EditTableOf = Session(EditCurrentTable & intField & "OwnerTable")

  Else
    EditTableOf = EditCurrentTable

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditTableOf

REM -------------------------------------------------------------------------
REM Rotina que le os dados e tamanhos dos campos ja adicionados na tabela
REM -------------------------------------------------------------------------
Public Function EditReadFieldValues(blnEdit)
Dim sql, i

  Application.Lock
  Application(EditCurrentTable & EditMakeTarget) = ""
  Application.Unlock
  If (Not Application(EditCurrentTable & EditMakeTarget) > "") Or (Not blnEdit) Then
    SetPutAnd False

    sql = "SELECT "

    For i = 1 to Session(EditCurrentTable & "FieldNumber")
      If i > 1 Then
        sql = sql & ", "

      End If

      sql = sql & EditTableOf(i) & "." & Session(EditCurrentTable & i & "Field")

    Next

    sql = sql & " FROM " & EditCurrentTable & Session(EditCurrentTable & "JoinFormUnit")

    If sparEditFilter > "" Then
      sql = sql & " WHERE " & sparEditFilter

    End If

REM    ShowMessageError sql

    sobjConn.Close
    sobjConn.Open Session("EditConnString")

    sobjCMD.ActiveConnection = sobjConn
    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    sobjRS.CursorLocation = adUseClient

REM    ShowMessageError sql
    sobjRS.Open sobjCMD, , adOpenStatic, adLockReadOnly

REM   Lista(sobjRS)

    If Not sobjRS.Eof And Not sobjRS.Bof Then
      If (sparEditMove <> conMoveFirst) And (sparEditMove <> conMoveLast) Then
        If sparEditFind > "" Then
          sobjRS.Find sparEditFind

        End If
      End If

      Select Case sparEditMove
        Case conMoveFirst
          sobjRS.MoveFirst

        Case conMovePrevious
          sobjRS.MovePrevious

        Case conMoveNext
          sobjRS.MoveNext

        Case conMoveLast
          sobjRS.MoveLast

      End Select

      If Not sobjRS.EOF And Not sobjRS.BOF Then
        For i = 1 to Session(EditCurrentTable & "FieldNumber")
          Session(EditCurrentTable & i & "OldValue") = sobjRS(Session(EditCurrentTable & i & "Field"))
          Session(EditCurrentTable & i & "Value")    = sobjRS(Session(EditCurrentTable & i & "Field"))
          REM Response.Write Session(EditCurrentTable & i & "Field")
        Next

      End If
      sintAbsolutePosition = sobjRS.AbsolutePosition
      sintRecordCount		  = sobjRS.RecordCount
REM      ShowMessageError "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""Porra Me-o"");</SCRIPT>"

    Else
      If Not blnEdit Then
        If Not sparEditFilter > "" Then
          Session("RecordAction") = 1
          sparEditOption = conOptionAdd
          EditClear sparEditOption

        Else
          If Session(EditCurrentTable & "ClientValidation") Then
            Session("Message") = "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""Registro não localizado"");</SCRIPT>"

          End If
          Session("Message") = Session("Message") & "<FONT COLOR=RED>" & "<B>Registro não localizado</B></FONT>"
          sparEditFilter     = ""
          EditReadFieldValues = EditReadFieldValues(blnEdit)
REM          ShowMessageError "<SCRIPT LANGUAGE=JAVASCRIPT>alert(""Ura Meu"");</SCRIPT>"

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

REM    ShowMessageError "Total:" & Session(EditCurrentTable & "UnitOperationNumber")
REM    ShowMessageError "Parcial:" & sintLastNumber
REM    ShowMessageError "Numero de colunas:" & sintNumCol
REM    ShowMessageError "ColSpan:" & intColSpan

  If intColSpan > 1 Then
    TableSetColumnColSpan intColSpan

  Else
    TableSetColumnColSpan ""

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditColumnColSpan

REM -------------------------------------------------------------------------
REM Rotina cria os javascripts das mensagens de exclusao e salvamento
REM -------------------------------------------------------------------------
Public Sub JavaScriptDefaultMessages

  ShowHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"

  ShowHTMLCR "function ValueToHidden(valor)"
  ShowHTMLCR "  {"
  ShowHTMLCR "  switch (valor){"
  ShowHTMLCR "  case """ & constrDelete & """ :"
  ShowHTMLCR "  return confirm(""Você tem certeza que deseja excluir este registro ?""); break;"
  If Session("RecordAction") > "" Then
    ShowHTMLCR "  case """ & constrCancel & """ :"
    ShowHTMLCR "  return confirm(""Você tem certeza que deseja cancelar a inclusão do registro ? \nOs dados digitados serão perdidos.""); break;"

  End If

  If Session(EditCurrentTable & "ClientValidation") And Session("RecordAction") > "" Then
    ShowHTMLCR "  case """ & constrSave & """ :"
    ShowHTMLCR "  case """ & constrSaveAdd & """ :"
    ShowHTMLCR "  case """ & constrSaveCopy & """ :"
    ShowHTMLCR "  return ValidateForm(); break;"

  End If

  ShowHTMLCR "  default : return true;"
  ShowHTMLCR "  }}"

  ShowHTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DeleteSaveMessages

REM -------------------------------------------------------------------------
REM Inclui uma imagem na tela
REM -------------------------------------------------------------------------
Private Sub ShowImage(strImagem, strID, strAlt, strHRef, strAlign, intWidth, intHeight)

  If strHRef > "" Then
    Response.Write "<A HREF=" & strHRef & ">"

  End If

  Response.Write "<IMG SRC=""" & strImagem & """ BORDER=0"

  If strAlt > "" Then
    Response.Write " ALT=""" & strAlt & """"

  End If

  If strAlign > "" Then
    Response.Write " ALIGN=""" & strAlign & """"

  End If

  If strID > "" Then
    Response.Write " ID=""" & strID & """"

  End If

  If intWidth > "" Then
    Response.Write " WIDTH=" & intWidth

  End If

  If intHeight > "" Then
    Response.Write " HEIGHT=" & intHeight

  End If

	If sFormInputonClick > "" Then
		Response.Write " onClick=""" & sFormInputonClick & """"

	End If

  Response.Write ">"

  If strHRef > "" Then
    Response.Write "</A>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DeleteSaveMessages

REM -------------------------------------------------------------------------
REM Entra botao Grafico
REM -------------------------------------------------------------------------
Sub FormInputImage(txtName, txtValue, txtImage, strAlt)

	FormSetAlt strAlt
	FormSetNoCrLf
	FormInput "image", txtName, 0, txtValue, 0, "", -1, txtImage

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputImage

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

Private Sub EditGrid
Dim sql, strAux, strAux2, arrFields(30), intCount, i

  sql = "SELECT * FROM " & EditCurrentTable & Session(EditCurrentTable & "JoinFormUnit")

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Session(EditCurrentTable & i & "EditOption") = conLookupField Then
       sql = sql & " LEFT OUTER JOIN " & Session(EditCurrentTable & i & "LookupTable") &_
       " on " & Session(EditCurrentTable & i & "LookupTable") & "." & Session(EditCurrentTable & i & "LookupField") & " = " & _
       EditTableOf(i) & "." & Session(EditCurrentTable & i & "Field")

    End If

  Next

  If sparEditFilter > "" Then
    sql = sql & " WHERE " & sparEditFilter

  End If

  If Not Session("OrderField") > "" Then
    EditBeginReadOperation Session(EditCurrentTable & "FieldList")

    Do While EditIsThereOperation
      If EditFindField(EditNextOperation) Then
        sstrText = ""
        Session("OrderField") = Session("CurrentField")

      End If

    Loop

  End If

  Dim blnFirst
  EditBeginReadOperation Session("OrderField")
  sql = sql & " ORDER BY "

  blnFirst = True

  Do While EditIsThereOperation
    Select Case EditNextOperation
       Case ","

       Case Else
         Session("CurrentField") = EditLastOperation
         If Not blnFirst Then
           sql = sql & ", "

         Else
           sql = sql & " "

         End If

         Select Case Session(EditCurrentField & "EditOption")
           Case conLookUpField
             sql = sql & Session(EditCurrentField & "LookupTable") & "." & Session(EditCurrentField & "LookupFieldText")

           Case Else
             sql = sql & Session(EditCurrentField & "Field")

        End Select

        blnFirst = False
        If Session("OrderDesc") = "1" Then
          sql = sql & " DESC "

        End If

    End Select

  Loop

  BrowseSetActual "Grid"
  BrowseSetPageSize 5

  Nbsp

  sobjCMD.CommandText = sql
  sobjCMD.CommandType = adCmdText

REM  ShowMessageError sql

  sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

  If Not sobjRS.EOF And Not sobjRS.BOF Then
    BrowseBegin
    BrowseHeader

    EditBeginReadOperation Session(EditCurrentTable & "UnitList")

    ShowMessageError "" 

    Dim strList

    Do While EditIsThereOperation
      Select Case EditNextOperation
         Case ",", ";", "-"
         Case "="
           EditIsThereOperation

         Case Else
           If EditFindField(EditLastOperation) then
             If strList > "" Then
               strList = strList & ","

             End If
             strList = strList & Session("CurrentField")
 
           End If

       End Select

    Loop

    ClassAss

    TableBegin "", Session(EditCurrentTable & "ListOperationNumber") + 1, "600"

    TableSetColumnAlign ""
    TableSetColumnColSpan ""
    TableBeginHeadRow 2

    EditBeginReadOperation strList

    Do While EditIsThereOperation
      Select Case EditNextOperation
        Case ",",";","-","="

        Case Else
          Session("CurrentField") = EditLastOperation
          EditTableHeader EditCaption(False), EditLastOperation, Session(EditCurrentTable & EditLastOperation & "ListSize")

       End Select

    Loop

    TableEndHeadRow

Dim blnActualRecord

    Do While Not sobjRS.EOF And IsBrowsePage
      EditBeginReadOperation strList

      TableBeginRow 1

      blnActualRecord = True
      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "isKey") Then
          If Session(EditCurrentTable & i & "Value") <> sobjRS(Session(EditCurrentTable & i & "Field")) Then
            blnActualRecord = False

          End If
        End If

      Next

      Do While EditIsThereOperation
        Select Case EditNextOperation
          Case ",",";","-","="

          Case Else
            Session("CurrentField") = EditLastOperation
            If TypeName(EditFieldValue) = "Long" Or TypeName(EditFieldValue) = "Integer" Or _
               TypeName(EditFieldValue) = "Currency" Or TypeName(EditFieldValue) = "Double" Or TypeName(EditFieldValue) = "Single" Then
              TableSetColumnAlign "Right"

            End If

            arrFields(intCount) = Session(EditCurrentField & "Value")

            TableBeginColumn
            EditBeginMakeLink False
            If Not blnActualRecord Then
              Select Case Session(EditCurrentField & "EditOption")
                Case conLookupField, conRadioField, conCheckField, conSelectField
                  ShowHTMLCR BreakString(EditFieldValue, Session(EditCurrentField & "ListChars"))

                Case Else
                  ShowHTMLCR BreakString(PutMask(FormatVariable(EditFieldValue, Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars"))

              End Select

            Else
              Select Case Session(EditCurrentField & "EditOption")
                Case conLookupField, conRadioField, conCheckField, conSelectField
                  ShowHTMLCR "<font color=red>" & BreakString(EditFieldValue, Session(EditCurrentField & "ListChars")) & "</font>"

                Case Else
                  ShowHTMLCR "<font color=red>" & BreakString(PutMask(FormatVariable(EditFieldValue, Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars")) & "</font>"

              End Select

            End If
            EditEndMakeLink
            TableSetColumnAlign ""
            TableEndColumn

        End Select

      Loop
      TableEndRow

      sobjRS.MoveNext
      BrowseLineCounter
    Loop

    TableEnd

    BrowseSetStatus True
    BrowseEnd
REM    BrowseBar

  End If

  sobjRS.Close

End Sub

REM -------------------------------------------------------------------------
REM Rotina que mostra a tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditShowEditor
Dim blnNotStarting, FontSize, strURL, i, blnFirstEditable, strLastOperation, blnDisable

  Response.Write "<STYLE> .BTM{font: 8pt Arial}</STYLE>"

  FormSetClass "BTM"

  Session("LastToFind") = conOptionFPEdit
  blnFirstEditable = True
  BR
  Center

  JavaScriptDefaultMessages

  If Session(EditCurrentTable & "ClientValidation") Then
    If Session("RecordAction") > "" Then
      ClientValidates

    End If

  End If

  If Session("Message") > "" Then
    TableBegin "", 1, "65%"
    TableBeginRow 2
    TableBeginColumn
    ShowHTMLCR Session("Message")
    TableEndColumn
    TableEndRow
    TableEnd

    Session("Message") = ""

  End If

  strURL = sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionExeEdition & GetFilter & EditWhatWhereHow

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Request.QueryString("F" & i) > "" Then
      strURL = strURL & "&" & "F" & i & "=" & Request.QueryString("F" & i)

    End If

  Next

  FormBegin strURL, "POST", "frmEditForm", "", ""

  If Session(EditCurrentTable & "EditOption") <> 0 Then
    FontSize = 1

  Else
    FontSize = 2

  End If

  EditBeginReadNumbers Session(EditCurrentTable & "UnitOperations")

  Dim strMessage

  If Not Session("RecordAction") > "" Then
    strMessage = Session(EditCurrentTable & "Caption")
    FormSetDisabled "1"
    blnDisable = True

  ElseIf Session("RecordAction") = 1 Then
    strMessage = Session(EditCurrentTable & "Caption") & " - Inclusão"
    blnDisable = False

  ElseIf Session("RecordAction") = 2 Then
    strMessage = Session(EditCurrentTable & "Caption") & " - Alteração"
    blnDisable = False

  End If

  TableBegin strMessage, 1, "65%"

  blnNotStarting = False

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")

  TableBeginHeadRow 2
  TableSetColumnAlign "LEFT"
  TableBeginHeadColumn
  Response.Write "<FONT SIZE=1>"
  If sparEditFilter > "" Then
    ShowHTMLCR "Tabela Filtrada<BR>&nbsp;&nbsp;"

  End If

  ShowHTMLCR "Nº de registros&nbsp;:&nbsp;" & sintRecordCount

  If sparEditFilter > "" Then
    ShowHTMLCR "<BR>"

  End If
  ShowHTMLCR "</FONT>"
  TableEndHeadColumn
  TableSetColumnAlign ""
  TableEndHeadRow

  If Left(sstrText, 1) <> "=" Then
    StartRow
    TableBeginRow 1
    TableBeginColumn
    TableNormalBegin "100%"
    blnNotStarting = True
    EditColumnColSpan
    TableBeginRow FontSize
    TableBeginColumn

  End If

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ","
        IncNumCol
        EditColumnColSpan
        TableEndColumn
        TableBeginColumn

      Case ";"
        StartRow
        TableEndColumn
        TableEndRow
        EditColumnColSpan
        TableBeginRow FontSize
        TableBeginColumn

      Case "-"
        StartRow
        TableEndColumn
        TableEndRow
        TableNormalEnd
        TableEndColumn
        TableEndRow
        TableBeginRow 1
        TableBeginColumn
        TableNormalBegin "100%"
        EditColumnColSpan
        TableBeginRow FontSize
        TableBeginColumn

      Case "="
        StartRow
        If blnNotStarting Then
          TableEndColumn
          TableEndRow
          TableNormalEnd
          TableEndColumn
          TableEndRow

        End If

        blnNotStarting = True

        TableBeginHeadRow 2
        TableSetColumnAlign "CENTER"
        TableHeadColumn EditNextOperation
        TableSetColumnAlign ""
        TableEndHeadRow
        TableBeginRow 1
        TableBeginColumn
        TableNormalBegin "100%"
        EditColumnColSpan
        TableBeginRow FontSize
        TableBeginColumn
        EditNextOperation

      Case Else
        If (Session(EditCurrentTable & "EditOption") > 0) And (sparEditEditable = 1) Then
          blnFirstEditable = EditShowField(blnFirstEditable And Session("RecordAction") > "", False, blnDisable, False)

        Else
          EditShowFieldAsText

        End If

    End Select

  Loop

  TableEndColumn
  TableEndRow
  TableNormalEnd

  If Session(EditCurrentTable & "UnitList") > "" Then
    EditGrid

  End If

  TableEndColumn
  TableEndRow

  TableBeginRow ""
  TableBeginColumn
  Center

  FormSetNoCrLf

  If Not Session("RecordAction") > "" Then

    If sintAbsolutePosition > 1 Then
      FormSetDisabled ""
      FormInputSetOnClick "return ValueToHidden('" & constrFirst & "')"

    Else
      FormSetDisabled "1"

    End If
    FormSetTitle "Primeiro Registro"
    FormInputSubmit "btmEditor", constrFirst
REM FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdPrimeiroWin.gif", "Vai para o primeiro registro"

    If sintAbsolutePosition > 1 Then
      FormSetDisabled ""
      FormInputSetOnClick "return ValueToHidden('" & constrPrevious & "')"

    Else
      FormSetDisabled "1"

    End If
    FormSetTitle "Registro Anterior"
    FormInputSubmit "btmEditor", constrPrevious
REM    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdAnteriorWin.gif", "Vai para o registro anterior"

  End If

  FormSetDisabled ""

  If sparEditEditable = 1 Then
    Select Case Session(EditCurrentTable & "EditOption")
      Case "1"
        If Not Session("RecordAction") > "" Then
          FormSetAccessKey "I"
          FormInputSetOnClick "return ValueToHidden('" & constrAdd & "')"

          FormSetTitle "inclui um registro"
          FormInputSubmit "btmEditor", constrAdd
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdIncluirWin.gif", "Inclui um registro"
          FormSetAccessKey ""
          FormInputSetOnClick "return ValueToHidden('" & constrCopy & "')"
          FormSetTitle "Inclui um registro copiando os dados do registro corrente"
          FormInputSubmit "btmEditor", constrCopy
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdCopiarWin.gif", "Inclui um registro copiando os dados deste"
          FormSetAccessKey "A"
          FormInputSetOnClick "return ValueToHidden('" & constrEdit & "')"
          FormSetTitle "Edita o registro corrente"
          FormInputSubmit "btmEditor", constrEdit
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdAlterarWin.gif", "Entra em modo de alteração"
          FormSetAccessKey "P"
          FormInputSetOnClick "return ValueToHidden('" & constrDelete & "')"
          FormSetTitle "Exclui o registro corrente"
          FormInputSubmit "btmEditor", constrDelete
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdApagarWin.gif", "Exclui este registro"
          FormSetAccessKey ""

        End If

        If Not Session("RecordAction") > "" Then
REM        FormInputSubmit "btmEditor", constrFind

        Else
          FormSetAccessKey "S"
          FormInputSetOnClick "return ValueToHidden('" & constrSave & "')"
          FormSetTitle "Salta as alterações efetuadas"
          FormInputSubmit "btmEditor", constrSave
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarWin.gif", "Salva as alterações feitas neste registro"
          FormSetAccessKey ""
          FormInputSetOnClick "return ValueToHidden('" & constrSaveAdd & "')"
          FormSetTitle "Salta as alterações efetuadas e entra em modo de inclusão"
          FormInputSubmit "btmEditor", constrSaveAdd
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarIncluirWin.gif", "Salva os dados deste registro e inclui outro"
          FormInputSetOnClick "return ValueToHidden('" & constrSaveCopy & "')"
          FormSetTitle "Salta as alterações efetuadas e entra em modo de inclusão copiando os dados do registro corrente"
          FormInputSubmit "btmEditor", constrSaveCopy
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarCopiarWin.gif", "Salva os dados deste registro e copiando os dados deste"
          FormSetAccessKey "C"
          FormInputSetOnClick "return ValueToHidden('" & constrCancel & "')"
          FormSetTitle "Cancela todas as alterações efetuadas"
          FormInputSubmit "btmEditor", constrCancel
REM          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdCancelarWin.gif", "Cancela a inclusão deste registro"
          FormSetAccessKey ""

        End If

    End Select
  End If

  If Not Session("RecordAction") > "" Then
    FormInputSetOnClick "return ValueToHidden('" & constrFind & "')"
    FormSetTitle "Entra em modo de localização de registros"
    FormInputSubmit "btmEditor", constrFind
REM    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdLocalizarWin.gif", "Entra em modo de localização de registros"
    FormInputSetOnClick "return ValueToHidden('" & constrList & "')"
    FormSetTitle "Entra no modo de visualização dos registros em lista"
    FormInputSubmit "btmEditor", constrList
REM    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdListagemWin.gif", "Vai para a forma de visualização de listagem"

    If sintAbsolutePosition = sintRecordCount Then
      FormSetDisabled "1"

    Else
      FormInputSetOnClick "return ValueToHidden('" & constrNext & "')"

    End If
    FormSetTitle "Próximo Registro"
    FormInputSubmit "btmEditor", constrNext
REM    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdProximoWin.gif", "Vai para o próximo registro"

    If sintAbsolutePosition = sintRecordCount Then
      FormSetDisabled "1"

    Else
      FormInputSetOnClick "return ValueToHidden('" & constrLast & "')"

    End If
    FormSetTitle "Último Registro"
    FormInputSubmit "btmEditor", constrLast
REM    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdUltimoWin.gif", "Vai para o úlitmo registro"
    FormInputSetOnClick ""

  End If

  FormInputSetOnClick ""
  FormSetDisabled ""
REM  FormInputReset "Apaga"

  TableEndColumn
  TableEndRow

  TableEnd

  FormEnd

  EditFPLPLink

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowEditor

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
  Center

  Response.Write "<FONT SIZE=2>"
  
  If sparEditOption <> conOptionFirstPage Then
    If Session("EditFPLPLink") > "" Then
      Response.Write "<A HREF=" & Session("EditFPLPLink")

    Else
      Response.Write "<A HREF=" & sstrThisScriptName

    End If

    ShowHTMLCR ">Página Inicial</a>"

  End If

  If QueryStringLast > "" Then
    ShowHTMLCR "<A HREF=" & QueryStringLast & ">Última Opção</a>"

  End If

  Response.Write "</FONT>"

  CenterEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFPLPLink

REM -------------------------------------------------------------------------
REM Limpa as variaveis da tabela corrente
REM -------------------------------------------------------------------------
Private Sub EditClear(intClearOption)
Dim sql, i, intPos, intPos2, strAux
  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    Select Case  Left(Session(EditCurrentTable & i & "DefaultValue"), 8)
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

        sobjRS2.Open sobjCMD, , adOpenStatic, adLockReadOnly

        If Not sobjRS2.EOF And Not sobjRS2.BOF Then
          sobjRS2.MoveLast
          Session(EditCurrentTable & i & "Value") = sobjRS2(strField) + 1
          Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

        Else
          Session(EditCurrentTable & i & "Value") = 1
          Session(EditCurrentTable & i & "OldValue") = 1

        End If
        sobjRS2.Close

      Case Else
        If intClearOption = conOptionAdd Then
          If sparEditDefaults > "" Then
            intPos = InStr(sparEditDefaults, LCase(Session(EditCurrentTable & i & "Field")) & "=")
            If intPos > 0 Then
              intPos = intPos + Len(Session(EditCurrentTable & i & "Field"))
              strAux = Right(sparEditDefaults, Len(sparEditDefaults) - intPos)
              intPos2 = InStr(strAux, ";")
              If intPos2 > 0 Then
                Session(EditCurrentTable & i & "Value") = EditFormatField(Left(strAux, intPos2 - 1), Session(EditCurrentTable & i & "Type"), Session(EditCurrentTable & i & "Size"))
                
              Else
                Session(EditCurrentTable & i & "Value") = EditFormatField(strAux, Session(EditCurrentTable & i & "Type"), Session(EditCurrentTable & i & "Size"))
              
              End If  

            Else
              Session(EditCurrentTable & i & "Value") = Session(EditCurrentTable & _
                                                                  i & "DefaultValue")
            End If

          Else
            Session(EditCurrentTable & i & "Value") = Session(EditCurrentTable & _
                                                                i & "DefaultValue")
          End If

        End If

    End Select
REM ShowMessageError Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "Value")
  Next

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditClear

REM -------------------------------------------------------------------------
REM Rotina que retorna o tamanho do campo para visualizacao
REM -------------------------------------------------------------------------
Public Function EditSizeToShow
  If Session(EditCurrentField & "ShowSize") > "" Then
    EditSizeToShow = Session(EditCurrentField & "ShowSize")

  Else
    EditSizeToShow = EditSize
    
  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditSizeToShow

REM -------------------------------------------------------------------------
REM Rotina que retorna o tamanho do campo em caracteres
REM -------------------------------------------------------------------------
Public Function EditSize
  Select Case Session(EditCurrentField & "Type")
    Case conDataMoney, conDataInt, conDataFloat
      EditSize = Len(CStr(256^Session(EditCurrentField & "Size"))) - 3

    Case conDataDateTime
      EditSize = 10
      
    Case Else
      EditSize = Session(EditCurrentField & "Size")
      
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
    ShowImage "/ztitools/img/editform/CampoRequerido.gif", "", "Campo obrigatório", "", "", 5, 5
    Nbsp

  End If
  If Session(EditCurrentField & "ValidationError") > "" Then
    ShowImage "/ztitools/img/editform/CampoErrado.gif", "", Session(EditCurrentField & "ValidationError"), "", "", 5, 5
    Session(EditCurrentField & "ValidationError") = ""

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Function EditShowFieldImages

REM -------------------------------------------------------------------------
REM Retorna a letra correspondente ao parametro, sem acentuacoes
REM -------------------------------------------------------------------------
Private Function ChangeChar(Char)
  If InStr("áâãàa", LCase(Char)) > 0 Then
    ChangeChar = "A"

  ElseIf InStr("éêe", LCase(Char)) > 0 Then
    ChangeChar = "E"

  ElseIf InStr("íi", LCase(Char)) > 0 Then
    ChangeChar = "I"

  ElseIf InStr("óõôo", LCase(Char)) > 0 Then
    ChangeChar = "O"

  ElseIf InStr("úu", LCase(Char)) > 0 Then
    ChangeChar = "U"

  ElseIf InStr("çc", LCase(Char)) > 0 Then
    ChangeChar = "C"

  Else
    ChangeChar = UCase(Char)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub ChangeChar

REM -------------------------------------------------------------------------
REM Inclui um texto de caption tratado na tela
REM -------------------------------------------------------------------------
Private Sub EditShowCaption(blnForm)
Dim i

  i = InStr(Session(EditCurrentField & "Caption"), "&")

  If i > 0 Then
    If blnForm Then
      ShowHTMLCR Left(Session(EditCurrentField & "Caption"), i - 1) & "<U>" & Mid(Session(EditCurrentField & "Caption"), i + 1 , 1) & "</U>" & Mid(Session(EditCurrentField & "Caption"), i + 2)
      FormSetAccessKey ChangeChar(Mid(Session(EditCurrentField & "Caption"), i + 1 , 1))

    Else
      ShowHTMLCR Left(Session(EditCurrentField & "Caption"), i - 1) & Mid(Session(EditCurrentField & "Caption"), i + 1)

    End If

  Else
    ShowHTMLCR Session(EditCurrentField & "Caption")

  End If


End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowCaption

REM -------------------------------------------------------------------------
REM Retorna um texto de caption tratado
REM -------------------------------------------------------------------------
Private Function EditCaption(blnForm)
Dim i

  i = InStr(Session(EditCurrentField & "Caption"), "&")

  If i > 0 Then
    If blnForm Then
      EditCaption = Left(Session(EditCurrentField & "Caption"), i - 1) & "<U>" & Mid(Session(EditCurrentField & "Caption"), i + 1 , 1) & "</U>" & Mid(Session(EditCurrentField & "Caption"), i + 2)
      FormSetAccessKey ChangeChar(Mid(Session(EditCurrentField & "Caption"), i + 1 , 1))

    Else
      EditCaption = Left(Session(EditCurrentField & "Caption"), i - 1) & Mid(Session(EditCurrentField & "Caption"), i + 1)

    End If

  Else
    EditCaption = Session(EditCurrentField & "Caption")

  End If


End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCaption

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
REM Retorna a divisao do primeiro valor pelo segundo arredondado para cima, tornando o reultado inteiro
REM -------------------------------------------------------------------------
Private Function DivRoundUp(intValue, intValue2)

  If (intValue / intValue2) > (intValue \ intValue2) Then
    DivRoundUp = (intValue \ intValue2) + 1

  Else
    DivRoundUp = intValue \ intValue2

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub DivRoundUp

REM -------------------------------------------------------------------------
REM Retorna o proximo caracter numerico do texto a ser mascarado
REM -------------------------------------------------------------------------
Private Function MaskNextNumber(strMaskCompleter)
Dim blnOK, strChar

  blnOK = False
  Do While (Not blnOK) And (Len(sstrText3) > 0)
    If (Left(sstrText3, 1) >= "0") And (Left(sstrText3, 1) <= "9") Then
      strChar = Left(sstrText3, 1)
      blnOK = True

    End If
    sstrText3 = Right(sstrText3, Len(sstrText3) - 1)

  Loop

  If (Not sstrText3 > "") And (Not strChar > "") Then
    MaskNextNumber = strMaskCompleter

  Else
    MaskNextNumber = strChar

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function MaskNextNumber

REM -------------------------------------------------------------------------
REM Retorna o proximo caracter alfa do texto a ser mascarado
REM -------------------------------------------------------------------------
Private Function MaskNextAlfa(strMaskCompleter)
Dim blnOK

  blnOK = False
  Do While (Not blnOK) And (Len(sstrText3) > 0)
    If ((Left(sstrText3, 1) >= "a") And (Left(sstrText3, 1) <= "z")) Or ((Left(sstrText3, 1) >= "A") And (Left(sstrText3, 1) <= "Z")) Then
      strChar = Left(sstrText3, 1)
      blnOK = True

    End If
    sstrText3 = Right(sstrText3, Len(sstrText3) - 1)

  Loop

  If (Not sstrText3 > "") And (Not strChar > "") Then
    MaskNextAlfa = strMaskCompleter

  Else
    MaskNextAlfa = strChar

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function MaskNextAlfa

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo do form
REM -------------------------------------------------------------------------
Public Function PutMask(strText, strMask, strMaskCompleter)
Dim strAux, i

  If strMask > "" Then
    strAux = ""
    sstrText3 = CStr(ZTIReverse(strText))
    For i = Len(strMask) to 1 Step -1
REM          ShowMessageError Mid(strMask, i , 1)
      Select Case Mid(strMask, i , 1)
        Case "0"
          strAux = strAux & MaskNextNumber(strMaskCompleter)

        Case "A"
          strAux = strAux & MaskNextAlfa(strMaskCompleter)

        Case Else
          strAux = strAux & Mid(strMask, i , 1)

      End Select
    Next
    PutMask = ZTIReverse(strAux)

  Else
    PutMask = strText

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function PutMask

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo do form
REM -------------------------------------------------------------------------
Public Function EditShowField(blnFocus, blnFieldOnly, blnDisabled, blnEmpty)
Dim sql, i, Value, blnFirst

  If EditFindField(EditLastOperation) Then
    If Session(EditCurrentField & "ShowOption") > -1 Then
      FormSetTitle Session(EditCurrentField & "Hint")
      If blnEmpty Then
        Value = ""

      Else
        If IsNull(Session(EditCurrentField & "Value")) Then
          Value = ""

        Else
          Value = Session(EditCurrentField & "Value")

        End If
      End If

      Select Case Session(EditCurrentField & "EditOption")
        Case conTextField, conHTTPField, conEMailField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If

          FormInputTextMaxLength Session(EditCurrentField & "Field"), EditSizeToShow, _
          FormatVariable(Value, Session(EditCurrentField & "Type")), EditSizeToGet
 
        Case conFileField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            Response.Write "&nbsp;&nbsp;<FONT COLOR=#000066>" & Value & "</FONT>"
            BR          

          End If
          
          FormInputFile Session(EditCurrentField & "Field"), 0, Value, EditSizeToGet
REM          Response.Write "<SCRIPT LANGUAGE=JavaScript>" & "document.frmEditForm." & Session(EditCurrentField & "Field") & ".defaultValue=""" & Value & """</SCRIPT>"
REM          Response.Write "<SCRIPT LANGUAGE=JavaScript>alert(" & "document.frmEditForm." & Session(EditCurrentField & "Field") & ".defaultValue);</SCRIPT>"
REM          Response.Write "<SCRIPT LANGUAGE=JavaScript>alert(" & "document.frmEditForm." & Session(EditCurrentField & "Field") & ".value);</SCRIPT>"

        Case conPasswordField
          EditShowCaption True
          EditShowFieldImages
          BR
          FormInputPassword Session(EditCurrentField & "Field"), EditSizeToShow

        Case conTextAreaField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If
          FormTextArea Session(EditCurrentField & "Field"), Session(EditCurrentField & "TextAreaHeight"), EditSizeToShow, Value

        Case conUFField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If
            If blnDisabled Then
              FormSelectBegin Session(EditCurrentField & "Field"), 1, False
              If Value > "" Then
                FormSelectOption Value, Value, True

              Else
                FormSelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

              End If
              FormSelectEnd

            Else
              FormSelectUF Session(EditCurrentField & "Field"), 1, False, Value

            End If

        Case conLookupField
          If (Session(EditCurrentField & "LookupTable") > "") And (Session(EditCurrentField & "LookupField") > "")_
          And (Session(EditCurrentField & "LookupFieldText") > "") Then

            If Not blnFieldOnly Then
              EditShowCaption True
              EditShowFieldImages
              BR

            End If

            If blnDisabled Then
              sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")
              sql = sql & " WHERE " & Session(EditCurrentField & "LookupField") & " = " & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))


            Else
              sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")
              sql = sql & " ORDER BY " & Session(EditCurrentField & "LookupFieldText")

            End If

            sobjCMD.CommandText = sql
            sobjCMD.CommandType = adCmdText
REM            ShowMessageError sql & "&" & Value & "&" & (Not Value > "")

            FormSelectBegin Session(EditCurrentField & "Field"), 1, False
            If Not Value > "" And blnDisabled Then
              FormSelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

            Else
              If Not Value > "" Then
                FormSelectOption "", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", True

              End If
              sobjRS2.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly
              Do While Not sobjRS2.EOF
                FormSelectOption sobjRS2(Session(EditCurrentField & "LookupField")), _
                                             sobjRS2(Session(EditCurrentField & "LookupFieldText")), _
                                             (ToString(sobjRS2(Session(EditCurrentField & "LookupField"))) _
                                             = ToString(Value))

                sobjRS2.MoveNext

              Loop
              sobjRS2.Close

            End If
            FormSelectEnd

          Else
            Err.Raise 300, "EditShowField", "O campo """ & Session(EditCurrentField & "Field") & """, da tabela """ & EditCurrentTable & _
            """, está definido como lookup, mas sua definição de lookup não foi feita ou não está completa"

          End If

        Case conCheckField
          If Not sstrCheckAsRadio > "" Then
            FormInputCheckBox Session(EditCurrentField & "Field"), Session(EditCurrentField & ",1Value"),_
            "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",1Value")))
            If Not blnFieldOnly Then
              EditShowCaption True
              EditShowFieldImages

            End If

          Else
            ShowHTMLCR Session(EditCurrentField & ",1Caption")
            FormInputRadio Session(EditCurrentField & "Field"), Session(EditCurrentField & ",1Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",1Value")))
            Nbsp
            ShowHTMLCR Session(EditCurrentField & ",2Caption")
            FormInputRadio Session(EditCurrentField & "Field"), Session(EditCurrentField & ",2Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",2Value")))

          End If

        Case conRadioField
          If Not blnFieldOnly Then
            EditShowCaption True

          End If
          Dim j
          TableBegin "", 1, "100%"
          TableBeginRow 1
          TableBeginColumn
          TableNormalBegin "100%"
          blnFirst = True
          For i = 1 to Session(EditCurrentField & "ValueNumber")
            If ((i - 1) Mod DivRoundUp(Session(EditCurrentField & "ValueNumber"), Session(EditCurrentField & "RadioColumns"))) = 0 Then
              If not blnFirst Then
                TableNormalEnd
                TableEndColumn

              End If
              blnFirst = False
REM              ShowMessageError i
              TableSetColumnVAlign "Top"
              TableBeginColumn
              TableNormalBegin "100%"

            End If
            TableBeginRow 1
            TableSetColumnWidth 10
            TableBeginColumn
            FormInputRadio Session(EditCurrentField & "Field"), Session(EditCurrentField & "," & i & "Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & "," & i & "Value")))
            TableEndColumn
            TableSetColumnWidth ""
            TableBeginColumn
            ShowHTMLCR Session(EditCurrentField & "," & i & "Caption")
            TableEndColumn
            TableEndRow

          Next
          TableNormalEnd
          TableEndColumn
          TableEndRow
          TableNormalEnd
          TableEndColumn
          TableEndRow
          TableEnd

        Case conSelectField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If

          FormSelectBegin Session(EditCurrentField & "Field"), 1, False
          If Not Value > "" Then
            FormSelectOption "", "", True

          End If
          For i = 1 to Session(EditCurrentField & "ValueNumber")
            FormSelectOption Session(EditCurrentField & "," & i & "Value"), Session(EditCurrentField & "," & i & "Caption"), _
                                             (ToString(Value) = ToString(Session(EditCurrentField & "," & i & "Value")))

          Next
          FormSelectEnd

        Case conSeparateDateField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If

          FormSelectBegin Session(EditCurrentField & "Field") & "DayValue", 1, False

          For i = 1 to 31
            FormSelectOption LeadingZeroes(i, 2), LeadingZeroes(i, 2), i = Day(Value)

          Next

          FormSelectEnd

          Response.Write "<B>/</B>"

          FormSelectBegin Session(EditCurrentField & "Field") & "MonthValue", 1, False

          For i = 1 to 12
            FormSelectOption LeadingZeroes(i, 2), LeadingZeroes(i, 2), i = Month(Value)

          Next

          FormSelectEnd

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

          FormSelectBegin Session(EditCurrentField & "Field") & "YearValue", 1, False

          For i = BeginYear to EndYear
            FormSelectOption i, i, i = Year(Value)

          Next

          FormSelectEnd

      End Select

      If blnFocus Then
        blnFocus = False
        ShowHTMLCR "<SCRIPT LANGUAGE=""JavaScript"">"
        ShowHTMLCR "  document.frmEditForm." & Session(EditCurrentField & "Field") & ".focus();"
        ShowHTMLCR "</SCRIPT>"

      End If

    Else
      EditShowFieldAsText

    End If

  Else
    Response.Write "<FONT SIZE=-1>" & EditLastOperation & "</FONT>"

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
REM Retorna o valor 1 se o parametro passado for verdadeiro
REM -------------------------------------------------------------------------
Private Function ToString(strText)

  If strText > "" Then
    ToString = CStr(strText)

  Else
    ToString = ""

  End if

End Function
REM -------------------------------------------------------------------------
REM Final da Function BitBln

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo como texto
REM -------------------------------------------------------------------------
Public Sub EditShowFieldAsText
Dim sql, i

  If EditFindField(EditLastOperation) Then
    Response.Write "<FONT SIZE=-1>"
    Select Case Session(EditCurrentField & "EditOption")
      Case conTextField, conUFField, conTextAreaField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write "&nbsp;:&nbsp;"

        If (Session(EditCurrentField & "ExternalLink") > "") Then
          EditBeginMakeLink True

        End If

        Response.Write PutMask(FormatVariable(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter"))
        EditEndMakeLink

      Case conHTTPField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write "&nbsp;:&nbsp;"
        Response.Write "<A HRef=http://" & Session(EditCurrentField & "Value") & ">" & Session(EditCurrentField & "Value") & "</A>"

      Case conEMailField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write "&nbsp;:&nbsp;"
        Response.Write "<A HRef=mailto:" & Session(EditCurrentField & "Value") & ">" & Session(EditCurrentField & "Value") & "</A>"

      Case conLookupField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Response.Write "&nbsp;:&nbsp;"
        If Session(EditCurrentField & "Value") > "" Then
          sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")
          sql = sql & " WHERE " & Session(EditCurrentField & "LookupField") & "="
          sql = sql & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))
                    
          sobjCMD.CommandText = sql

REM      ShowMessageError sql

          sobjRS2.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

          If Not sobjRS2.EOF And Not sobjRS2.BOF Then
            EditBeginMakeLink True
            Response.Write sobjRS2(Session(EditCurrentField & "LookupFieldText"))
            EditEndMakeLink

          End If
          sobjRS2.Close

        End If

      Case conCheckField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Nbsp
        EditBeginMakeLink True
        If Session(EditCurrentField & "Value") = Session(EditCurrentField & ",1Value") Then
          ShowHTMLCR Session(EditCurrentField & ",1Caption")

        Else
          ShowHTMLCR Session(EditCurrentField & ",2Caption")

        End If
        EditEndMakeLink

      Case conRadioField, conSelectField
        Response.Write "<B>" & EditCaption(False) & "</B>"
        Nbsp
        EditBeginMakeLink True
        For i = 1 to Session(EditCurrentField & "ValueNumber")
          If Session(EditCurrentField & "Value") = Session(EditCurrentField & "," & i & "Value") Then
            ShowHTMLCR Session(EditCurrentField & "," & i & "Caption")

          End If

        Next
        EditEndMakeLink

    End Select
    ShowHTMLCR "</FONT>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowFieldAsText

REM -------------------------------------------------------------------------
REM Rotina que apresenta um campo como texto
REM -------------------------------------------------------------------------
Public Function EditFieldValue
Dim sql, Value, i

  Select Case Session(EditCurrentField & "EditOption")
    Case conTextField, conUFField, conTextAreaField
REM      ShowMessageError Session(EditCurrentField & "Field")
      Value = sobjRS(Session(EditCurrentField & "Field"))

    Case conHTTPField
      Value = "<A HRef=" & sobjRS(Session(EditCurrentField & "Field")) & ">" & sobjRS(Session(EditCurrentField & "Field")) & "</A>"

    Case conEMailField
      Value = "<A HRef=mailto" & sobjRS(Session(EditCurrentField & "Field")) & ">" & sobjRS(Session(EditCurrentField & "Field")) & "</A>"

    Case conLookupField
      Value = sobjRS(Session(EditCurrentField & "LookupFieldText"))

    Case conCheckField
      If sobjRS(Session(EditCurrentField & "Field")) = Session(EditCurrentField & ",1Value") Then
        Value = Session(EditCurrentField & ",1Caption")

      Else
        Value = Session(EditCurrentField & ",2Caption")

      End If

    Case conRadioField, conSelectField
      For i = 1 to Session(EditCurrentField & "ValueNumber")
        If sobjRS(Session(EditCurrentField & "Field")) = Session(EditCurrentField & "," & i & "Value") Then
          Value = Session(EditCurrentField & "," & i & "Caption")

        End If

      Next

  End Select

  EditFieldValue = Value

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditShowFieldAsText

REM -------------------------------------------------------------------------
REM Formata uma variavel de acordo com seu tipo
REM -------------------------------------------------------------------------
Public Function FormatVariable(Value, strType)
REM  ShowMessageError Value & "&" & strType
  If Value > "" Then
    Select Case strType
      Case "int"
        FormatVariable = ZTIInt(Value)

      Case "money"
        FormatVariable = ZTIMoney(Value)

      Case "float"
        FormatVariable = ZTINumber(Value)

      Case "datetime"
        FormatVariable = ZTILongDate(Value)

      Case Else
        FormatVariable = Value

    End Select

  Else
    FormatVariable = Null

  End If
End Function
REM -------------------------------------------------------------------------
REM Final da Function FormatVariable

REM -------------------------------------------------------------------------
REM Formata uma variavel de acordo com seu tipo para ser usada numa chamada SQL
REM -------------------------------------------------------------------------
Public Function FormatForSQL(ByVal Value, ByVal strType)
REM  ShowMessageError Value & "&" & strType
  If Value > "" Then
    Select Case LCase(strType)
      Case "int"
        FormatForSQL = NormalizeString(Value)

      Case "varchar"
        FormatForSQL = SQLString(Value)

      Case "date"
        FormatForSQL = SQLLongDateBR(Value)

      Case Else
        FormatForSQL = Value

    End Select

  Else
    FormatForSQL = Null

  End if
End Function
REM -------------------------------------------------------------------------
REM Final da Function FormatForSQL

REM -------------------------------------------------------------------------
REM Rotina que monta a listagem
REM -------------------------------------------------------------------------
Public Sub EditShowAZ
Dim sql, i, strAux, strAux2, strTip, strTip2, arrFields(30), intCount

  Session("LastToFind") = conOptionAZEdit

  TableSetColumnNoWrap False

  sql = "SELECT * FROM " & EditCurrentTable & Session(EditCurrentTable & "JoinFormUnit")

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    If Session(EditCurrentTable & i & "EditOption") = conLookupField Then
       sql = sql & " LEFT OUTER JOIN " & Session(EditCurrentTable & i & "LookupTable") &_
       " on " & Session(EditCurrentTable & i & "LookupTable") & "." & Session(EditCurrentTable & i & "LookupField") & " = " & _
       EditTableOf(i) & "." & Session(EditCurrentTable & i & "Field")

    End If

  Next

  If sparEditFilter > "" Then
    sql = sql & " WHERE " & sparEditFilter

  End If

  If Not Session("OrderField") > "" Then
    EditBeginReadOperation Session(EditCurrentTable & "FieldList")

    Do While EditIsThereOperation
      If EditFindField(EditNextOperation) Then
        sstrText = ""
        Session("OrderField") = Session("CurrentField")

      End If

    Loop

  End If

  Dim blnFirst
  EditBeginReadOperation Session("OrderField")
  sql = sql & " ORDER BY "

  blnFirst = True

  Do While EditIsThereOperation
    Select Case EditNextOperation
       Case ","

       Case Else
         Session("CurrentField") = EditLastOperation
         If Not blnFirst Then
           sql = sql & ", "

         Else
           sql = sql & " "

         End If

         Select Case Session(EditCurrentField & "EditOption")
           Case conLookUpField
             sql = sql & Session(EditCurrentField & "LookupTable") & "." & Session(EditCurrentField & "LookupFieldText")

           Case Else
             sql = sql & Session(EditCurrentField & "Field")

        End Select

        blnFirst = False
        If Session("OrderDesc") = "1" Then
          sql = sql & " DESC "

        End If

    End Select

  Loop

  BrowseSetActual EditCurrentTable

  Nbsp

  sobjCMD.CommandText = sql
  sobjCMD.CommandType = adCmdText
REM  ShowMessageError sql

  sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

  If Not sobjRS.EOF And Not sobjRS.BOF Then
    Dim strMessage

    strMessage = "Lista de " & Session(EditCurrentTable & "Caption")

    Center

    BrowseBegin
    BrowseHeader

    EditBeginReadOperation Session(EditCurrentTable & "FieldList")

    Dim Width, strList

    Width = 0
    Do While EditIsThereOperation
      Select Case EditNextOperation
         Case ",", ";", "-"
         Case "="
           EditIsThereOperation

         Case Else
           If EditFindField(EditLastOperation) Then
             If Session(EditCurrentField & "Type") <> "varchar" Then
               Width = Width + EditSize + 3

             Else
               Width = Width + EditSize * 0.8

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

    TableBegin strMessage, Session(EditCurrentTable & "ListOperationNumber") + 1, Width & "%"

REM    ShowMessageError Session(EditCurrentTable & "ListOperationNumber")

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
              If EditFindField(EditLastOperation) Then
                strAux = strAux & EditCaption(False)
                If strAux2 > "" Then
                  strAux2 = strAux2 & ","

                End If
                strAux2 = strAux2 & Session("CurrentField")

              End If

          End Select

        Loop
REM        ShowMessageError strAux2
        If Session("OrderField") = strAux2 Then
          strAux2 = Server.URLEncode(strAux2)
          If Session("OrderDesc") = "1" Then
            strTip = strTip & "<IMG SRC=""/ZTITools/img/EditForm/OrderDes.gif"" BORDER=0>" & _
            "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
            EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & GetFilter & """>" & _
            "&nbsp;" & strAux & "</A><BR>"

          Else
            strTip = strTip & "<IMG SRC=""/ZTITools/img/EditForm/OrderAsc.gif"" BORDER=0>" & _
            "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
            EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & conPEditOrderDescStr & GetFilter & """>" & _
             "&nbsp;" & strAux & "</A><BR>"

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

    TableSetColumnColSpan Session(EditCurrentTable & "ListOperationNumber") + 1

    TableBeginRow 1
    TableBeginColumn

    TableNormalBegin "100%"
    TableSetColumnAlign "Left"
    TableSetColumnVAlign "Top"
    TableBeginRow 1
    TableBeginColumn
    Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd & _
               EditWhatList & conPEditEditableStr & ">Incluir</a></B>"
    Nbsp
    Nbsp
    Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
               EditWhatList & conPEditEditableStr & GetFilter & ">Editar</a></B>"

    If Session(EditCurrentTable & "FieldFind") > "" Then
      Nbsp
      Nbsp
      Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindAdvanced & _
               EditWhatList & ">Localizar</a></B>"
    Else
      Nbsp
      Nbsp
      Response.Write "<B><a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindCommon & _
                 EditWhatList & ">Localizar</a></B>"

    End If
    TableEndColumn

    If Session(EditCurrentTable & "ListNumber") > 0 Then
      TableSetColumnAlign "Right"
      TableSetColumnVAlign "Top"
      TableBeginColumn
      ShowHTMLCR "<FONT SIZE=2><DIV ID=F2 class=SS onMouseOver=""makeCool(F2);"" onMouseOut=""makeNormal(F2);"" onClick=""OpenDesc(T2);"" STYLE=""cursor:hand""><B>Outras listagens</B></DIV>"
      ShowHTMLCR "<table><tr><td align=justify>"
      ShowHTMLCR "<DIV ID=T2 class=dbody style=""display:'none';"">"
      Response.Write strTip2
      ShowHTMLCR "</DIV></td></tr></table>"
      TableEndColumn

    End If

    If Session(EditCurrentTable & "OrderNumber") > 0 Then
      TableSetColumnAlign "Right"
      TableSetColumnVAlign "Top"
      TableBeginColumn
      ShowHTMLCR "<FONT SIZE=2><DIV ID=F1 class=SS onMouseOver=""makeCool(F1);"" onMouseOut=""makeNormal(F1);"" onClick=""OpenDesc(T1);"" STYLE=""cursor:hand""><B>Outras ordenações</B></DIV>"
      ShowHTMLCR "<table><tr><td align=justify>"
      ShowHTMLCR "<DIV ID=T1 class=dbody style=""display:'none';"">"
      Response.Write strTip
      ShowHTMLCR "</DIV></td></tr></table>"
      TableEndColumn

    End If
    TableEndRow
    TableNormalEnd
    TableEndColumn
    TableEndRow
    TableSetColumnAlign ""
    TableSetColumnColSpan ""
    TableBeginHeadRow 2

    EditBeginReadOperation strList

Dim intNumHeader, intNumArrHeaders, intAux

    intNumHeader     = 0
    intCount         = 1
    intNumArrHeaders = 0

    If Session(EditCurrentTable & "HeaderNumber") > 0 Then
      TableSetColumnRowSpan 2

    End If

    Do While EditIsThereOperation
      Select Case EditNextOperation
        Case ",",";","-","="

        Case Else
          Session("CurrentField") = EditLastOperation
          intNumHeader = intNumHeader + 1
          REM ShowMessageError Session(EditCurrentTable & "HeaderNumber")
          If Session(EditCurrentTable & "HeaderNumber") > 0 Then
REM            ShowMessageError "&" & intNumHeader & "=" & Session(EditCurrentTable & intCount & "HeaderBegin")
            If intNumHeader = Session(EditCurrentTable & intCount & "HeaderBegin") Then
              TableSetColumnRowSpan ""
              TableSetColumnColSpan Session(EditCurrentTable & intCount & "HeaderWidth")
              TableHeadColumn Session(EditCurrentTable & intCount & "HeaderText")
              TableSetColumnRowSpan 2
              TableSetColumnColSpan ""
              intAux           = 1
              intNumArrHeaders = intNumArrHeaders + 1
              arrFields(intNumArrHeaders) = EditLastOperation
              If Session(EditCurrentTable & "HeaderNumber") > intCount Then
                intCount = intCount + 1

              End If

            Else
              If intAux > 0 And intAux < Session(EditCurrentTable & intCount & "HeaderWidth") Then
                intNumArrHeaders = intNumArrHeaders + 1
                arrFields(intNumArrHeaders) = EditLastOperation
                intAux = intAux + 1

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
      TableSetColumnRowSpan ""
      TableEndHeadRow
      TableBeginHeadRow 2
      For i = 1 to intNumArrHeaders
        Session("CurrentField") = arrFields(i)
        EditTableHeader EditCaption(False), Session("CurrentField"), Session(EditCurrentField & "ListSize")

      Next

    End If

    TableEndHeadRow

    Do While Not sobjRS.EOF And IsBrowsePage
      EditBeginReadOperation strList

      TableBeginRow 1

REM      intCount = 0

      Do While EditIsThereOperation
        Select Case EditNextOperation
          Case ",",";","-","="

          Case Else
REM            intCount = intCount + 1
REM            If arrFields(intCount) <> Session(EditCurrentField & "Value") Then
              Session("CurrentField") = EditLastOperation
                If TypeName(EditFieldValue) = "Long" Or TypeName(EditFieldValue) = "Integer" Or _
                  TypeName(EditFieldValue) = "Currency" Or TypeName(EditFieldValue) = "Double" Or TypeName(EditFieldValue) = "Single" Then
                  TableSetColumnAlign "Right"

                End If

                arrFields(intCount) = Session(EditCurrentField & "Value")

                TableBeginColumn
                EditBeginMakeLink False
                Select Case Session(EditCurrentField & "EditOption")
                  Case conLookupField, conRadioField, conCheckField, conSelectField
                    ShowHTMLCR BreakString(EditFieldValue, Session(EditCurrentField & "ListChars"))

                  Case Else
                    ShowHTMLCR BreakString(PutMask(FormatVariable(EditFieldValue, Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars"))

                End Select
                TableSetColumnAlign ""
                EditEndMakeLink

                TableEndColumn

REM             Else
REM              TableBeginColumn
REM              EditBeginMakeLink False
REM              ShowHTMLCR PutMask(FormatVariable(EditFieldValue, Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter"))
REM              TableSetColumnAlign ""
REM              EditEndMakeLink

REM              TableEndColumn

REM            End If

        End Select

      Loop
      TableEndRow

      sobjRS.MoveNext
      BrowseLineCounter
    Loop

    TableEnd

    BrowseSetStatus True
    BrowseEnd
    BrowseBar

    CenterEnd

  Else
    ShowMessageError "<center>Não há detalhamento na Tabela " & Session(EditCurrentTable &_
    "Caption") & " para as opções escolhidas !!!</center><P>"
    ShowHTMLCR "&nbsp;<a href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd &_
               "&" & EditWhatList &_
               ">Incluir um registro</a><P>"

  End If

  BR
  sobjRS.Close

  EditFPLPLink

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowEditor

REM -------------------------------------------------------------------------
REM Monta um header de listagem
REM -------------------------------------------------------------------------
Private Sub EditTableHeader(strField, strOrderField, intSize)
  TableSetColumnWidth intSize
  TableBeginHeadColumn
  If LCase(CStr(Session("OrderField"))) = LCase(CStr(strOrderField)) Then
    If Session("OrderDesc") = "1" Then
      ShowImage "/ZTITools/img/EditForm/OrderDes.gif", "", "Ordenado descendentemente pelo campo " & strField, sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & GetEditable & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetList & GetFilter, "", 8, 7
      Nbsp
      Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & GetList & GetFilter & """>" & strField & "</A>"

    Else
      ShowImage "/ZTITools/img/EditForm/OrderAsc.gif", "", "Ordenado ascendentemente pelo campo " & strField, sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & GetEditable & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & conPEditOrderDescStr & GetList & GetFilter, "", 8, 7
      Nbsp
      Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & conPEditOrderDescStr & GetList & GetFilter & """>" & strField & "</A>"

    End If
  Else
    Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & sparEditOption & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrderField) & GetEditable & GetList & GetFilter & """>" & strField & "</A>"

  End If
  TableEndHeadColumn

End Sub
REM -------------------------------------------------------------------------
REM Final da Function EditTableHeader

Private Function EditListStr
 EditListStr = "&" & conPEditList & "=" & sparEditList

End Function

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
    If sobjRS.State = 0 Then
      strAux = strAux & Session(EditCurrentField & "," & i & "ILTarget") & "=" & FormatForSQL(Session(EditCurrentField & "Value"), Session(EditCurrentField & "Type"))

    Else
      intCurrentField = Session("CurrentField")
      EditFindField Session(EditCurrentField & "," & i & "ILTarget")
REM      ShowMessageError Session(EditCurrentField & "," & i & "ILTarget")
      strAux = strAux & Session(EditCurrentField & "Field") & "=" & FormatForSQL(sobjRS(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"))
      Session("CurrentField") = intCurrentField

    End If

  Next
  If strAux > "" Then
    EditTargetOfLinks = "&" & conPEditFind & "=" & Server.URLEncode(strAux)

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditMakeInternalLink

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
REM      ShowMessageError Session(EditCurrentTable & i & "OldValue") & "ksjhgfsk"
REM      ShowMessageError Session(EditCurrentTable & i & "Field") & EditCurrentTable & i & "Field"

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
    ShowHTMLCR "</A>"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditEndMakeLink

REM -------------------------------------------------------------------------
REM Armazena os campos da tabela do form em variaveis
REM -------------------------------------------------------------------------
Private Sub EditGetForm
Dim i

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")
  Dim Width

  Width = 0
  Do While EditIsThereOperation
    Select Case EditNextOperation
       Case ",", ";", "-", ""
       Case "="
         EditIsThereOperation

       Case Else
         If EditFindField(EditLastOperation) Then
           Select Case Session(EditCurrentField & "EditOption")
             Case conSeparateDateField
                Session(EditCurrentField & "Value") = ImportDate(Request.Form(Session(EditCurrentField & "Field") & _
                                                "DayValue") & Request.Form(Session(EditCurrentField & "Field") & _
                                                "MonthValue") & Request.Form(Session(EditCurrentField & "Field") &_
                                                "YearValue"), conDateTypeDDMMAAAA)

              Case Else
REM                Session(EditCurrentField & "Value") = EditFormatField(Request.Form(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"), Session(EditCurrentField & "Size"))
                Session(EditCurrentField & "Value") = Request.Form(Session(EditCurrentField & "Field"))
REM                ShowMessageError Request.Form(Session(EditCurrentField & "Field")) & "&" & Session(EditCurrentField & "Field")

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

  If Session("RecordAction") = "1" Then
    sobjCMD.CommandText = EditCurrentTable
    sobjCMD.CommandType = adCmdTable

    sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

    sobjRS.AddNew

    strAnotherTable = ""
    For i = 1 to Session(EditCurrentTable & "FieldNumber")
      If Not Session(EditCurrentTable & i & "ExternalAddress") > "" Then
REM        ShowMessageError Session(EditCurrentTable & i & "Field") & "&" & Session(EditCurrentTable & i & "Value")
REM        ShowMessageError Session(EditCurrentTable & (i+1) & "Field") & "&" & Session(EditCurrentTable & (i+1) & "Value")
        sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatField(Session(EditCurrentTable & _
                                                                 i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                 Session(EditCurrentTable & i & "Size"))
        Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

      Else
        If Not strAnotherTable > "" Then
          strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

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
          If EditFindField(Session(strTable & "," & j & "RelationField1")) Then
            sobjRS(Session(strTable & "," & j & "RelationField2")) = EditFormatField(Session(EditCurrentField & "Value"), _
                                                                     Session(EditCurrentField & "Type"), EditSize)
            Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

          End If
        Next

      End If

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "OwnerTable") = strCurrentTable Then
REM           ShowMessageError Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "Value")
           sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatField(Session(EditCurrentTable & _
                                                                    i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                    Session(EditCurrentTable & i & "Size"))
           Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

        Else
          If (Not strAnotherTable > "") And (InStr(strSaved, Session(EditCurrentTable & i & "OwnerTable")) = 0) Then
            strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

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

REM    ShowMessageError sql

    sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

    strAnotherTable = ""
    For i = 1 to Session(EditCurrentTable & "FieldNumber")
      If Not Session(EditCurrentTable & i & "ExternalAddress") > "" Then
        sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatField(Session(EditCurrentTable & _
                                                                 i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                 Session(EditCurrentTable & i & "Size"))
        Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

      Else
        If Not strAnotherTable > "" Then
          strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

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
          If EditFindField(Session(strTable & "," & j & "RelationField1")) Then
            sql = sql & PutAnd & Session(strTable & "," & j & "RelationField2") & " = " & _
                      FormatForSQL(Session(EditCurrentField & "OldValue"), Session(EditCurrentField & "Type"))

          End If
        Next

      End If

      sobjCMD.CommandText = sql
      sobjCMD.CommandType = adCmdText

REM      ShowMessageError sql

      sobjRS.Open sobjCMD, , adOpenDynamic, adLockOptimistic

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "OwnerTable") = strCurrentTable Then
            sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatField(Session(EditCurrentTable & _
                                                                     i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                     Session(EditCurrentTable & i & "Size"))
            Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

        Else
          If (Not strAnotherTable > "") And (InStr(strSaved, Session(EditCurrentTable & i & "OwnerTable")) = 0) Then
            strAnotherTable   = Session(EditCurrentTable & i & "OwnerTable")

          End If

        End If
      Next

      If strTable > "" Then
        For j = 1 to Session(strTable & "FieldtoFieldNumber")
          If EditFindField(Session(strTable & "," & j & "RelationField1")) Then
            sobjRS(Session(strTable & "," & j & "RelationField2")) = EditFormatField(Session(EditCurrentField & "Value"), _
                                                                     Session(EditCurrentField & "Type"), EditSize)
            Session(EditCurrentTable & i & "OldValue") = Session(EditCurrentTable & i & "Value")

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

REM    ShowMessageError sql

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    sobjRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

  Else
    sql = "DELETE FROM " & strTable

    If sparEditFind > "" Then
      sql = sql & " WHERE " & sparEditFind

    End If

REM    ShowMessageError sql & "<BR>"

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    sobjRS.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditDelete

REM -------------------------------------------------------------------------
REM Prepara datas MMDDAA
REM -------------------------------------------------------------------------
Public Function ZTILongDateBR(datDate)
  Dim strAux
  Dim strYear, strMonth, strDay

  If VarType(datDate) = vbString Then
    datDate = NormalizeString(datDate)

    strYear  = Mid(datDate, 5, 4)
    strMonth = Mid(datDate, 3, 2)
    strDay   = Mid(datDate, 1, 2)

    If Len(datDate) = 6 Then
      strYear  = "19" & Mid(datDate, 5, 2)

    End If

    datDate = strMonth & "/" & strDay & "/" & strYear

  End If

	ZTILongDateBR =  datDate

End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTILongDateBR

REM -------------------------------------------------------------------------
REM Prepara Money para serem utilizadas dentro de uma Query SQL
REM -------------------------------------------------------------------------
Public Function ImportDouble(strValue)
dim target
dim i
dim cOneChar
dim intPos, monAux

  If Trim(strValue) > "" Then
  	target = ""
  	strValue = UCase(Trim(strValue))

  	For i = 1 to Len(strValue)
  		cOneChar = Mid(strValue, i, 1)
  		If cOneChar = "." Or cOneChar = "," Or cOneChar = "-" Or (cOneChar >= "0" and cOneChar <= "9") then
        target = target & cOneChar

  		End if

  	Next

    strValue = target

    intPos = InStr(strValue, ",")

    monAux = 0

    If intPos > 0 Then
      monAux = CLng(Left(strValue, intPos))
        monAux = monAux + (CLng(Mid(strValue, intPos + 1)) / Exp((Len(strValue) - intPos)*Log(10)))

    Else
      monAux = CLng(strValue)

    End If


    ImportDouble = monAux

 Else
    ImportDouble = Null

 End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function ZTILongDateBR

REM -------------------------------------------------------------------------
REM Formata uma variavel
REM -------------------------------------------------------------------------
Public Function EditFormatField(ByVal Value, ByVal strType, ByVal intSize)
  If Value > "" Then
    Select Case LCase(strType)
      Case "int"
        EditFormatField = CLng(TotalNormalizeString(Value))

      Case "datetime"
        If IsValidDate(Value) Then
          EditFormatField = CDate(ZTILongDateBR(Value))

        Else
          EditFormatField = Value

        End If

      Case "varchar"
        EditFormatField = Left(CStr(Value), intSize)

      Case "money"
        EditFormatField = ImportDouble(Value)

      Case "float"
        EditFormatField = ImportDouble(Value)

      Case Else
        EditFormatField = Value

      End Select

  Else
    EditFormatField = Null

  End If
End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormatField

REM -------------------------------------------------------------------------
REM Rotina que emite uma Mensagem
REM -------------------------------------------------------------------------
Private Sub Message(ByVal strCabecalho, strTexto, strRetorno)

  BR
  Center

  FormBegin strRetorno, "POST", "frmMensagem", "", "40%"

  TableBegin strCabecalho, 1, ""

  TableBeginRow 2
  TableBeginColumn
  Center
  ShowHTMLCR strTexto
  TableEndColumn
  TableEndRow

  TableBeginRow 2
  TableBeginColumn
  Center
  FormInputSubmit "btmVolta", constrOK
  TableEndColumn
  TableEndRow

  TableEnd

  FormEnd

  CenterEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub Mensagem

REM -------------------------------------------------------------------------
REM Redireciona o cliente para o form ou para a opcao do parametro
REM -------------------------------------------------------------------------
Private Sub MakeServerValidation(strRedirectOption)
  If ServerValidates Then
    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & strRedirectOption & _
    EditWhatWhereHow

  Else
    Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
    EditWhatWhereHow

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MakeServerValidation

REM -------------------------------------------------------------------------
REM Retorna o HTML que monta a imagem de erro
REM -------------------------------------------------------------------------
Private Function ErrorImage

  ErrorImage = "<IMG SRC=/ztitools/img/editform/CampoErrado.gif width=5 Height=5>&nbsp;"

End Function
REM -------------------------------------------------------------------------
REM Final da Function ErrorImage

REM -------------------------------------------------------------------------
REM Rotina que faz as validacoes no servidor
REM -------------------------------------------------------------------------
Private Function ServerValidates
Dim i, blnIsValid

  blnIsValid = True

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditLastOperation) Then
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
  REM    ShowHTMLCR EditLastOperation

    End Select

  Loop

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditLastOperation) Then
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
      If Session(strField & "Value") > "" Then
        If Not IsValidDate(Session(strField & "Value")) Then
          blnIsValid = False

        End If
      End If

    Case conValidSepDate
REM      FormJavaValidSepDate Session(EditCurrentTable & i & "Field") & "DayValue", Session(EditCurrentTable & i & "Field") & "MonthValue", Session(EditCurrentTable & i & "Field") & "YearValue", "V" & Session(EditCurrentTable & i & "Field"), False, ""

    Case conValidDateMToday
      If IsValidDate(Session(strField & "Value")) Then
        If CDate(ZTILongDateBR(Session(strField & "Value"))) > Now + 1 Then
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

  If Not blnIsValid Then
    EditFindField(strField)
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
  blnStrExt = True

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditLastOperation) Then
          If blnStrExt Then
            StrExtBegin
            blnStrExt = False

          End If
          Select Case Session(EditCurrentField & "ValidationOption")
            Case conValidCGC
              FormJavaValidCGC Session(EditCurrentField & "Field"), "V" & Session(EditCurrentField & "Field")

            Case conValidCPF
              FormJavaValidCPF Session(EditCurrentField & "Field"), "V" & Session(EditCurrentField & "Field")

            Case conValidDate
              FormJavaValidDate Session(EditCurrentField & "Field"), "V" & Session(EditCurrentField & "Field"), False, ""

            Case conValidSepDate
              FormJavaValidSepDate Session(EditCurrentField & "Field") & "DayValue", Session(EditCurrentField & "Field") & "MonthValue", Session(EditCurrentField & "Field") & "YearValue", "V" & Session(EditCurrentField & "Field"), False, ""

            Case conValidDateMToday
              FormJavaValidDate Session(EditCurrentField & "Field"), "V" & Session(EditCurrentField & "Field"), True, EditCaption(False) & " deve ser maior que hoje"

            Case conValidSepDateMToday
              FormJavaValidSepDate Session(EditCurrentField & "Field") & "DayValue", Session(EditCurrentField & "Field") & "MonthValue", Session(EditCurrentField & "Field") & "YearValue", "V" & Session(EditCurrentField & "Field"), True, EditCaption(False) & " deve ser maior que hoje"

            Case conValidEmail
              FormJavaValidEmail Session(EditCurrentField & "Field"), "V" & Session(EditCurrentField & "Field")

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

  FormJavaBegin "ValidateForm", "frmEditForm"

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(EditLastOperation) Then
          If Session(EditCurrentField & "RequireMessage") > "" Then
            FormJavaRequired Session(EditCurrentField & "Field"), Session(EditCurrentField & "RequireMessage")

          ElseIf Not Session(EditCurrentField & "Null") Then
            FormJavaRequired Session(EditCurrentField & "Field"), "O campo " &  EditCaption(False) & " deve ser preenchido"

          End If
        End If

    End Select

  Loop
  If sstrLastValidate > "" Then
    ShowHTMLCR "return (" & sstrLastValidate & ");"

  Else
    ShowHTMLCR "return true;"

  End If

  Response.Write "}"
  Response.Write "//-->"
  ShowHTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ClientValidates

REM -------------------------------------------------------------------------
REM Monta um Grid na
REM -------------------------------------------------------------------------
Private Sub EditShowGrid(strGrid)
Dim sql, i

  EditTotalOperation strGrid, "GridOperationNumber", False, ""

  FormBegin sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionExeEdition, "POST", "frmEditForm", "", ""

  sql = "SELECT * FROM " & EditCurrentTable

  sobjCMD.CommandText = sql
  sobjCMD.CommandType = adCmdText

REM  ShowMessageError sql

  sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

  If Not sobjRS.EOF And Not sobjRS.BOF Then
    Dim strMessage

    strMessage = Session(EditCurrentTable & "Caption")

    TableBegin strMessage, Session(EditCurrentTable & "GridOperationNumber") + 1, "60%"

    TableSetColumnAlign ""
    TableSetColumnColSpan ""
    TableBeginHeadRow 1

    EditBeginReadOperation strGrid

    Do While EditIsThereOperation
      Select Case EditNextOperation
         Case ",", ";", "-", ""
         Case "="
           EditIsThereOperation

         Case Else
           If EditFindField(EditLastOperation) Then
             TableHeadColumn EditCaption(False)

           End If

       End Select

    Loop
    TableEndHeadRow

    Do While Not sobjRS.EOF
      EditBeginReadOperation strGrid

      TableBeginRow 1

      Do While EditIsThereOperation
        Select Case EditNextOperation
          Case ",", ";", "-", ""
          Case "="
            EditIsThereOperation

          Case Else
            If EditFindField(EditLastOperation) Then
              If TypeName(EditFieldValue) = "Long" Or TypeName(EditFieldValue) = "Integer" Or _
                 TypeName(EditFieldValue) = "Currency" Or TypeName(EditFieldValue) = "Double" Or TypeName(EditFieldValue) = "Single" Then
                TableSetColumnAlign "Right"

              End If

              TableBeginColumn
              EditBeginMakeLink False
              ShowHTMLCR FormatVariable(EditFieldValue, Session(EditCurrentField & "Type"))
              TableSetColumnAlign ""
              EditEndMakeLink

              TableEndColumn

            End If

        End Select

      Loop
      TableEndRow

      sobjRS.MoveNext
    Loop

    TableEnd

    FormEnd

  End If

  sobjRS.Close

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditExeEdition

REM -------------------------------------------------------------------------
REM Inicia a criacao do URL - 19/01/99 22:28
REM -------------------------------------------------------------------------
Public Sub EditCreateURLBegin(ByVal strLocate)

  sstrCreateURL = strLocate

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLBegin

REM -------------------------------------------------------------------------
REM Termina a criacao do URL - 19/01/99 22:28
REM -------------------------------------------------------------------------
Public Sub EditCreateURLEnd

  sstrCreateURL = ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLEnd

REM -------------------------------------------------------------------------
REM Determina o modo da edicao - 19/01/99 22:28
REM -------------------------------------------------------------------------
Public Sub EditCreateURLModeSet(ByVal strMode)

  sstrCreateURLMode = strMode

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLModeSet

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit - Ruben 1/18/99 12:57AM
REM -------------------------------------------------------------------------
Public Sub EditCreateURLTarget(ByVal strTarget, ByVal strOption, ByVal strDescription)

  EditCreateURLModeSet "Null"

  sstrCreateURLTarget = strTarget

  EditCreateURL strOption, strDescription

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLModeSet

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit - Ruben 1/18/99 12:57AM
REM -------------------------------------------------------------------------
Public Sub EditCreateURL(ByVal strOption, ByVal strDescription)

  If sstrCreateURL > "" Then
    URLBegin sstrCreateURL

  Else
    URLBegin sstrThisScriptName

  End If

  URLAdd conPEditOption, conOptionFPEdit
  URLAdd conPEditWhat, strOption

  If sstrCreateURLTarget > "" Then
    URLAdd "F1", sstrCreateURLTarget

  End If

  If sstrCreateURLMode > "" Then
    If LCase(sstrCreateURLMode) = "null" Then
      URLShow "", strDescription

    Else
      URLShow sstrCreateURLMode, strDescription

    End If
  Else
    URLShow conPEditEditableStr, strDescription

  End If

  URLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURL

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
REM Retorna filtro se tiver filtro no instante
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
REM Retorna filtro se tiver filtro no instante
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

Private Sub ClassAss
  ShowHTMLCR "<STYLE>"
  ShowHTMLCR "A.SS {text-decoration: none; color: black; font-size: 8pt}"
  ShowHTMLCR "A:HOVER.SS {color: purple}"
  ShowHTMLCR "</STYLE>"

  ShowHTMLCR "<script type=""text/javascript"">"
  ShowHTMLCR "function OpenDesc(text) {"
  ShowHTMLCR "  if (text.style.display =="""") {"
  ShowHTMLCR "    text.style.display = ""none"";"
  ShowHTMLCR "  }"
  ShowHTMLCR "  else {"
  ShowHTMLCR "  text.style.display="""";"
  ShowHTMLCR "  }"
  ShowHTMLCR "}"
  ShowHTMLCR "function makeCool(text) {"
  ShowHTMLCR "  text.style.color = ""Purple"""
  ShowHTMLCR "}"
  ShowHTMLCR "function makeNormal(text) {"
  ShowHTMLCR "  text.style.color = ""Black"""
  ShowHTMLCR "}"
  ShowHTMLCR "</script>"

End Sub

REM -------------------------------------------------------------------------
REM Recolhe as informacoes de um resumo
REM -------------------------------------------------------------------------
Private Sub EditSummaryPars(strText)
Dim blnParameter, i, j

  blnParameter = False

  Session("FieldNumber") = 0
  Session("TotalNumber") = 0
  sblnHaveSum   = False
  sblnHaveGraph   = False

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
              sblnHaveSum                      = True
              sblnHaveGraph                      = True

            Case "gc"
              Session(Session("FieldNumber") & "IsGraph")  = "1"
              Session(Session("FieldNumber") & "IsCount")  = "1"
              Session(Session("FieldNumber") & "OrderSQL") = "COUNT(*)"
              sblnHaveGraph                      = True

            Case "a"
              sstrGroupField = Session(Session("FieldNumber") & "ShowField")

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
          If EditFindField(EditLastOperation) Then
            Session(Session("FieldNumber") & "ShowField") = Session("CurrentField")
            If Session(EditCurrentField & "EditOption") = conLookupField Then
              Session(Session("FieldNumber") & "Field")   = Session(EditCurrentField & "LookupFieldText")

            Else
              Session(Session("FieldNumber") & "Field")   = Session(EditCurrentField & "Field")
            
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
REM   ShowMessageError Session(i & "Field") & Session(i & "Operation")
    If Not blnFirst And Not Session(i & "Operation") > "" Then
      sstrFields = sstrFields & ","

    ElseIf Not Session(i & "Operation") > "" Then
      blnFirst    = False
      
    End If  

    If Not blnFirst2 And Not Session(i & "IsCalcField") > "" Then
      sstrFields2 = sstrFields2 & ","

    ElseIf Not Session(i & "IsCalcField") > "" Then
      blnFirst2    = False
      
    End If  

    If Session(i & "IsSum") > "" Then
      sstrFields = sstrFields & "SUM(" & Session(i & "Field") & ") as s" & Session(i & "Field")
      If Not Session(i & "IsCalcField") > "" Then
        sstrFields = sstrFields &  ", " & Session(i & "Field")
        sstrFields2 = sstrFields2 & Session(i & "Field")

      End if

    ElseIf Session(i & "IsCount") > "" Then
      sstrFields = sstrFields & "COUNT(*) as s" & Session(i & "Field")
      If Not Session(i & "IsCalcField") > "" Then
        sstrFields = sstrFields &  ", " & Session(i & "Field")

      End if

    Else
      If Not Session(i & "IsCalcField") > "" Then
          sstrFields = sstrFields & Session(i & "Field")
          sstrFields2 = sstrFields2 & Session(i & "Field")
        
      End if

    End If

  Next
REM  ShowMessageError sstrFields
REM  ShowMessageError sstrFields2

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
REM      If Not Session(i & "IsSum") > "" And Not Session(i & "IsCount") > "" And Not Session(i & "IsCalcField") > "" Then 
      Session("CurrentField") = Session(i & "ShowField")
        If Session(EditCurrentField & "EditOption") = conLookupField Then
           strJoin = strJoin & " LEFT OUTER JOIN " & Session(EditCurrentField & "LookupTable") &_
           " on " & Session(EditCurrentField & "LookupTable") & "." & Session(EditCurrentField & "LookupField") & " = " & _
           EditTableOf(Session("CurrentField")) & "." & Session(EditCurrentField & "Field")

        End If
REM      End If
    Next

    If sparEditFilter > "" Then
      strWhere = " WHERE " & sparEditFilter

    End If

    strGroup = " GROUP BY " & sstrFields2

    If Not Session("OrderField") > "" Then
      EditBeginReadOperation sstrFields

      Do While EditIsThereOperation
        If EditFindField(EditNextOperation) Then
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
      strOrder = " ORDER BY " & Session("OrderField")
      If Session("OrderDesc") = "1" Then
        strOrder = strOrder & " DESC "

      End If

    End If

    Dim intWidth

    intWidth = 0

    For i = 1 to Session("FieldNumber")
      Session("CurrentField") = Session(i & "ShowField")
      intWidth = intWidth + EditSize + 5

    Next

    If intWidth < 70 Then
      intWidth = 70

    ElseIf intWidth > 95 Then
      intWidth = 95

    End If

    BrowseSetActual EditCurrentTable

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

REM      ShowMessageError sql

      sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

      For i = 1 to Session("FieldNumber")
        If Session(i & "IsSum") > ""Or Session(i & "IsCount") > "" Then
          Session(i & "TotalValue") = sobjRS("s" & Session(i & "Field"))

        End If

      Next
      sobjRS.Close

    End If

    sql = strSelect & sstrFields & strFrom & strJoin & strWhere & strGroup & strOrder

    sobjCMD.CommandText = sql
    sobjCMD.CommandType = adCmdText

    ShowMessageError sql

    sobjRS.Open sobjCMD, , adOpenKeySet, adLockReadOnly

Dim strMax
    If Not sobjRS.EOF And Not sobjRS.BOF Then
    Dim strMessage
      If sblnHaveGraph Then
        Do While Not sobjRS.EOF
          For i = 1 to Session("FieldNumber")
            If Session(i & "IsGraph") > "" Then
              If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
                If Session(i & "MaxValue") < sobjRS("s" & Session(i & "Field")) Then
                  Session(i & "MaxValue") = sobjRS("s" & Session(i & "Field"))

                End If
              Else
                If Session(i & "MaxValue") < sobjRS(Session(i & "Field")) Then
                  Session(i & "MaxValue") = sobjRS(Session(i & "Field"))

                End If
              End If
            End If
          Next
          sobjRS.MoveNext

        Loop
        sobjRS.MoveFirst

      End If

      strMessage = Session(EditCurrentList & "ListCaption")

      Nbsp
      Center

      BrowseBegin
      BrowseHeader

      TableBegin strMessage, intNumCols, intWidth & "%"

      strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
      conOptionAZEdit & EditWhatList & EditRemakeTarget & """>Listagem Completa</A><BR>"

      For i = 1 to Session(EditCurrentTable & "ListNumber")
        If Session(EditCurrentTable & i & "ListName") <> sparEditList Then
          strTip2 = strTip2 & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
          conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditList & "=" & Session(EditCurrentTable & i & "ListName") & """>" & Session(EditCurrentTable & i & "ListCaption") & "</A><BR>"

        End If

      Next

      TableSetColumnColSpan intNumCols

      TableBeginRow 1
      TableBeginColumn
      TableNormalBegin "100%"
      TableSetColumnAlign "Right"

      TableBeginRow 1
      TableBeginColumn
      ShowHTMLCR "<FONT SIZE=2><DIV ID=F2 class=SS onMouseOver=""makeCool(F2);"" onMouseOut=""makeNormal(F2);"" onClick=""OpenDesc(T2);"" STYLE=""cursor:hand""><B>Outras listagens</B></DIV>"
      ShowHTMLCR "<table><tr><td align=justify>"
      ShowHTMLCR "<DIV ID=T2 class=dbody style=""display:'none';"">"
      Response.Write strTip2
      ShowHTMLCR "</DIV></td></tr></table>"
      TableEndColumn
      TableEndRow

      TableNormalEnd
      TableEndColumn
      TableEndRow
      TableSetColumnAlign ""
      TableSetColumnColSpan ""

      TableBeginHeadRow 1

      For i = 1 to Session("FieldNumber")
REM        ShowMessageError Session(i & "Field") & Session(i & "MaxValue") & "*" & Session(i & "TotalValue") & "$"
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

      TableEndHeadColumn

      TableEndHeadRow

      TableMultiGraph

      sstrGroupFieldLastValue = "kapushinefrebow"
      blnFirst = True
      Do While Not sobjRS.EOF And IsBrowsePage
        If sstrGroupField > "" Then
          If sstrGroupFieldLastValue <> "kapushinefrebow" Then
            If sstrGroupFieldLastValue <> sobjRS(Session(EditCurrentTable & sstrGroupField & "Field")) Then
              EditShowSummarySubTotal intNumCols

            End If
          End If

          sstrGroupFieldLastValue = sobjRS(Session(EditCurrentTable & sstrGroupField & "Field"))

        End If

        TableBeginRow 1
        For i = 1 to Session("FieldNumber")
          If Not Session(i & "Operation") > "" Then
            If Session(i & "IsCount") > "" Or Session(i & "IsSum") > "" Then
              Session(i & "Value") = sobjRS("s" & Session(i & "Field"))

            Else
              Session(i & "Value") = sobjRS(Session(i & "Field"))

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
            TableColumnGraph Session(i & "Value"), "<a href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & "&" & conPEditFilter & "=" & EditMakeFilter(sstrFields2) & ">" & FormatVariable(Session(i & "Value"), Session(i & "Type")) & "</a>"

          Else
            If Session(i & "IsCalcField") > "" Or Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Then
              TableSetColumnAlign "Right"

              TableBeginColumn
              EditBeginMakeLink False
              ShowHTMLCR FormatVariable(Session(i & "Value"), Session(i & "Type"))
              EditEndMakeLink
              TableEndColumn
              TableSetColumnAlign ""

            Else
              Session("CurrentField") = Session(i & "ShowField")
              If TypeName(EditFieldValue) = "Long" Or TypeName(EditFieldValue) = "Integer" Or _
                TypeName(EditFieldValue) = "Currency" Or TypeName(EditFieldValue) = "Double" Or TypeName(EditFieldValue) = "Single" Then
                TableSetColumnAlign "Right"

              End If

              TableBeginColumn
              EditBeginMakeLink False
              ShowHTMLCR BreakString(PutMask(FormatVariable(EditFieldValue, Session(EditCurrentField & "Type")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter")), Session(EditCurrentField & "ListChars"))
              TableSetColumnAlign ""
              EditEndMakeLink
              TableEndColumn

            End If
          End If

        Next
        TableEndRow

        sobjRS.MoveNext
        BrowseLineCounter

      Loop

      If sstrGroupField > "" Then
      Dim strColor
        EditShowSummarySubTotal intNumCols

        strColor = sTableSetRowBGColor
        TableSetRowBGColor sTableSetHeadRowBGColor
        TableBeginRow 1
        For i = 1 to Session("FieldNumber")
          If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
            If Session(i & "IsGraph") > "" Then
              TableBeginColumn
              TableNormalBegin ""
              TableBeginRow 1
              TableSetColumnWidth "70"
              TableSetColumnAlign "Right"
              TableCompleteColumn FormatVariable(Session(i & "Total"), Session(i & "Type")) & "&nbsp;"
              TableEndRow
              TableNormalEnd
              TableSetColumnWidth ""
              TableSetColumnAlign ""

              TableEndColumn

            Else
              TableSetColumnAlign "Right"
              TableCompleteColumn FormatVariable(Session(i & "Total"), Session(i & "Type"))
              TableSetColumnAlign ""

            End If

          Else
            TableCompleteColumn "&nbsp;"

          End If
        Next
        TableEndRow
        TableSetRowBGColor strColor

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
      TableEnd

      FormEnd

      BrowseSetStatus True
      BrowseEnd
      BrowseBar

      CenterEnd

    Else
      ShowMessageError "<center>Não há detalhamento na Tabela " & Session(EditCurrentList &_
      "Caption") & " para as opções escolhidas !!!</center><P>"
      ShowHTMLCR "&nbsp;<a href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd &_
                 "&" & EditWhatList &_
                 ">Incluir um registro</a><P>"

    End If

    BR
    sobjRS.Close

    EditFPLPLink

  Else
    ShowMessageError "Opção não existente - Listagem (" & sparEditList & ") !"

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
  strColor = sTableSetRowBGColor
  TableSetRowBGColor "WHITE"
  TableBeginRow 1
  For i = 1 to Session("FieldNumber")
    If Session(i & "IsSum") > "" Or Session(i & "IsCount") > "" Or Session(i & "IsTotal") > "" Then
      If Session(i & "IsGraph") > "" Then
        TableBeginColumn
        TableNormalBegin ""
        TableBeginRow 1
        TableSetColumnWidth "70"
        TableSetColumnAlign "Right"
        TableCompleteColumn FormatVariable(Session(i & "SubTotal"), Session(i & "Type")) & "&nbsp;"
        If Session(i & "TotalValue") > 0 Then
          TableSetColumnWidth "50"
          TableCompleteColumn "&nbsp;" & ZTINumber(100 * Session(i & "SubTotal") / Session(i & "TotalValue")) & "%" & "&nbsp;"

        End If
        TableSetColumnWidth ""
        TableEndRow
        TableNormalEnd
        TableSetColumnAlign ""

        TableEndColumn
        Session(i & "Total") = Session(i & "Total") + Session(i & "SubTotal")
        Session(i & "SubTotal") = 0

      Else
        TableSetColumnAlign "Right"
        TableCompleteColumn FormatVariable(Session(i & "SubTotal"), Session(i & "Type"))
        Session(i & "Total") = Session(i & "Total") + Session(i & "SubTotal")
        Session(i & "SubTotal") = 0
        TableSetColumnAlign ""

      End If
REM     ElseIf sstrGroupField = Session(i & "ShowField") Then
REM       TableCompleteColumn Session(i & "SubTotal")
    Else
      TableCompleteColumn "&nbsp;"

    End If

  Next
  TableEndRow

  TableSetRowBGColor strColor
  TableSetColumnColSpan intNumCols
  TableBeginRow 1
  TableCompleteColumn "&nbsp;"
  TableEndRow
  TableSetColumnColSpan ""

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

REM  ShowMessageError strType
  sintGraphNum = sintGraphNum + 1
  sintGraphActual = sintGraphNum

  If Session("1sintCor") = "" Then
     Session("1sintCor") = -1
     Session("1sintR")   = 255
     Session("1sintG")   = -85
     Session("1sintB")   = 0

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
  TableBeginRow 1
  TableBeginColumn
  If LCase(CStr(Session("OrderField"))) = LCase(CStr(strOrder)) Then
    If Session("OrderDesc") = "1" Then
      ShowImage "/ZTITools/img/EditForm/OrderDes.gif", "", "Ordenado descendentemente pelo campo " & EditCaption(False), sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter, "", 8, 7
      Nbsp
      Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter & """>" & strText & "</A>"

    Else
      ShowImage "/ZTITools/img/EditForm/OrderAsc.gif", "", "Ordenado ascendentemente pelo campo " & EditCaption(False), sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & conPEditOrderDescStr & GetFilter, "", 8, 7
      Nbsp
      Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & conPEditOrderDescStr & GetFilter & """>" & strText & "</A>"

    End If

  Else
    Response.Write "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionSummary & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Server.URLEncode(strOrder) & EditListStr & GetFilter & """>" & strText & "</A>"

  End If

  TableSetColumnAlign "Left"
  If intMax > 2 Then
    TableSetColumnWidth 60
    TableSetColumnAlign "Left"
    TableCompleteColumn "0"
    TableSetColumnAlign "Center"
    If intMax < 10 Then
      TableCompleteColumn FormatVariable(Round(intMax / 2, 1), strType)

    Else
      TableCompleteColumn FormatVariable(intMax \ 2, strType)

    End If
    TableSetColumnAlign "Right"
    TableCompleteColumn FormatVariable(intMax, strType)
    TableEndRow

    TableBeginRow 1

    TableSetColumnWidth 60
    If intTotal > 0 Then
       TableCompleteColumn "&nbsp;<Font Size=-1><B>" & FormatVariable(sarrTotal(sintGraphActual), strType) & "<B></Font>&nbsp;"

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
      TableCompleteColumn FormatVariable(i, strType)
      TableSetColumnAlign "Center"

    Next

    TableSetColumnAlign "Right"
    TableCompleteColumn FormatVariable(intMax, strType)
    TableEndRow

    TableBeginRow 1

    TableSetColumnWidth 60
    If intTotal > 0 Then
      TableCompleteColumn "&nbsp;<B>" & FormatVariable(sarrTotal(sintGraphActual), strType) & "<B>&nbsp;"

    Else
       TableCompleteColumn "&nbsp;"

    End If

    TableSetColumnAlign "Left"
    TableSetColumnWidth 120 \ (intMax + 1)
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
REM Desenha a Legenda do Grafico de Barras
REM -------------------------------------------------------------------------
Private Function EditMakeFilter(ByVal strFields)
Dim strAux

  EditBeginReadOperation strFields

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",",";","-","="

      Case Else
        EditFindField(EditLastOperation)
        If strAux > "" Then
          strAux = strAux + " AND "

        End If
        strAux = strAux & Session(EditCurrentField & "Field") & "=" & FormatForSQL(sobjRS(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"))

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
    EditWhatWhereHow = EditMakeTarget & EditWhatList & GetEditable

  Else
    EditWhatWhereHow = EditMakeTarget & EditWhatList

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditWhatWhereHow

REM -------------------------------------------------------------------------
REM Rotina que seleciona
REM -------------------------------------------------------------------------
Public Sub EditShowOptions(ByVal strShowOption)

  If Request.Form("Order") > "" Then
    Session("Descendente") = Request.Form("Descendente")
    Session("Order")       = Request.Form("Order")

  Else
    If Session(EditCurrentTable & "OrderNumber") > 0 Then
      Session("Order") = 1

    End If
  End If

  Select Case strShowOption
    Case conOptionFirstPage
      QueryStringSave
      Session("RecordAction") = ""
      ShowFirstPage

    Case conOptionEditor
      If Session("EditOption") <> Session("LastEditOption") Then
        QueryStringSave

      End If
      EditShowEditor

    Case conOptionFPEdit
      QueryStringSave
      Session("RecordAction") = ""
      EditReadFieldValues False
      sobjRS.Close
      EditShowEditor

    Case conOptionAZEdit
      QueryStringSave
      Session("RecordAction") = ""
      EditShowAZ

    Case conOptionSummary
      QueryStringSave
      Session("RecordAction") = ""
      EditShowSummary

    Case conOptionShowGrid
      QueryStringSave
      Session("RecordAction") = ""
      EditShowGrid "Codigo,Valor,Descricao,Data"

    Case conOptionFindCommon, conOptionFindAdvanced
      QueryStringSave
      FindFirstPage

  End Select

REM  ClearAllSessions

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditShowOptions

REM -------------------------------------------------------------------------
REM Rotina que mostra a tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditExeOptions(ByVal strShowOption)
  Select Case strShowOption
    Case conOptionExeEdition
      EditExeEdition

    Case conOptionAdd, conOptionCopy
      Session("RecordAction") = 1
      EditClear strShowOption
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
      EditWhatWhereHow

    Case conOptionEdit
      Session("RecordAction") = 2
      If EditReadFieldValues(True) Then
        Application.Lock
        Application(EditCurrentTable & EditMakeTarget) = Now()
        Application.Unlock
        Session("InEdition") = EditCurrentTable & EditMakeTarget
        sobjRS.Close
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEditor & _
        EditWhatWhereHow

      Else
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhatWhereHow

      End If

    Case conOptionDelete
      Session("RecordAction") = ""
      EditDelete
      If sintAbsolutePosition = 1 Then
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhereHow & "&" & conPEditMove & "=" & conMoveFirst

      ElseIf sintAbsolutePosition = sintRecordCount Then
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhereHow & "&" & conPEditMove & "=" & conMoveLast

      Else
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhatWhereHow & "&" & conPEditMove & "=" & conMovePrevious

      End If

    Case conOptionSave
      EditSave
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock
      Session("InEdition") = ""
      Session("RecordAction") = ""
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow

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
      EditWhatWhereHow & GetFilter & GetDefaults

    Case constrCopy
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionCopy & _
      EditWhatWhereHow & GetFilter

    Case constrEdit
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionEdit & _
      EditWhatWhereHow & GetFilter

    Case constrDelete
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionDelete & _
      EditWhatWhereHow & GetFilter

    Case constrSave
      EditGetForm
      MakeServerValidation conOptionSave

    Case constrSaveAdd
      EditGetForm
      MakeServerValidation conOptionSaveAdd

    Case constrSaveCopy
      EditGetForm
      MakeServerValidation conOptionSaveCopy

    Case constrList
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
      EditWhatList

    Case constrCancel
      Application.Lock
      Application(EditCurrentTable & EditMakeTarget) = ""
      Application.Unlock
      Session("InEdition") = ""
      Response.Redirect QueryStringLast

    Case constrFirst
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow & "&" & conPEditMove & "=" & conMoveFirst & GetFilter

    Case constrPrevious
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow & "&" & conPEditMove & "=" & conMovePrevious & GetFilter

    Case constrNext
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow & "&" & conPEditMove & "=" & conMoveNext & GetFilter

    Case constrLast
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
      EditWhatWhereHow & "&" & conPEditMove & "=" & conMoveLast & GetFilter

    Case constrFind
      If Session(EditCurrentTable & "FieldFind") > "" Then
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindAdvanced & _
        EditWhatWhereHow

      Else
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindCommon & _
        EditWhatWhereHow

      End If
  End Select

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditExeEdition

%>