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
  Const conImageField         = 8

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

  Const conOptionFirstPage    = "0"
  Const conOptionExeFirstPage = "1"
  Const conOptionAZEdit       = "2"
  Const conOptionFPEdit       = "3"
  Const conOptionEditor       = "4"
  Const conOptionExeEdition   = "5"
  Const conOptionAdd          = "6"
  Const conOptionCopy         = "7"
  Const conOptionEdit         = "8"
  Const conOptionDelete       = "9"
  Const conOptionSave         = "10"
  Const conOptionSaveAdd      = "11"
  Const conOptionSaveCopy     = "12"

  Const conOptionShowGrid     = "13"

  Const conOptionFindCommon   = "14"
  Const conOptionFindAdvanced = "15"
  Const conOptionExeFind      = "16"

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

Dim sstrText, sstrText2, sstrText3, sstrLastValidate, sintSize, sintNumCol, sstrCheckAsRadio

  Public sstrCreateURL, sstrCreateURLMode, sstrCreateURLTarget

  Session("TableNumber") = 0
  sintSize = 0

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

  If Not sparEditOption > "" Then
    sparEditOption   = conOptionFirstPage

  End If
  If Not sparEditEditable > "" Then
    sparEditEditable = 0

  End If

  If Request.QueryString(conPEditOrderField) > "" Then
    Session("OrderField") = lCase(Request.QueryString(conPEditOrderField))

  End If
  Session("OrderDesc")  = lCase(Request.QueryString(conPEditOrderDesc))

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditQueryString

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

  Session("CurrentTable")                   = strTable
  Session("CurrentTableNum")                = Session("TableNumber")
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
Public Sub DataAddField(ByVal strField, ByVal strType, ByVal intSize, ByVal blnIsNull)
Dim i

  Session(EditCurrentTable & "FieldNumber") = Session(EditCurrentTable & "FieldNumber") + 1

  i = Session(EditCurrentTable & "FieldNumber")

  Session(EditCurrentTable & i & "Field")  = strField
  Session(EditCurrentTable & i & "Size")   = intSize
  Session(EditCurrentTable & i & "Type")   = strType
  Session(EditCurrentTable & i & "IsNull") = blnIsNull

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddField

REM -------------------------------------------------------------------------
REM Rotina que cria uma nova chave primaria
REM -------------------------------------------------------------------------
Public Sub DataAddPrimaryKey(ByVal strField)
  If EditFindField(strField) Then
    Session(EditCurrentField & "IsKey") = True

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

REM  ShowHTML Session(EditCurrentField & "Field") & " " & Session(EditCurrentField & "Caption")

  DataFindField = (i = Session("TableNumber"))

End Function
REM -------------------------------------------------------------------------
REM Final da Function DataFindField

REM -------------------------------------------------------------------------
REM Rotina que move para a proxima tabela
REM -------------------------------------------------------------------------
Public Sub DataNextTable(ByVal strField)

  If Session("CurrentTableNum") = Session("TableNumber") Then
    Session("CurrentTableNum") = 1
    Session("CurrentTable") = Session("Table1")

  Else
    Session("CurrentTableNum") = Session("CurrentTableNum") + 1
    Session("CurrentTable") = Session("Table" & Session("CurrentTableNum"))

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
  Session("CurrentTable")                = strTable
  Session(strTable & "Caption")          = strCaption
  Session(strTable & "EditOption")       = intEditOption
  Session(strTable & "OrderNumber")      = 0
  Session(strTable & "WhatList")         = strWhatList
  Session(strTable & "ClientValidation") = blnClientValidation

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
      Session("CurrentField")                          = Session(EditCurrentTable & "FieldNumber")
      i                                                = Session("CurrentField")

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
    Session(EditCurrentField & "RadioColumns")     = 1

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
REM Rotina que altera o numero de colunas de um campo radio
REM -------------------------------------------------------------------------
Public Function EditFormRadioColumns(strField, intColumns)
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
      REM ShowHTML EditCurrentTable & "*" & Session(EditCurrentField & "OwnerTable") & EditFindRelation(EditCurrentTable, Session(EditCurrentField & "OwnerTable"))
REM      Session(EditCurrentField & "Field") & "->" & Session(EditCurrentField & "OwnerTable") & " "
      If InStr(strJoinFormUnit, Session(EditCurrentField & "OwnerTable")) = 0 Then
        strTable = ""
        If EditFindRelation(EditCurrentTable, Session(EditCurrentField & "OwnerTable")) Then
           strTable       = EditCurrentTable & Session("CurrentRelation")
           strFirstTable  = EditCurrentTable
           strSecondTable = Session(EditCurrentField & "OwnerTable")

        Else
          If EditFindRelation(Session(EditCurrentField & "OwnerTable"), EditCurrentTable) Then
            strTable = Session(EditCurrentField & "OwnerTable") & Session("CurrentRelation")
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

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormEnd

REM -------------------------------------------------------------------------
REM Rotina que cria o formato da tela de listagem da tabela
REM -------------------------------------------------------------------------
Public Sub EditFormFind(ByVal strFieldFind)
  Session(EditCurrentTable & "FieldFind") = strFieldFind

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormList

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
  Do While (Session(strTable & i & "RelationTable") <> strTable2) And _
           (i < Session(strTable & "RelationNumber"))
    i = i + 1

  Loop

  If Session(strTable & i & "RelationTable") = strTable2 Then
    EditFindRelation = True
    Session("CurrentRelation") = i

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
REM Rotina retorna o nome do campo encontrado pela function EditFindField
REM -------------------------------------------------------------------------
Public Function EditCurrentField
  EditCurrentField = EditCurrentTable & Session("CurrentField")

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
          Session("CurrentTable") = Session("strSubText")

        End If
        sstrText = Mid(sstrText, 2)
        strChar = Left(sstrText, 1)

      Case "," , ";", "-", "=", ""
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
   Session("strSubText") = strSubText

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditNextOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna um dado ou operacao do texto de formato mas nao o exclui
REM -------------------------------------------------------------------------
Public Function EditSeeNextOperation

  EditSeeNextOperation = EditNextOperation
  sstrText             = Session("strSubText") & sstrText

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
    Session("intLastNumber") = CInt(strSubText)

  Else
    EditNextNumber = 1
    Session("intLastNumber") = 1

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

    IF Session("strSubText") = "=" Then
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

REM   ShowMessageError sql

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
          REM ShowHTML Session(EditCurrentTable & i & "Field")
        Next

      End If
      Session("AbsolutePosition") = sobjRS.AbsolutePosition
      Session("RecordCount")		  = sobjRS.RecordCount

    Else
      If Not blnEdit Then
        Session("RecordAction") = 1
        sparEditOption = conOptionAdd
        EditClear sparEditOption

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
    intColSpan = Session(EditCurrentTable & "UnitOperationNumber") \ Session("intLastNumber")

  End If

  If sintNumCol <= (Session(EditCurrentTable & "UnitOperationNumber") mod Session("intLastNumber")) Then
    intColSpan = intColSpan + 1

  End If

REM    ShowMessageError "Total:" & Session(EditCurrentTable & "UnitOperationNumber")
REM    ShowMessageError "Parcial:" & Session("intLastNumber")
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
  ShowHTMLCR "  case """ & constrDelete & """ :if (confirm(""Você tem certeza que deseja excluir este registro ?"")){"
  ShowHTMLCR "  document.frmEditForm.hidEditor.value = valor; return true;} else {return false;}  break;"
  If Session("RecordAction") > "" Then
    ShowHTMLCR "  case """ & constrCancel & """ :if (confirm(""Você tem certeza que deseja cancelar a inclusão do registro ? \nOs dados digitados serão perdidos."")){"
    ShowHTMLCR "  document.frmEditForm.hidEditor.value = valor; return true;} else {return false;} break;"
  
  End If

  If Session(EditCurrentTable & "ClientValidation") And Session("RecordAction") > "" Then
    ShowHTMLCR "  case """ & constrSave & """ :if (ValidateForm()){"
    ShowHTMLCR "  document.frmEditForm.hidEditor.value = valor; return true;} else {return false;} break;"
    ShowHTMLCR "  case """ & constrSaveAdd & """ :if (ValidateForm()){"
    ShowHTMLCR "  document.frmEditForm.hidEditor.value = valor; return true;} else {return false;} break;"
    ShowHTMLCR "  case """ & constrSaveCopy & """ :if (ValidateForm()){"
    ShowHTMLCR "  document.frmEditForm.hidEditor.value = valor; return true;} else {return false;} break;"

  End If

  ShowHTMLCR "  default : document.frmEditForm.hidEditor.value = valor; return true;"
REM  ShowHTMLCR "  document.frmEditForm.submit();"
  ShowHTMLCR "  }}"

  ShowHTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DeleteSaveMessages

REM -------------------------------------------------------------------------
REM Inclui uma imagem na tela
REM -------------------------------------------------------------------------
Private Sub ShowImage(strImagem, strID, strAlt, strHRef, strAlign)

  If strHRef > "" Then
    ShowHTML "<A HREF=" & strHRef & ">"

  End If

  ShowHTML "<IMG SRC=""" & strImagem & """ BORDER=0"

  If strAlt > "" Then
    ShowHTML " ALT=""" & strAlt & """"

  End If

  If strAlign > "" Then
    ShowHTML " ALIGN=""" & strAlign & """"

  End If

  If strID > "" Then
    ShowHTML " ID=""" & strID & """"

  End If

	If sFormInputonClick > "" Then
		ShowHTML  " onClick=""" & sFormInputonClick & ""

	End If

  ShowHTML """>"

  If strHRef > "" Then
    ShowHTML "</A>"

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

REM -------------------------------------------------------------------------
REM Rotina que mostra a tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditShowEditor
Dim blnNotStarting, FontSize, strURL, i, blnFirstEditable, strLastOperation, blnDisable

  Session("LastToFind") = conOptionFPEdit
  blnFirstEditable = True
  BR
  Center

  JavaScriptDefaultMessages

  If Session(EditCurrentTable & "ClientValidation") Then
    If Session("RecordAction") > "" Then
      ClientValidates
    
    End If  

  Else
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

  If Left(sstrText, 1) <> "=" And Not sparEditFilter > "" Then
    StartRow
    TableBeginRow 1
    TableBeginColumn
    TableNormalBegin "100%"
    blnNotStarting = True
    EditColumnColSpan
    TableBeginRow FontSize
    TableBeginColumn

  End If

  If sparEditFilter > "" Then
    TableBeginHeadRow 2
    TableSetColumnAlign "LEFT"
    TableHeadColumn "sdfsdfsd"
    TableSetColumnAlign ""
    TableEndHeadRow
    TableBeginRow 1
    TableBeginColumn
    TableNormalBegin "100%"
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

      Case Else
        If (Session(EditCurrentTable & "EditOption") > 0) And (sparEditEditable = 1) Then
          blnFirstEditable = EditShowField(blnFirstEditable And Session("RecordAction") > "", False, blnDisable, False)

        Else
          EditShowFieldAsText

        End If

    End Select

  Loop

  TableNormalEnd

  TableEndColumn
  TableEndRow

REM  TableSetRowBGColor "Black"

  TableBeginRow 2
  TableBeginColumn
  Center

  FormSetDisabled ""

  FormInputHidden "hidEditor", ""

  If Not Session("RecordAction") > "" Then

    If Session("AbsolutePosition") > 1 Then
      FormInputSetOnClick "return ValueToHidden('" & constrFirst & "')"

    Else
      FormInputSetOnClick "return false"

    End If
    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdPrimeiroWin.gif", "Vai para o primeiro registro"

    If Session("AbsolutePosition") > 1 Then
      FormInputSetOnClick "return ValueToHidden('" & constrPrevious & "')"

    Else
      FormInputSetOnClick "return false"

    End If

    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdAnteriorWin.gif", "Vai para o registro anterior"

  End If

  If sparEditEditable = 1 Then
    Select Case Session(EditCurrentTable & "EditOption")
      Case "1"
        If Not Session("RecordAction") > "" Then
          FormSetAccessKey "I"
          FormInputSetOnClick "return ValueToHidden('" & constrAdd & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdIncluirWin.gif", "Inclui um registro"
          FormSetAccessKey ""
          FormInputSetOnClick "return ValueToHidden('" & constrCopy & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdCopiarWin.gif", "Inclui um registro copiando os dados deste"
          FormSetAccessKey "A"
          FormInputSetOnClick "return ValueToHidden('" & constrEdit & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdAlterarWin.gif", "Entra em modo de alteração"
          FormSetAccessKey "P"
          FormInputSetOnClick "return ValueToHidden('" & constrDelete & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdApagarWin.gif", "Exclui este registro"
          FormSetAccessKey ""

        End If

        If Not Session("RecordAction") > "" Then
REM        FormInputSubmit "btmEditor", constrFind

        Else
          FormSetAccessKey "S"
          FormInputSetOnClick "return ValueToHidden('" & constrSave & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarWin.gif", "Salva as alterações feitas neste registro"
          FormSetAccessKey ""
          FormInputSetOnClick "return ValueToHidden('" & constrSaveAdd & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarIncluirWin.gif", "Salva os dados deste registro e inclui outro"
          FormInputSetOnClick "return ValueToHidden('" & constrSaveCopy & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdSalvarCopiarWin.gif", "Salva os dados deste registro e copiando os dados deste"
          FormSetAccessKey "C"
          FormInputSetOnClick "return ValueToHidden('" & constrCancel & "')"
          FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdCancelarWin.gif", "Cancela a inclusão deste registro"
          FormSetAccessKey ""

        End If

    End Select
  End If

  If Not Session("RecordAction") > "" Then
    FormInputSetOnClick "return ValueToHidden('" & constrFind & "')"
    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdLocalizarWin.gif", "Entra em modo de localização de registros"
    FormInputSetOnClick "return ValueToHidden('" & constrList & "')"
    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdListagemWin.gif", "Vai para a forma de visualização de listagem"

    If Session("AbsolutePosition") = Session("RecordCount") Then
      FormInputSetOnClick "return false"

    Else
      FormInputSetOnClick "return ValueToHidden('" & constrNext & "')"

    End If
    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdProximoWin.gif", "Vai para o próximo registro"

    If Session("AbsolutePosition") = Session("RecordCount") Then
      FormInputSetOnClick "return false"

    Else
      FormInputSetOnClick "return ValueToHidden('" & constrLast & "')"

    End If
    FormInputImage "btmEditor", "", "/ZTITools/Img/EditForm/cmdUltimoWin.gif", "Vai para o úlitmo registro"
    FormInputSetOnClick ""

  End If

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

  If sparEditOption <> conOptionFirstPage Then
    If Session("EditFPLPLink") > "" Then
      ShowHTML "<A HREF=" & Session("EditFPLPLink")

    Else
      ShowHTML "<A HREF=" & sstrThisScriptName

    End If

    ShowHTMLCR ">Página Inicial</a>"

  End If

  If QueryStringLast > "" Then
    ShowHTMLCR "<A HREF=" & QueryStringLast & ">Última Opção</a>"

  End If

  CenterEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFPLPLink

REM -------------------------------------------------------------------------
REM Limpa as variaveis da tabela corrente
REM -------------------------------------------------------------------------
Private Sub EditClear(intClearOption)
Dim sql, i
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

        Else
          Session(EditCurrentTable & i & "Value") = 1

        End If
        sobjRS2.Close
REM        ShowMessageError Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "Value")

      Case Else
        If intClearOption = conOptionAdd Then
          Session(EditCurrentTable & i & "Value") = Session(EditCurrentTable & _
                                                           i & "DefaultValue")

        End If

    End Select

  Next

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditClear

REM -------------------------------------------------------------------------
REM Rotina que seleciona e adiciona um form ou texto de um campo
REM -------------------------------------------------------------------------
Public Function EditSizeToShow
  If Session(EditCurrentField & "ShowSize") > "" Then
    EditSizeToShow = Session(EditCurrentField & "ShowSize")

  Else
    EditSizeToShow = Session(EditCurrentField & "Size")

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditSizeToShow

REM -------------------------------------------------------------------------
REM Rotina que coloca um indicador de campo requerido
REM -------------------------------------------------------------------------
Public Sub EditShowFieldImages
  If Session(EditCurrentField & "RequireMessage") > "" Then
    ShowImage "/ztitools/img/editform/CampoRequerido.gif", "", "Campo obrigatório", "", ""
    Nbsp

  End If
  If Session(EditCurrentField & "ValidationError") > "" Then
    ShowImage "/ztitools/img/editform/CampoErrado.gif", "", Session(EditCurrentField & "ValidationError"), "", ""
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

  If EditFindField(Session("strSubText")) Then
    If Session(EditCurrentField & "ShowOption") > -1 Then
      FormSetTitle Session(EditCurrentField & "Hint")
      If blnEmpty Then
        Value = ""

      Else
        Value = Session(EditCurrentField & "Value")

      End If

      Select Case Session(EditCurrentField & "EditOption")
        Case conTextField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If
          FormInputTextMaxLength Session(EditCurrentField & "Field"), EditSizeToShow, _
          FormatVariable(Value), Session(EditCurrentField & "Size")

        Case conTextAreaField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If
          FormTextArea Session(EditCurrentField & "Field"), 4, EditSizeToShow, Value

        Case conUFField
          If Not blnFieldOnly Then
            EditShowCaption True
            EditShowFieldImages
            BR

          End If
          FormSelectUF Session(EditCurrentField & "Field"), 1, False, Value

        Case conLookupField
          If (Session(EditCurrentField & "LookupTable") > "") Or (Session(EditCurrentField & "LookupField") > "")_
          Or (Session(EditCurrentField & "LookupFieldText") > "") Then

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

REM            ShowMessageError sql

            sobjRS2.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly
            FormSelectBegin Session(EditCurrentField & "Field"), 1, False
            If Not Value > "" Then
              FormSelectOption "", "", True

            End If
            Do While Not sobjRS2.EOF
              FormSelectOption sobjRS2(Session(EditCurrentField & "LookupField")), _
                                             sobjRS2(Session(EditCurrentField & "LookupFieldText")), _
                                             (ToString(sobjRS2(Session(EditCurrentField & "LookupField"))) _
                                             = ToString(Value))

              sobjRS2.MoveNext

            Loop
            FormSelectEnd

            sobjRS2.Close

          Else
            Err.Raise 300, "EditShowField", "O campo """ & strField & """, da tabela """ & EditCurrentTable & _
            """, está definido como lookup, mas sua definição de lookup não foi feita ou não está completa"

          End If

        Case conCheckField
          If Not blnFieldOnly Then
            EditShowCaption True

          End If
          If Not sstrCheckAsRadio > "" Then
            FormInputCheckBox Session(EditCurrentField & "Field"), Session(EditCurrentField & ",1Value"),_
            "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & ",1Value")))
            If Not blnFieldOnly Then
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
REM          TableBegin "", 1, "100%"
REM          TableBeginRow 1
REM          TableBeginColumn
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
            TableBeginColumn 
            FormInputRadio Session(EditCurrentField & "Field"), Session(EditCurrentField & "," & i & "Value"), "left", BitBln(ToString(Value) = ToString(Session(EditCurrentField & "," & i & "Value")))
            TableEndColumn
            TableBeginColumn 
            ShowHTMLCR Session(EditCurrentField & "," & i & "Caption")
            TableEndColumn
            TableEndRow
            
          Next           
          TableNormalEnd
          TableEndColumn
          TableEndRow
          TableNormalEnd
REM          TableEndColumn
REM          TableEndRow
REM          TableEnd
          
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

          ShowHTML "<B>/</B>"

          FormSelectBegin Session(EditCurrentField & "Field") & "MonthValue", 1, False

          For i = 1 to 12
            FormSelectOption LeadingZeroes(i, 2), LeadingZeroes(i, 2), i = Month(Value)

          Next

          FormSelectEnd

          ShowHTML "<B>/</B>"

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
    ShowHTML "<FONT SIZE=-1>" & Session("strSubText") & "</FONT>"

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

  If EditFindField(Session("strSubText")) Then
    ShowHTML "<FONT SIZE=-1>"
    Select Case Session(EditCurrentField & "EditOption")
      Case conTextField, conUFField, conTextAreaField
        BoldText EditCaption(False)
        ShowHTML "&nbsp;:&nbsp;"

        If (Session(EditCurrentField & "ExternalLink") > "") Then
          EditBeginMakeLink True

        End If                                              
                              
        ShowHTML PutMask(FormatVariable(Session(EditCurrentField & "Value")), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter"))
        EditEndMakeLink

      Case conLookupField
        BoldText EditCaption(False)
        ShowHTML "&nbsp;:&nbsp;"
        sql = "SELECT * FROM " & Session(EditCurrentField & "LookupTable")
        sql = sql & " WHERE " & Session(EditCurrentField & "LookupField") & "=" & _
                    Session(EditCurrentField & "Value")

        sobjCMD.CommandText = sql

        sobjRS2.Open sobjCMD, , adOpenForwardOnly, adLockReadOnly

        If Not sobjRS2.EOF And Not sobjRS2.BOF Then
          EditBeginMakeLink True
          ShowHTML FormatVariable(sobjRS2(Session(EditCurrentField & "LookupFieldText")))
          EditEndMakeLink

        End If
        sobjRS2.Close

      Case conCheckField
        BoldText EditCaption(False)
        Nbsp
        EditBeginMakeLink True
        If Session(EditCurrentField & "Value") = Session(EditCurrentField & ",1Value") Then
          ShowHTMLCR FormatVariable(Session(EditCurrentField & ",1Caption"))

        Else
          ShowHTMLCR FormatVariable(Session(EditCurrentField & ",2Caption"))

        End If
        EditEndMakeLink

      Case conRadioField, conSelectField
        BoldText EditCaption(False)
        Nbsp
        EditBeginMakeLink True
        For i = 1 to Session(EditCurrentField & "ValueNumber")
          If Session(EditCurrentField & "Value") = Session(EditCurrentField & "," & i & "Value") Then
            ShowHTMLCR FormatVariable(Session(EditCurrentField & "," & i & "Caption"))

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
    Case conTextField, conUFField
      Value = sobjRS(Session(EditCurrentField & "Field"))

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
Public Function FormatVariable(Value)
  Select Case LCase(TypeName(Value))
    Case "long", "integer", "byte"
      FormatVariable = ZTIInt(Value)

    Case "currency"
      FormatVariable = ZTIMoney(Value)

    Case "single", "double"
      FormatVariable = ZTINumber(Value)

    Case "date"
      FormatVariable = ZTILongDate(Value)

    Case Else
      FormatVariable = Value

  End Select

End Function
REM -------------------------------------------------------------------------
REM Final da Function FormatVariable

REM -------------------------------------------------------------------------
REM Formata uma variavel de acordo com seu tipo para ser usada numa chamada SQL
REM -------------------------------------------------------------------------
Public Function FormatForSQL(ByVal Value, ByVal strType)
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
REM Rotina que mostra a tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditShowAZ
Dim sql, i

  Session("LastToFind") = conOptionAZEdit

  TableSetColumnNoWrap False
  FormBegin sstrThisScriptURL & "?" & sstrThisQueryString, "POST", "frmEditForm", "", ""

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
         Session("CurrentField") = Session("strSubText")
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
         Case ",", ";", "-", ""
         Case "="
           EditIsThereOperation

         Case Else
           If EditFindField(Session("strSubText")) Then
             Width = Width + Session(EditCurrentField & "Size")
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

    If Session(EditCurrentTable & "OrderNumber") > 0 Then
      Dim strAux, strAux2, strTip

      For i = 1 to Session(EditCurrentTable & "OrderNumber")
        strAux = ""
        strAux2 = ""
        EditBeginReadOperation Session(EditCurrentTable & i & "Order")
        Do While EditIsThereOperation
          Select Case EditNextOperation
            Case ",", ";", "-", "", "="
              strAux = strAux & Session("strSubText") & " "

            Case Else
              If EditFindField(Session("strSubText")) Then
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
            EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & """>" & _
            "&nbsp;" & strAux & "</A><BR>"

          Else
            strTip = strTip & "<IMG SRC=""/ZTITools/img/EditForm/OrderAsc.gif"" BORDER=0>" & _
            "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & _
            EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & conPEditOrderDescStr & """>" & _
             "&nbsp;" & strAux & "</A><BR>"

          End If
        Else
          strAux2 = Server.URLEncode(strAux2)
          strTip = strTip & "&nbsp;&nbsp;&nbsp;<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & _
          conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & strAux2 & """>" & strAux & "</A><BR>"

        End If

      Next

    End If

    TableSetColumnColSpan Session(EditCurrentTable & "ListOperationNumber") + 1

    TableBeginRow 1
    TableBeginColumn

    TableNormalBegin "100%"
    TableSetColumnAlign "Left"
    TableSetColumnVAlign "Top"
    TableBeginRow 1
    TableBeginColumn
    BoldText "<a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd & _
               EditWhatList & conPEditEditableStr & ">Incluir</a>"
    Nbsp
    Nbsp
    BoldText "<a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
               EditWhatList & conPEditEditableStr & ">Editar</a>"
    Nbsp
    Nbsp
    BoldText "<a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindCommon & _
               EditWhatList & ">Localizar</a>"

    If Session(EditCurrentTable & "FieldFind") > "" Then
      Nbsp
      Nbsp
      BoldText "<a CLASS=SS href=" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFindAdvanced & _
               EditWhatList & ">Localizar avançado</a>"

    End If
    TableEndColumn

    If Session(EditCurrentTable & "OrderNumber") > 0 Then
      TableSetColumnAlign "Right"
      TableSetColumnVAlign "Top"
      TableBeginColumn
      ShowHTMLCR "<FONT SIZE=2><DIV ID=F1 class=SS onMouseOver=""makeCool(F1);"" onMouseOut=""makeNormal(F1);"" onClick=""OpenDesc(T1);"" STYLE=""cursor:hand""><B>Outras formas de ordenação</B></DIV>"
      ShowHTMLCR "<table><tr><td align=justify>"
      ShowHTMLCR "<DIV ID=T1 class=dbody style=""display:'none';"">"
      BoldText strTip
      ShowHTMLCR "</DIV></td></tr></table>"
      TableEndColumn

    End If

    TableEndRow
    TableNormalEnd
    TableEndColumn
    TableEndRow
    TableSetColumnAlign ""
    TableSetColumnColSpan ""
    TableBeginHeadRow 1

    EditBeginReadOperation strList

    Do While EditIsThereOperation
      Select Case EditNextOperation
         Case ","

         Case Else
           Session("CurrentField") = Session("strSubText")
           TableHeadColumnBegin
           If CStr(Session("OrderField")) = CStr(Session("CurrentField")) Then
             If Session("OrderDesc") = "1" Then
               ShowImage "/ZTITools/img/EditForm/OrderDes.gif", "", "Ordenado descendentemente pelo campo " & EditCaption(False), sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Session("CurrentField"), ""
               Nbsp
               ShowHTML "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Session("CurrentField") & """>" & EditCaption(False) & "</A>"

             Else
               ShowImage "/ZTITools/img/EditForm/OrderAsc.gif", "", "Ordenado ascendentemente pelo campo " & EditCaption(False), sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Session("CurrentField") & conPEditOrderDescStr, ""
               Nbsp
               ShowHTML "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Session("CurrentField") & conPEditOrderDescStr & """>" & EditCaption(False) & "</A>"

             End If
           Else
             ShowHTML "<A CLASS=SS HRef=""" & sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAZEdit & EditWhatList & EditRemakeTarget & "&" & conPEditOrderField & "=" & Session("CurrentField") & """>" & EditCaption(False) & "</A>"

           End If
           TableHeadColumnEnd

       End Select

    Loop
    TableEndHeadRow

    Do While Not sobjRS.EOF And IsBrowsePage
      EditBeginReadOperation strList

      TableBeginRow 1

      Do While EditIsThereOperation
        Select Case EditNextOperation
          Case ","

          Case Else
            Session("CurrentField") = Session("strSubText")
            If TypeName(EditFieldValue) = "Long" Or TypeName(EditFieldValue) = "Integer" Or _
              TypeName(EditFieldValue) = "Currency" Then
              TableSetColumnAlign "Right"

            End If

            TableBeginColumn
            EditBeginMakeLink False
            ShowHTMLCR PutMask(FormatVariable(EditFieldValue), Session(EditCurrentField & "Mask"), Session(EditCurrentField & "MaskCompleter"))
            TableSetColumnAlign ""
            EditEndMakeLink

            TableEndColumn

        End Select

      Loop
      TableEndRow

      sobjRS.MoveNext
      BrowseLineCounter
    Loop

    TableEnd

    FormEnd

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
REM Monta o link do campo corrente
REM -------------------------------------------------------------------------
Public Sub EditBeginMakeLink(ByVal blnPrefforExtenal)
  If blnPrefforExtenal Then
    If Session(EditCurrentField & "ExternalLink") > "" Then
      ShowHTML "<A HREF=" & Session(EditCurrentField & "ExternalLink") & ">"

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
        ShowHTML "<A HREF=" & Session(EditCurrentField & "ExternalLink") & ">"

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

  ShowHTML strURL

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditMakeInternalLink

REM -------------------------------------------------------------------------
REM Monta o link interno do campo corrente
REM -------------------------------------------------------------------------
Public Function EditTargetOfLinks
Dim strAux, i

  For i = 1 to Session(EditCurrentField & "ILTargetNumber")
    If strAux > "" Then
      strAux = strAux & " And "

    End If
    If sobjRS.State = 0 Then
      strAux = strAux & Session(EditCurrentField & "," & i & "ILTarget") & "=" & Session(EditCurrentField & "Value")
    
    Else
      strAux = strAux & Session(EditCurrentField & "," & i & "ILTarget") & "=" & sobjRS(Session(EditCurrentField & "Field"))
    
    End If
  
  Next

  EditTargetOfLinks = "&" & conPEditFind & "=" & Server.URLEncode(strAux)

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
      strAux = strAux & Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "OldValue")

    End If
  Next

  EditMakeTarget = "&" & conPEditFind & "=" & Server.URLEncode(strAux)

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditMakeTarget

REM -------------------------------------------------------------------------
REM Remonta o ultimo link dos valores campos
REM -------------------------------------------------------------------------
Public Function EditRemakeTarget
Dim strAux, i

  EditRemakeTarget = "&" & conPEditFind & "=" & sparEditFind

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
         If EditFindField(Session("strSubText")) Then
           Select Case Session(EditCurrentField & "EditOption")
             Case conSeparateDateField
                Session(EditCurrentField & "Value") = ImportDate(Request.Form(Session(EditCurrentField & "Field") & _
                                                "DayValue") & "/" & Request.Form(Session(EditCurrentField & "Field") & _
                                                "MonthValue") & "/" & Request.Form(Session(EditCurrentField & "Field") &_
                                                "YearValue"), conDateTypeDDMMAAAA)

              Case Else                                                   
REM                Session(EditCurrentField & "Value") = EditFormatField(Request.Form(Session(EditCurrentField & "Field")), Session(EditCurrentField & "Type"), Session(EditCurrentField & "Size"))
                Session(EditCurrentField & "Value") = Request.Form(Session(EditCurrentField & "Field"))

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
        sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatField(Session(EditCurrentTable & _
                                                                 i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                 Session(EditCurrentTable & i & "Size"))
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
         strTable = EditCurrentTable & Session("CurrentRelation")

      ElseIf EditFindRelation(strCurrentTable, EditCurrentTable) Then
         strTable = strCurrentTable & Session("CurrentRelation")

      End If

      If strTable > "" Then
        For j = 1 to Session(strTable & "FieldtoFieldNumber")
          If EditFindField(Session(strTable & "," & j & "RelationField1")) Then
            sobjRS(Session(strTable & "," & j & "RelationField2")) = EditFormatField(Session(EditCurrentField & "Value"), _
                                                                     Session(EditCurrentField & "Type"), _
                                                                     Session(EditCurrentField & "Size"))
          End If
        Next

      End If

      For i = 1 to Session(EditCurrentTable & "FieldNumber")
        If Session(EditCurrentTable & i & "OwnerTable") = strCurrentTable Then
REM           ShowMessageError Session(EditCurrentTable & i & "Field") & "=" & Session(EditCurrentTable & i & "Value")
           sobjRS(Session(EditCurrentTable & i & "Field")) = EditFormatField(Session(EditCurrentTable & _
                                                                    i & "Value"), Session(EditCurrentTable & i & "Type"), _
                                                                    Session(EditCurrentTable & i & "Size"))
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
         strTable = EditCurrentTable & Session("CurrentRelation")

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
                                                                     Session(EditCurrentField & "Type"), _
                                                                     Session(EditCurrentField & "Size"))
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
        sql = sql & " WHERE " & sparEditFind

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

  ErrorImage = "<IMG SRC=""/ztitools/img/editform/Ball.gif"">&nbsp;"

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
        If EditFindField(Session("strSubText")) Then
          If (Session(EditCurrentField & "RequireMessage") > "") Then
            If Not Session(EditCurrentField & "Value") > "" Then
              Session("Message") = Session("Message") & ErrorImage & "O campo <B>" &  EditCaption(False) & "</B> deve ser preenchido</B><BR>"
              Session(EditCurrentField & "ValidationError") = "O campo " &  EditCaption(False) & " deve ser preenchido"
              blnIsValid = False

            End If
          End If
        End If
  REM    ShowHTMLCR Session("strSubText")

    End Select

  Loop

  EditBeginReadOperation Session(EditCurrentTable & "FieldUnit")

  Do While EditIsThereOperation
    Select Case EditNextOperation
      Case ",", ";", "-", ""
      Case "="
        EditIsThereOperation

      Case Else
        If EditFindField(Session("strSubText")) Then
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
        If EditFindField(Session("strSubText")) Then
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
        If EditFindField(Session("strSubText")) Then
          If Session(EditCurrentField & "RequireMessage") > "" Then
            FormJavaRequired Session(EditCurrentField & "Field"), Session(EditCurrentField & "RequireMessage")

          End If
        End If

    End Select

  Loop
  If sstrLastValidate > "" Then
    ShowHTMLCR "return (" & sstrLastValidate & ");"

  Else
    ShowHTMLCR "return true;"

  End If

  ShowHTML "}"
  ShowHTML "//-->"
  ShowHTMLCR "</SCRIPT>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ClientValidates

REM -------------------------------------------------------------------------
REM Retorna a parte da QueryString referente a qual listagem esta sendo usada
REM -------------------------------------------------------------------------
Private Function EditWhatList

  EditWhatList = "&" & conPEditWhat & "=" & Session(EditCurrentTable & "WhatList")

End Function
REM -------------------------------------------------------------------------
REM Final da Sub EditWhatList

REM -------------------------------------------------------------------------
REM Retorna a parte da QueryString referente a listagem, e modo de insercao
REM -------------------------------------------------------------------------
Private Function EditWhereHow

  If sparEditEditable = "1" Then
    EditWhereHow = EditWhatList & conPEditEditableStr

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
    EditWhatWhereHow = EditMakeTarget & EditWhatList & conPEditEditableStr

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

    Case conOptionShowGrid
      QueryStringSave
      Session("RecordAction") = ""
      EditShowGrid "Codigo,Valor,Descricao,Data"

    Case conOptionFindCommon, conOptionFindAdvanced
      FindFirstPage

  End Select

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
      If Session("AbsolutePosition") = 1 Then
        Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionFPEdit & _
        EditWhereHow & "&" & conPEditMove & "=" & conMoveFirst

      ElseIf Session("AbsolutePosition") = Session("RecordCount") Then
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
  Select Case Request.Form("hidEditor")
    Case constrAdd
      Response.Redirect sstrThisScriptURL & "?" & conPEditOption & "=" & conOptionAdd & _
      EditWhatWhereHow & GetFilter

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
           If EditFindField(Session("strSubText")) Then
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
            If EditFindField(Session("strSubText")) Then
              If TypeName(EditFieldValue) = "Long" Or TypeName(EditFieldValue) = "Integer" Or _
                TypeName(EditFieldValue) = "Currency" Then
                TableSetColumnAlign "Right"

              End If

              TableBeginColumn
              EditBeginMakeLink False
              ShowHTMLCR FormatVariable(EditFieldValue)
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

%>