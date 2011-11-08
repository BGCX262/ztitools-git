<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/FormPars.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 10:00h 16/12/1998
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Rotina que cria uma nova secao de tabela com seus parametros principais
REM -------------------------------------------------------------------------
Public Sub EditFormBegin(ByVal strTable, ByVal strCaption, ByVal intEditOption, ByVal strWhatList, _
                         ByVal blnClientValidation)
  Session("CurrentTable")                = strTable
  Session(strTable & "Caption")          = strCaption
  Session(strTable & "EditOption")       = intEditOption
  Session(strTable & "WhatList")         = strWhatList
  Session(strTable & "ClientValidation") = blnClientValidation
  Session(strTable & "ListNumber")       = 0
  Session(strTable & "OrderNumber")      = 0
  Session(strTable & "HeaderNumber")     = 0
  Session(strTable & "FilterNumber")     = 0
  Session(strTable & "GridNumber")       = 0

  ClearSession strTable & "UnitList"
  ClearSession strTable & "UnitTabs"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormBegin

REM -------------------------------------------------------------------------
REM Rotina que finaliza a criacao da secao de tabela
REM -------------------------------------------------------------------------
Public Sub EditFormEnd
Dim strJoinFormUnit, i, j, strTable, strFirstTable, strSecondTable

  strJoinFormUnit = ""

  For i = 1 to Session(EditCurrentTable & "FieldNumber")
    Session("CurrentField") = i
    If IsExternalField Then
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
        If Session(EditCurrentTable & sintCurrentRelation & "RelationType") = conRelation1to1 Then
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
    End If

  Next

  Session(EditCurrentTable & "JoinFormUnit") = strJoinFormUnit
REM  sobjShow.MessageError Session(EditCurrentTable & "JoinFormUnit")

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormEnd

REM -------------------------------------------------------------------------
REM Rotina que adiciona um campo da tabela na edicao
REM -------------------------------------------------------------------------
Public Sub EditAddFormField(ByVal strField, ByVal strCaption, ByVal intEditOption, ByVal intValidationOption, _
                            ByVal intShowOption, ByVal DefaultValue, ByVal strRequireMessage)
Dim i
  If EditFindField(EditCurrentTable, strField) Or (InStr(strField, ".") <> 0) Then
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
    Session(EditCurrentField & "DisableNumber")    = 0
    ClearSession EditCurrentField & "Hint"
    ClearSession EditCurrentField & "ExternalLink"
    ClearSession EditCurrentField & "InternalLink"
    ClearSession EditCurrentField & "ShowSize"
    ClearSession EditCurrentField & "BeginYear"
    ClearSession EditCurrentField & "EndYear"
    ClearSession EditCurrentField & "LookupTable"
    ClearSession EditCurrentField & "LookupField"
    ClearSession EditCurrentField & "LookupFieldText"
    ClearSession EditCurrentField & "ListSize"
    ClearSession EditCurrentField & "ListChars"
    Session(EditCurrentField & "CharCase")     = conNCase
    Session(EditCurrentField & "RadioColumns")     = 1
    If intEditOption = conTextAreaField Then
      Session(EditCurrentField & "TextAreaHeight") = 4

    End If
    If (InStr(strField, ".") <> 0) Then
      Dim strCurrentField, strCurrentTable

      strCurrentField  = EditCurrentField
      strCurrentTable  = EditCurrentTable

      Session("CurrentTable") = Left(strField, InStr(strField, ".") - 1)

      If EditFindField(EditCurrentTable, Mid(strField, InStr(strField, ".") + 1)) Then
        Session(strCurrentTable & i & "ExternalAddress") = EditCurrentField
        Session(strCurrentTable & i & "OwnerTable")     = Left(strField, InStr(strField, ".") - 1)
        Session(strCurrentTable & i & "Field")          = Mid(strField, InStr(strField, ".") + 1)
        Session(strCurrentTable & i & "Size")           = Session(EditCurrentField & "Size")
        Session(strCurrentTable & i & "Null")           = Session(EditCurrentField & "Null")
        Session(strCurrentTable & i & "Type")           = Session(EditCurrentField & "Type")

      Else
        Err.Raise 300, "EditAddFormField", "O campo """ & Mid(strField, InStr(strField, ".") + 1) & """, da tabela """ & _
              Left(strField, InStr(strField, ".") - 1) & """, não existe"

      End If
      Session("CurrentField") = i
      Session("CurrentTable") = strCurrentTable

    End If

  Else
    Err.Raise 300, "EditAddFormField", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditAddFormField

REM -------------------------------------------------------------------------
REM Rotina que altera a altura de um text area
REM -------------------------------------------------------------------------
Public Function EditFieldTextAreaHeight(strField, intHeight)
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "TextAreaHeight") = intHeight

  Else
    Err.Raise 300, "EditFormTextAreaHeight", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormTextAreaHeight

REM -------------------------------------------------------------------------
REM Rotina que altera a altura de um text area
REM -------------------------------------------------------------------------
Public Function EditFormFieldImage(strField, strImageRoot, strWidth)
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "ImageRoot")  = strImageRoot 
    Session(EditCurrentField & "ImageWidth") = strWidth

  Else
    Err.Raise 300, "EditFormFieldImage", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormTextAreaHeight

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Function EditFormAddFilter(strField, strValue)
Dim i
  Session(EditCurrentTable & "FilterNumber") = Session(EditCurrentTable & "FilterNumber") + 1
  i = Session(EditCurrentTable & "FilterNumber")
  
  Session(EditCurrentTable & i & "FilterField") = strField
  Session(EditCurrentTable & i & "FilterValue") = strValue

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormRadioColumns

REM -------------------------------------------------------------------------
REM Rotina que altera o numero de colunas de um campo radio
REM -------------------------------------------------------------------------
Public Function EditFieldRadioColumns(strField, intColumns)
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "RadioColumns") = intColumns

  Else
    Err.Raise 300, "EditFormRadioColumns", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormRadioColumns

REM -------------------------------------------------------------------------
REM Rotina que altera a opcao caso de caracter
REM -------------------------------------------------------------------------
Public Function EditFieldCharCase(strField, intCase)
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "CharCase") = intCase

  Else
    Err.Raise 300, "EditFieldCharCase", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldCharCase

REM -------------------------------------------------------------------------
REM Rotina que altera o tamanho do campo na apresentacao de listagens
REM -------------------------------------------------------------------------
Public Function EditFieldDisableValue(strField, strField2, Value, blnDisable)
Dim i, j, k, blnFind
  If EditFindField(EditCurrentTable, strField) Then
    If EditFindField(EditCurrentTable, strField2) Then
      i = 1
      blnFind = False
      Do While i <= Session(EditCurrentField & "DisableNumber") And Not blnFind
        If Session(EditCurrentField & "," & i & "DisableField") = strField Then
          blnFind = True
  
        Else
          i = i + 1
        
        End If

      Loop

      If blnFind Then
        j = Session(EditCurrentField & "," & i & "DisableValueNumber") + 1

      Else
        j = 1
        i = Session(EditCurrentField & "DisableNumber") + 1
        Session(EditCurrentField & "DisableNumber")                = i
        Session(EditCurrentField & "," & i & "DisableField")       = strField

      End If  
      Session(EditCurrentField & "," & i & "DisableValueNumber")     = j
      Session(EditCurrentField & "," & i & "DisableCondition")       = blnDisable
      If Session(EditCurrentField & "EditOption") = conRadioField Then
        For k = 1 to Session(EditCurrentField & "ValueNumber")
          If CStr(Session(EditCurrentField & "," & k & "Value")) = CStr(Value) Then
            Session(EditCurrentField & "," & i & "," & j & "DisableValue") = k - 1

          End If
        Next
      Else
        Session(EditCurrentField & "," & i & "," & j & "DisableValue") = Value

      End If

    Else
      Err.Raise 300, "EditFieldDisableValue", "O campo """ & strField2 & """, da tabela """ & _
                EditCurrentTable & """, não existe"

    End If

  Else
    Err.Raise 300, "EditFieldDisableValue", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldListSize

REM -------------------------------------------------------------------------
REM Rotina que altera o tamanho do campo na apresentacao de listagens
REM -------------------------------------------------------------------------
Public Function EditFieldListSize(strField, intSize)
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "ListChars") = intSize

  Else
    Err.Raise 300, "EditFieldListChars", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldListChars

REM -------------------------------------------------------------------------
REM Rotina que altera o numero de colunas de um campo radio
REM -------------------------------------------------------------------------
Public Function EditFormFieldMask(strField, strMask, strCompleter)
  If EditFindField(EditCurrentTable, strField) Then
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
Public Sub EditFormAddGrid(ByVal strTable, ByVal strFields, ByVal strName, ByVal intOption)
Dim i
  If Session(strTable & "FieldNumber") > "" Then
    Session(EditCurrentTable & "GridNumber") = Session(EditCurrentTable & "GridNumber") + 1
    i = Session(EditCurrentTable & "GridNumber")

    Session(EditCurrentTable & i & "GridTable")  = strTable
    Session(EditCurrentTable & i & "GridFields") = strFields
    Session(EditCurrentTable & i & "GridName")   = strName
    Session(EditCurrentTable & i & "GridOption") = intOption

  Else
    Err.Raise 300, "EditFormAddGrid", "A tabela """ & strTable & """ deve ser definida antes de ser colocada num grid"

  End If

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

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormUnit

REM -------------------------------------------------------------------------
REM Rotina define os textos dos tabs da tela de edicao
REM -------------------------------------------------------------------------
Public Sub EditFormUnitTabs(ByVal strUnitTabs)
  Session(EditCurrentTable & "UnitTabs") = strUnitTabs

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditFormUnitTabs

REM -------------------------------------------------------------------------
REM Rotina define uma funcao chamada a partir de um tab 
REM -------------------------------------------------------------------------
Public Sub EditTabTask(ByVal strTab, ByVal strTask)
  Session(EditCurrentTable & strTab & "Task")  = strTask

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTabTask

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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
  If EditFindField(EditCurrentTable, strField) Then
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
REM 
REM -------------------------------------------------------------------------
Public Function EditFormAddList(strList, strCaption, strName)
Dim i
  Session(EditCurrentTable & "ListNumber") = Session(EditCurrentTable & "ListNumber") + 1
  i = Session(EditCurrentTable & "ListNumber")

  Session(EditCurrentTable & i & "ListName")    = strName
  Session(EditCurrentTable & i & "List")        = strList
  Session(EditCurrentTable & i & "ListCaption") = strCaption

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFormAddList

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Function EditFieldLookupFilter(strField, strDependentField, strFilterField, strValue, strType)
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "FilterField")     = strFilterField
    Session(EditCurrentField & "DependentField")  = strDependentField
    Session(EditCurrentField & "FilterValue")     = strValue
    Session(EditCurrentField & "FilterFieldType") = strType

  Else
    Err.Raise 300, "EditFieldLookupFilter", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldLookupFilter

REM -------------------------------------------------------------------------
REM 
REM -------------------------------------------------------------------------
Public Function EditFieldDoubleValidate(strField, strField2, strMessage)
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "DoubleValidateField")   = strField2
    Session(EditCurrentField & "DoubleValidateMessage") = strMessage

  Else
    Err.Raise 300, "EditFieldDoubleValidate", "O campo """ & strField & """, da tabela """ & _
              EditCurrentTable & """, não existe"

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFieldDoubleValidate

%>