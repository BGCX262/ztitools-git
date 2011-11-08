<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/DataPars.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 10:40h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

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
  If EditFindField(EditCurrentTable, strField) Then
    Session(EditCurrentField & "IsKey") = True
REM      sobjShow.MessageError Session(EditCurrentField & "Field")

  Else
    Err.Raise 300, "EditDataAddPrimaryKey", "O campo """ & strField & """, da tabela """ & EditCurrentTable & """, não existe"

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditDataAddPrimaryKey

REM -------------------------------------------------------------------------
REM Rotina que adiciona um relacionamento entre tabelas
REM -------------------------------------------------------------------------
Public Sub DataAddRelation(ByVal strTable, ByVal intDeleteOption, ByVal intUpdateOption, ByVal intRelationType)
Dim i

  Session(EditCurrentTable & "RelationNumber")     = Session(EditCurrentTable & "RelationNumber") + 1

  i = Session(EditCurrentTable & "RelationNumber")

  Session(EditCurrentTable & i & "RelationTable")      = strTable
  Session(EditCurrentTable & i & "DeleteOption")       = intDeleteOption
  Session(EditCurrentTable & i & "UpdateOption")       = intUpdateOption
  Session(EditCurrentTable & i & "RelationType")       = intRelationType
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

  Do While (Not EditFindField(EditCurrentTable, strField)) And (i < Session("TableNumber"))
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
REM Final do Include DataPars.inc

%>