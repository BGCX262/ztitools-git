<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Current.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 11:00h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
REM Retorna o nome da tabela atual
REM -------------------------------------------------------------------------
Public Function EditCurrentTable
  EditCurrentTable = Session("CurrentTable")

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditCurrentTable

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
Public Function EditCurrentGrid
  EditCurrentGrid = Session("CurrentGrid")

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
REM Rotina que encontra um campo em uma tabela
REM -------------------------------------------------------------------------
Public Function EditFindField(ByVal strTable, ByVal strField)
Dim i

  i = 1
  Do While (LCase(Session(strTable & i & "Field")) <> LCase(strField)) And _
           (i < Session(strTable & "FieldNumber"))

    i = i + 1

  Loop

  If LCase(Session(strTable & i & "Field")) = LCase(strField) Then
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
Public Function EditFindGrid(ByVal strGrid)
Dim i

  i = 1
  Do While (LCase(Session(EditCurrentTable & i & "GridName")) <> LCase(strGrid)) And _
           (i < Session(EditCurrentTable & "GridNumber"))

    i = i + 1

  Loop

  If LCase(Session(EditCurrentTable & i & "GridName")) = LCase(strGrid) Then
    EditFindGrid = True
    Session("CurrentGrid") = i

  Else
    EditFindGrid = False

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditFindGrid

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
REM Final da Function EditFindList

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
REM Rotina que le os dados e tamanhos dos campos ja adicionados na tabela
REM -------------------------------------------------------------------------
Private Function EditTableOf(strTable, intField)
  If Session(strTable & intField & "OwnerTable") > "" Then
    EditTableOf = Session(strTable & intField & "OwnerTable")

  Else
    EditTableOf = strTable

  End If

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditTableOf

Public Function FieldValue(strTable, strField)
Dim strCurrentField

  strCurrentField = Session("CurrentField")
  If EditFindField(strTable, strField) Then
    FieldValue = Session(strTable & Session("CurrentField") & "Value")

  Else
    FieldValue = ""

  End If
  Session("CurrentField") = strCurrentField

End Function

%>