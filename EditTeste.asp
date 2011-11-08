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
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataBegin

REM -------------------------------------------------------------------------
REM Rotina que adiciona um campo da tabela no Dic. de Dados
REM -------------------------------------------------------------------------
Public Sub DataAddField(ByVal strField, ByVal strType, ByVal intSize, ByVal blnIsNull)
Dim i
 
  Session(Session("CurrentTable") & "FieldNumber") = Session(Session("CurrentTable") & "FieldNumber") + 1

  i = Session(Session("CurrentTable") & "FieldNumber")
  
  Session(Session("CurrentTable") & i & "Field")  = strField
  Session(Session("CurrentTable") & i & "Size")   = intSize
  Session(Session("CurrentTable") & i & "Type")   = strType
  Session(Session("CurrentTable") & i & "IsNull") = blnIsNull
  
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
    Err.Raise 300, "EditDataAddPrimaryKey", "O campo """ & strField & """, da tabela """ & Session("CurrentTable") & """, não existe"

  End If  

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditDataAddPrimaryKey

REM -------------------------------------------------------------------------
REM Rotina que adiciona um relacionamento entre tabelas
REM -------------------------------------------------------------------------
Public Sub DataAddRelation(ByVal strTable, ByVal intDeleteOption, ByVal intUpdateOption)
Dim i
 
  Session(Session("CurrentTable") & "RelationNumber")     = Session(Session("CurrentTable") & "RelationNumber") + 1

  i = Session(Session("CurrentTable") & "RelationNumber")
  
  Session(Session("CurrentTable") & i & "RelationTable")  = strTable
  Session(Session("CurrentTable") & i & "DeleteOption")   = intDeleteOption
  Session(Session("CurrentTable") & i & "UpdateOption")   = intUpdateOption
  Session(Session("CurrentTable") & i & "FieldtoFieldNumber") = 0
  
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddRelation

REM -------------------------------------------------------------------------
REM Rotina que adiciona um campo da tabela no Dic. de Dados
REM -------------------------------------------------------------------------
Public Sub DataAddRelationFields(ByVal strField, ByVal strField2)
Dim i, j
                                                    
  i = Session(Session("CurrentTable") & "RelationNumber")
  
  j = Session(Session("CurrentTable") & i & "FieldtoFieldNumber")
  
  Session(Session("CurrentTable") & i & "FieldtoFieldNumber") = j + 1
  
  j = j + 1 
  
  Session(Session("CurrentTable") & i & "," & j & "RelationField1") = strField
  Session(Session("CurrentTable") & i & "," & j & "RelationField2") = strField2
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub DataAddRelationFields

REM -------------------------------------------------------------------------
REM Rotina que localiza um campo em qualquer tabela
REM -------------------------------------------------------------------------
Public Function DataFindField(ByVal strField)
Dim i, strCurrentTable

  i = 1
  
  strCurrentTable = Session("CurrentTable")
  
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