<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Data.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Dicionario de Dados
REM  Cria‡ao     : 10:00h 16/12/1998
REM  Local       : Brasilia/DF
REM  Elaborado   : Ueslei da Silva Reis <ueslei@zevallos.com.br>
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
REM Cria uma Tabela
REM -------------------------------------------------------------------------
Public Sub TableCreate(ByVal strTable)
Dim sql
Dim i, strType

  sql = ""
  
  strTable = Trim(strTable)
 
  If IsTable(strTable) Then
    TableDrop strTable
    
  End If
  
  For i = 1 To Session(strTable & "FieldNumber")

    If i > 1 Then
      sql = sql & ", "
      
    End If
  
    sql = sql & Session(strTable & i & "Field") & " "
    
    strType = LCase(Session(strTable & i & "Type"))
    
    Select Case strType
      Case "int", "money", "numeric", "datetime", "float", "real", "tinyint", "smallint", "bit", "text"
        sql = sql & strType
      
      Case Else
        sql = sql & Session(strTable & i & "Type") & "(" & Session(strTable & i & "Size") & ")"
      
    End Select

    If Not Session(strTable & i & "Null") Then
      sql = sql & " NOT"
      
    End If
    
    sql = sql & " NULL"

  Next

  sql = "CREATE TABLE " & strTable & " (" & sql & ")"
  
  ConnExecute sobjConn, sql

  ShowMessage "Criada a tabela " & strTable
  
  IndexCreate strTable
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub CreateTable

REM -------------------------------------------------------------------------
REM Exclui uma Tabela
REM -------------------------------------------------------------------------
Public Sub TableDrop(ByVal strTable)

  On Error Resume Next

  ConnExecute sobjConn, "DROP TABLE " & strTable

  On Error Goto 0

  ShowMessage "Excluida a tabela " & strTable

End Sub
REM -------------------------------------------------------------------------
REM End Sub DropTable

REM -------------------------------------------------------------------------
REM Verifica a Existencia de uma Tabela
REM -------------------------------------------------------------------------
Public Function IsTable(ByVal strTable)
Dim objRS

  sobjCMD.CommandText = "SELECT * FROM SysObjects Where Type='U' AND Name=" & SQLString(strTable)
  sobjCMD.CommandType = adCmdText
  
  Set objRS = Server.CreateObject("ADODB.RecordSet")
  objRS.Open sobjCMD, , adOpenDynamic, adLockPessimistic

  IsTable = False
    
   If Not objRS.EOF AND Not objRS.BOF Then
     IsTable = True

  End If  

  objRS.Close
  Set objRS = nothing

End Function
REM -------------------------------------------------------------------------
REM End Function IsTable

REM -------------------------------------------------------------------------
REM Atualiza uma Tabela
REM -------------------------------------------------------------------------
Public Sub TableUpdateBegin(ByVal strTable)

  Session("CurrentTable") = strTable                                                                                                                                              
  Session("FieldNumberUpdate") = 0
  Session("IndexNumberUpdate") = 0

End Sub
REM -------------------------------------------------------------------------
REM End Sub TableUpdateBegin

REM -------------------------------------------------------------------------
REM Atualiza uma Tabela
REM -------------------------------------------------------------------------
Public Sub FieldUpdateValue(ByVal strField, ByVal strValue)
                                                                                                                                              
  Session("FieldNumberUpdate") = Session("FieldNumberUpdate") + 1
  Session("FieldUpdate" & Session("FieldNumberUpdate")) = strField
  Session("FieldValue" & Session("FieldNumberUpdate")) = strValue
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub FieldUpdateValue

REM -------------------------------------------------------------------------
REM Atualiza uma Tabela
REM -------------------------------------------------------------------------
Public Sub FieldUpdateIndex(ByVal strField, ByVal strOperator, ByVal strValue)
                                                                                                                                              
  Session("IndexNumberUpdate") = Session("IndexNumberUpdate") + 1
  Session("FieldIndex" & Session("IndexNumberUpdate")) = strField
  Session("IndexOperator" & Session("IndexNumberUpdate")) = strOperator
  Session("IndexValue" & Session("FieldNumberUpdate")) = strValue
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub FieldUpdateIndex

REM -------------------------------------------------------------------------
REM Atualiza uma Tabela
REM -------------------------------------------------------------------------
Public Sub TableUpdateEnd
Dim objRS
Dim sql
Dim i
                                                                                                                                              
  sql = "UPDATE " & Session("CurrentTable") & " SET "

  For i = 1 To Session("FieldNumberUpdate")
    If i > 1 Then
      sql = sql & " AND "
    End If
    sql = sql & Session("FieldUpdate" & i) & " = " & Session("FieldValue" & i)
  Next

  If Session("IndexNumberUpdate") > 0 Then
     
    SetPutAnd  False
    
    For i = 1 To Session("IndexNumberUpdate")
      sql = sql & PutAnd & Session("FieldIndex" & i) & " " & Session("IndexOperator" & i) & " " & Session("IndexValue" & i)
    Next
     
  End If                                                                     
  
  ShowMessageError sql

  sobjCMD.CommandText = sql
  sobjCMD.CommandType = adCmdText
  
  Set objRS = Server.CreateObject("ADODB.RecordSet")
  objRS.Open sobjCMD, , adOpenDynamic, adLockPessimistic

End Sub
REM -------------------------------------------------------------------------
REM End Sub UpdateTableEnd

REM -------------------------------------------------------------------------
REM Cria um Indice
REM -------------------------------------------------------------------------
Public Sub IndexCreate(ByVal strTable)
  Dim i
  
  If Session(strTable & "IndexClusteredName") > "" Then
    On Error Resume Next
REM    ConnExecute sobjConn, "DROP INDEX " & strTable & "." & strTable & "_" & Session(strTable & "IndexClusteredName")

    On Error Goto 0

    ConnExecute sobjConn, "CREATE CLUSTERED INDEX " & strTable & "_" & Session(strTable & "IndexClusteredName") & " ON " & _
              strTable & "(" & Session(strTable & "IndexClusteredContent") & ")"

    ShowMessage "Criado o Indice " & strTable & "_" & Session(strTable & "IndexClusteredName")
  
  End If  
  
  If Session(strTable & "IndexNumber") > 0 Then
    For i = 1 To Session(strTable & "IndexNumber")
      On Error Resume Next
REM      ConnExecute sobjConn, "DROP INDEX " & strTable & "." & strTable & "_" & Session(strTable & "IndexName")

      On Error Goto 0
  
      ConnExecute sobjConn, "CREATE INDEX " & strTable & "_" & Session(strTable & i & "IndexName") & " ON " & _
                strTable & "(" & Session(strTable & i & "IndexContent") & ")"
  
      On Error Resume Next
  
      ShowMessage "Criado o Indice " & strTable & "_" & Session(strTable & i & "IndexName")
      
    Next
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub CreateIndex

REM -------------------------------------------------------------------------
REM Exclui um Indice
REM -------------------------------------------------------------------------
Public Sub IndexDrop(ByVal strIndex)

End Sub
REM -------------------------------------------------------------------------
REM End Sub DropIndice

REM -------------------------------------------------------------------------
REM Atualiza um Indice
REM -------------------------------------------------------------------------
Public Sub IndexUpdate(ByVal strIndex)

End Sub
REM -------------------------------------------------------------------------
REM End Sub UpdateIndice

REM -------------------------------------------------------------------------
REM Mostra o Registro Corrente
REM -------------------------------------------------------------------------
Public Sub ShowCurrentRecord(ByVal objRS)
  Dim i
  
  ShowHTMLCR "<P>"

  On Error Resume Next
    
  For i = 0 to objRS.Fields.Count - 1
    ShowHTML "<b>" & objRS(i).Name & "</b>(" & objRS(i).Type & ")=" & objRS(i).Value & "<BR>"
      
  Next

  On Error Goto 0

End Sub
REM -------------------------------------------------------------------------
REM End Sub ShowCurrentRecord
%>  