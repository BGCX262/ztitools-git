<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Database.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Databases
REM  Cria‡ao     : 11:23h 7/2/1998
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------
REM  ALTERACOES
REM -------------------------------------------------------------------------
REM  Responsavel : [Nome do executante da alteracao]
REM  Data/Hora   : [Data e hora da alteracao]
REM  Resumo      : [Resumo descritivo da alteracao executada]
REM -------------------------------------------------------------------------

Dim sblnPutAnd

Public Sub Close(obj)    
  If obj.State = 1 Then
    obj.Close

  End If

End Sub

REM -------------------------------------------------------------------------
REM Devolve um complemento para um SQL
REM -------------------------------------------------------------------------
Private Function PutAnd

  If sblnPutAnd Then 
    PutAnd = " AND "
    
  Else
    PutAnd = " WHERE "
    
  End If

  sblnPutAnd = True
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PutAnd

REM -------------------------------------------------------------------------
REM Inicia um Sql
REM -------------------------------------------------------------------------
Private Sub SetPutAnd(ByVal blnValor)
  sblnPutAnd = blnValor
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetPutAnd

REM -------------------------------------------------------------------------
REM Devolve um complemento para um SQL
REM -------------------------------------------------------------------------
Private Function PutOr

  If sblnPutAnd Then 
    PutOr = " OR "
    
  Else
    PutOr = " WHERE "
    
  End If

  sblnPutAnd = True
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PutAnd

REM -------------------------------------------------------------------------
REM Inicia um Sql
REM -------------------------------------------------------------------------
Private Sub SetPutOr(ByVal blnValor)
  sblnPutAnd = blnValor
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetPutAnd

REM -------------------------------------------------------------------------
REM Lista todos os Fields de um Connection e de uma query SQL
REM -------------------------------------------------------------------------
Public Sub ListFieldsSQL(ByVal objConnListFields, ByVal strSQL)
  Dim objRSListFields
  
  Set objRSListFields = objConnListFields.Execute(strSQL)
  
  ListFields(objRSListFields)
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub InicioSistema

REM -------------------------------------------------------------------------
REM Lista todos os Fields de um RecordSet
REM -------------------------------------------------------------------------
Public Sub ListFields(ByVal objRSSL)
  Dim i
  
  On Error Resume Next
  
  ShowHTMLCR "<table border=1>"
  ShowHTMLCR "<tr bgcolor=lightblue><td colspan=4>Fields: " & objRSSL.Fields.Count & "</td></tr>"
  ShowHTMLCR "<tr bgcolor=lightblue><td colspan=4>Records: " & objRSSL.RecordCount & "</td></tr>"

  ShowHTMLCR "<tr bgcolor=lightblue><td>Id</td><td>Name</td><td>Type</td><td>Value</td></tr>"

  On Error Goto 0

  For i = 0 to objRSSL.Fields.Count - 1
    If i Mod 2 = 0 Then
      ShowHTML "<tr bgcolor=white>"
      
    Else
      ShowHTML "<tr bgcolor=lightyellow>"
      
    End If
    
    ShowHTMLCR "<td>" & i & _
               "</td><td>" & objRSSL(i).Name & _
               "</td><td>" & objRSSL(i).Type & _
               "</td><td>" & objRSSL(i).Value & _
               "</td></tr>"
    
  Next
  
  ShowHTMLCR "</table>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ListFields

REM -------------------------------------------------------------------------
REM Lista os registros do RecordSet
REM -------------------------------------------------------------------------
Public Sub ListFieldsCurrent(ByVal objRSSL, ByVal i)

    ShowMessage "Id=" & i & _
                " - Name=" & objRSSL(i).Name & _
                " - Type=" & objRSSL(i).Type & _
                " - Value=" & objRSSL(i).Value

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ListFieldsCurrent

REM -------------------------------------------------------------------------
REM Lista os registros do RecordSet
REM -------------------------------------------------------------------------
Public Sub Lista(ByVal RS2)
  Dim x, i
  
    ShowHTMLCR "<p>"
    ShowHTMLCR "<table border=1 cellpadding=0 cellspacing=0>"
    ShowHTMLCR "<tr><td><b>Reg</b>"

    For i = 0 to RS2.Fields.Count - 1
      ShowHTML "<td><b>" & RS2(i).Name & "</b>"
      
    Next

    RS2.MoveFirst

    x = 0

    Do While Not RS2.EOF
      x = x + 1
      
      ShowHTMLCR "<TR><TD VALIGN=TOP>" & x
      
      For i = 0 to RS2.Fields.Count - 1
        ShowHTML "<TD VALIGN=TOP>" & RS2(i)
        
      Next

      RS2.MoveNext

    Loop

    RS2.Close
    
    ShowHTMLCR "</table>"

    ShowMessage "Registros Listados!!"

End Sub


REM -------------------------------------------------------------------------
REM Fim da Sub Lista

REM -------------------------------------------------------------------------
REM Grava em uma coluna do Record Set
REM -------------------------------------------------------------------------
Public Sub RSFieldLet(ByVal objConn, ByVal objRS, ByVal strField, ByVal strResult)
  On Error Resume Next
  
  objRS(strField) = strResult
  
  If ErrorConnHandler(objConn) <> 0 Then
    If objRS.EditMode Then
      objRS.CancelUpdate
      
    End If

    ShowMessageError "Field=" & strField & " - Result='" & strResult & "' - Type=" & VarType(strResult) & " (" & TypeName(strResult) & ") - Lenght=" & Len(strResult)

    Err.Clear

    On Error Goto 0

    objRS(strField) = strResult

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub RSFieldLet

REM -------------------------------------------------------------------------
REM Grava em uma coluna do Record Set
REM -------------------------------------------------------------------------
Public Sub RSUpdate(ByVal objConn, ByVal objRS)
  On Error Resume Next
  
  objRS.Update
  
  If ErrorConnHandler(objConn) <> 0 Then
    If objRS.EditMode Then
      objRS.CancelUpdate
      
    End If

    Err.Clear

    On Error Goto 0

    objRS.Update

  End If


End Sub
REM -------------------------------------------------------------------------
REM Final da Sub RSUpdate

REM -------------------------------------------------------------------------
REM Abre um RS
REM -------------------------------------------------------------------------
Public Sub RSOpen(ByVal objConn, ByRef objRS, ByVal sql, ByVal strPar1, ByVal strPar2)
  On Error Resume Next
  
  objRS.Open sql, objConn, strPar1, strPar1
  
  If ErrorConnHandler(objConn) <> 0 Then
    ShowItemListBegin "100%", 1
    ShowItemListError "SQL", sql
    ShowItemListError "objConn", objConn
    ShowItemListError "Parametro1", strPar1
    ShowItemListError "Parametro2", strPar2
    ShowItemListEnd
    
    Err.Clear

    On Error Goto 0

    objRS.Open sql, objConn, strPar1, strPar1

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub RSOpen

REM -------------------------------------------------------------------------
REM Abre um RS com o command
REM -------------------------------------------------------------------------
Public Sub RSOpenCommand(ByRef objRS, ByVal objCMD, ByVal objConn, ByVal strPar1, ByVal strPar2)
  On Error Resume Next
  
  objRS.Open objCMD, , strPar1, strPar1

  If ErrorConnHandler(objConn) <> 0 Then
    ShowItemListBegin "100%", 1
    ShowItemListError "objConn", objConn
    ShowItemListError "objCMD.CommandText", objCMD.CommandText
    ShowItemListError "Parametro1", strPar1
    ShowItemListError "Parametro2", strPar2
    ShowItemListEnd
    
    Err.Clear

    On Error Goto 0

    objRS.Open objCMD, , strPar1, strPar1

  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub RSOpen

REM -------------------------------------------------------------------------
REM Grava em uma coluna do Record Set
REM -------------------------------------------------------------------------
Public Sub RSConnExecute(ByVal objConn, ByVal sql, ByRef objRS)
  On Error Resume Next
  
  Set objRS = objConn.Execute(sql)
  
  If ErrorConnHandler(objConn) <> 0 Then
    ShowItemListBegin "100%", 1
    ShowItemListError "SQL", sql
    ShowItemListError "objConn", objConn
    ShowItemListEnd
    
    Err.Clear

    On Error Goto 0

    Set objRS = objConn.Execute(sql)

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub RSConnExecute

REM -------------------------------------------------------------------------
REM Executa um Connection Execute
REM -------------------------------------------------------------------------
Public Sub ConnExecute(ByVal objConn, ByVal sql)
  On Error Resume Next
  
  objConn.Execute(sql)
  
  If ErrorConnHandler(objConn) <> 0 Then
    ShowItemListBegin "100%", 1
    ShowItemListError "SQL", sql
    ShowItemListError "objConn", objConn
    ShowItemListEnd
    
    Err.Clear

    On Error Goto 0

    objConn.Execute(sql)

  End If

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ConnExecute

REM -------------------------------------------------------------------------
REM Criacao de tabelas
REM -------------------------------------------------------------------------
Public Sub CreateTable(ByVal strSQL, ByVal strTableName)

  Session("CreateTableTableName") = strTableName
  
  On Error Resume Next

  ConnExecute sobjConn, "DROP TABLE " & strTableName

  On Error Goto 0

  strSQL = "CREATE TABLE " & strTableName & " (" & strSQL & ")"

  ConnExecute sobjConn, strSQL

  ShowMessage "Criada a tabela " & strTableName

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub CreateTable

REM -------------------------------------------------------------------------
REM Criacao de Indices Clustered
REM -------------------------------------------------------------------------
Public Sub CreateIndex(ByVal strIndexName, ByVal strFields)

  On Error Resume Next

REM  ConnExecute sobjConn, "DROP INDEX " & strIndexName & "_" &  Session("CreateTableTableName")

  On Error Goto 0

  ConnExecute sobjConn, "CREATE INDEX " & strIndexName & "_" &  Session("CreateTableTableName") & " ON " & _
              Session("CreateTableTableName") & "(" & strFields & ")"

  ShowMessage "Criado o Indice " & strIndexName & "_" & Session("CreateTableTableName")

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub CreateIndex

REM -------------------------------------------------------------------------
REM Criacao de Indices Clustered
REM -------------------------------------------------------------------------
Public Sub CreateIndexClustered(ByVal strIndexName, ByVal strFields)

  On Error Resume Next

  ConnExecute sobjConn, "DROP TABLE " & strTableName

  On Error Goto 0

  ConnExecute sobjConn, "CREATE CLUSTERED INDEX " & strIndexName & "_" &  Session("CreateTableTableName") & " ON " & _
              Session("CreateTableTableName") & "(" & strFields & ")"

  ShowMessage "Criado o Indice " & strIndexName & "_" & Session("CreateTableTableName")

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub CreateIndex

REM -------------------------------------------------------------------------
REM Fim do Database.inc
%>