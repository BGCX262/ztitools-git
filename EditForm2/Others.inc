<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/Others.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 11:30h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

Public sblnShowMessage

sblnShowMessage = False

Public Sub DebugMessage(strText, intMessageType)
  If sblnShowMessage Then
    If (intMessageType = 1) Or (Len(strText) > 30 And intMessageType = 0) Then
      sobjShow.MessageError strText
    
    Else
      sobjShow.JSAlert strText
      
    End If
  End If  

End Sub

Public Sub DebugMode
  sblnShowMessage = True
  SetShowCR True

End Sub

Public Function Filter
Dim strfilter, i
               
  SetPutAnd False

  If sparEditFilter > "" Then
    strfilter = PutAnd & sparEditFilter

  End If          

  For i = 1 to Session(EditCurrentTable & "FilterNumber")
    If EditFindField(EditCurrentTable, Session(EditCurrentTable & i & "FilterField")) Then
      DebugMessage Session(EditCurrentField & "Type"), 0
      strfilter = strfilter & PutAnd &  EditTableOf(EditCurrentTable, Session("CurrentField")) & "." & Session(EditCurrentTable & i & "FilterField") & "=" & FormatForSQL(Session(EditCurrentTable & i & "FilterValue"), Session(EditCurrentField & "Type"))
    
    End If
  Next 
  
  DebugMessage strfilter, 2
  Filter = strfilter

End Function

Private Sub ClearSession(ByVal strSession)
  If Session(strSession) > "" Then
    Session(strSession) = ""

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
REM Inclui uma imagem na tela
REM -------------------------------------------------------------------------
Private Sub ShowImage(strImagem, strID, strAlt, strHRef, strAlign, intWidth, intHeight)

  If strHRef > "" Then
    Response.Write "<A HREF=" & strHRef & ">"

  End If

  Response.Write "<IMG SRC=""" & strImagem & """"
  
  If strHRef > "" Then
    Response.Write " BORDER=0"

  End If

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
REM Retorna a divisao do primeiro valor pelo segundo arredondado para cima, tornando o resultado inteiro
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

Private Sub ClassAss                         
Dim strLink, strALink
  If Session("BodyLink") > "" Then                                           
    strLink = Session("BodyLink")

  Else
    strLink = "Black"

  End If             
  If Session("BodyALink") > "" Then
    strALink = Session("BodyALink")                                             
  
  Else
    strALink = "Purple"
  
  End If                                              
  sobjShow.HTMLCR "<STYLE>"
  sobjShow.HTMLCR ".SS {text-decoration: none; color: " & strLink & "; font:bold 8pt}"
  sobjShow.HTMLCR ".SS:HOVER {color: " & strALink & "}"
  sobjShow.HTMLCR "</STYLE>"

  sobjShow.HTMLCR "<script type=""text/javascript"">"
  sobjShow.HTMLCR "function OpenDesc(text) {"
  sobjShow.HTMLCR "  if (text.style.display =="""") {"
  sobjShow.HTMLCR "    text.style.display = ""none"";"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "  else {"
  sobjShow.HTMLCR "  text.style.display="""";"
  sobjShow.HTMLCR "  }"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "function makeCool(text) {"
  sobjShow.HTMLCR "  text.style.color = """ & strALink & """"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "function makeNormal(text) {"
  sobjShow.HTMLCR "  text.style.color = """ & strLink & """"
  sobjShow.HTMLCR "}"
  sobjShow.HTMLCR "</script>"

End Sub

REM -------------------------------------------------------------------------
REM Entra Campo File para download de arquivos
REM -------------------------------------------------------------------------
Sub FormInputFile(ByVal strName, ByVal intSize, ByVal strValue, ByVal intMaxLength)
  FormInput "File", strName, intSize, strValue, intMaxLength, "", 0, ""

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub FormInputText

REM -------------------------------------------------------------------------
REM Salva a QueryString Corrente
REM -------------------------------------------------------------------------
Public Sub QSSave

  Session("QSLast") = Session("QSCurrent")
  
  Session("QSCurrent") = sstrThisScriptName & "?" & sstrThisQueryString
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub QSSave

REM -------------------------------------------------------------------------
REM Recupera a QueryString Corrent
REM -------------------------------------------------------------------------
Public Function QSCurrent

  QSCurrent = Session("QSCurrent")

End Function
REM -------------------------------------------------------------------------
REM Final da Function QSSave

REM -------------------------------------------------------------------------
REM Recupera a QueryString Anterior
REM -------------------------------------------------------------------------
Public Function QSLast

  QSLast = Session("QSLast")

End Function
REM -------------------------------------------------------------------------
REM Final da Function QSLast
%>