<%
REM -------------------------------------------------------------------------
REM  /ZTITools/EditForm/String.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Edicao
REM  Cria‡ao     : 11:10h 22/02/1999
REM  Local       : sobjShow.BRasilia/DF
REM  Elaborado   : Eduardo Alves Cortes <edualves@zevallos.com.br>
REM              : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

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

      Case "," , ";", "-", "=", "(", ")", "|", "#", ""
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
   
   strSubText = Replace(strSubText, "@~@", ".")
   strSubText = Replace(strSubText, "@!@", ";")
   strSubText = Replace(strSubText, "@?@", "=")
   strSubText = Replace(strSubText, "@@@", "-")
   strSubText = Replace(strSubText, "@%@", ",")
   EditNextOperation = strSubText
   sstrSubText = strSubText

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditNextOperation

REM -------------------------------------------------------------------------
REM Rotina que retorna o ultimo dado lida numa string
REM -------------------------------------------------------------------------
Private Function EditLastOperation

  EditLastOperation = sstrSubText

End Function
REM -------------------------------------------------------------------------
REM Final da Function EditLastOperation

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
  
  sstrSubText = ""                                                        
                                                          
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
      
      Case ";", "-", "|"
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

    If EditLastOperation = "=" Then
      EditNextOperation

    End If
   Loop

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

   Session(EditCurrentTable & strTotal) = intGreat

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditTotalOperation

Public Function PrepareForm(strText)
Dim strAux         
  strAux = strText
  strAux = Replace(strAux, ".", "@~@")
  strAux = Replace(strAux, ";", "@!@")
  strAux = Replace(strAux, "=", "@?@")
  strAux = Replace(strAux, "-", "@@@")
  strAux = Replace(strAux, ",", "@%@")
  PrepareForm = strAux

End Function

%>