<%
REM -------------------------------------------------------------------------
REM  /ZTITools/URL.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de URL
REM  Criacao     : 11:23h 7/2/1998
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

Dim sintURLCounter
Dim sstrURLLocation
Dim sarrURL(20)
Dim sblnAmpersand

sblnAmpersand = False


REM -------------------------------------------------------------------------
REM Inicia o URL
REM -------------------------------------------------------------------------
Private Sub URLBegin(ByVal strLocation)

  SetAmpersand False

  sstrURLLocation = strLocation

  sintURLCounter = 0

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub URLBegin

REM -------------------------------------------------------------------------
REM Termina o URL
REM -------------------------------------------------------------------------
Private Sub URLEnd

  sintURLCounter = 0

  SetAmpersand False

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub URLEnd

REM -------------------------------------------------------------------------
REM Adiciona a URL na Lista
REM -------------------------------------------------------------------------
Private Sub URLAdd(ByVal strOption, ByVal strValue)

  sintURLCounter = sintURLCounter + 1
  
  sarrURL(sintURLCounter) = strOption & "=" & Server.URLEncode(strValue)

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub URLAdd

REM -------------------------------------------------------------------------
REM Adiciona a URL na Lista
REM -------------------------------------------------------------------------
Private Function URLGet(ByVal strOption, ByVal strDescription)
  Dim strResult, i
  
  If Left(strOption, 1) = "&" Then
    strOption = Mid(strOption, 2)
    
  End If             

  strResult = "<A HREF=" & sstrURLLocation
  
  If strOption > "" Or sintURLCounter > 0 Then
    
    strResult = strResult & "?"
    
    If strOption > "" Then
      strResult = strResult & strOption
      
    End If   
  End If

  If sintURLCounter > 0 Then
    If strOption > "" Then
      strResult = strResult & "&"
      
    End If
    
    For i = 1 To sintURLCounter
      If i > 1 Then
        strResult = strResult & "&"
      
      End If

      strResult = strResult & sarrURL(i)
      
    Next
    
  End If
  
  strDescription = Trim(strDescription)
  
  If Not strDescription > "" Then
    strDescription = "Vazio"
    
  End If

  URLGet = strResult & ">" & strDescription & "</a>"

End Function
REM -------------------------------------------------------------------------
REM Final da Function URLGet

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit
REM -------------------------------------------------------------------------
Private Sub URLShow(ByVal strOption, ByVal strDescription)

  ShowHTMLCR URLGet(strOption, strDescription)

  SetAmpersand False
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub URLShow

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit
REM -------------------------------------------------------------------------
Private Function URLEqual(ByVal strOption, ByVal strValue)

  URLEqual = GetAmpersand & strOption & "=" & strValue
  
End Function
REM -------------------------------------------------------------------------
REM Final da Sub URLEqual

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit
REM -------------------------------------------------------------------------
Private Function GetAmpersand

  If sblnAmpersand Then
    GetAmpersand = "&"
    
  Else
    sblnAmpersand = True
    
  End If
  
End Function
REM -------------------------------------------------------------------------
REM Final da Sub GetAmpersand

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit
REM -------------------------------------------------------------------------
Private Sub SetAmpersand(ByVal blnValue)

  sblnAmpersand = blnValue
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub SetAmpersand

REM -------------------------------------------------------------------------
REM Inicia a criacao do URL - 19/01/99 22:28
REM -------------------------------------------------------------------------
Public Sub EditCreateURLBegin(ByVal strLocate)

  URLBegin strLocate

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLBegin

REM -------------------------------------------------------------------------
REM Termina a criacao do URL - 19/01/99 22:28
REM -------------------------------------------------------------------------
Public Sub EditCreateURLEnd

  URLEnd

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLEnd

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit - Ruben 1/18/99 12:57AM
REM -------------------------------------------------------------------------
Public Sub EditCreateURLTargetAdd(ByVal strTarget, ByVal strValue)

  URLAdd URLEqual(conPEditFind, strTarget), strValue

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLModeSet

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit - Ruben 1/18/99 12:57AM
REM -------------------------------------------------------------------------
Public Sub EditCreateURLTarget(ByVal strOption, ByVal strDescription)

  EditCreateURL strOption, strDescription

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURLModeSet

REM -------------------------------------------------------------------------
REM Monta URL para chamar o Edit - Ruben 1/18/99 12:57AM
REM -------------------------------------------------------------------------
Public Sub EditCreateURL(ByVal strOption, ByVal strDescription)

  URLShow URLEqual(conPEditOption, conOptionFPEdit) & URLEqual(conPEditWhat, strOption) & conPEditEditableStr, strDescription
 
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub EditCreateURL

REM -------------------------------------------------------------------------
REM Fim do URL.inc
%>