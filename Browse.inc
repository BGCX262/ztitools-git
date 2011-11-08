<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Browse.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Browse
REM  Cria‡ao     : 9/21/98 5:59AM
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

REM -------------------------------------------------------------------------
REM Constantes de Navegacao
REM -------------------------------------------------------------------------
  Const conBrowseOption     = "BO"
  Const conBrowseTarget     = "BT"
  Const conBrowseActualPage = "BA"
  Const conBrowseEncripted  = "BX"

  Const conBrowsePage     = 1
  Const conBrowsePageSize = 2

  Const conBrowseDefaulPS = 10
  Const conBrowseBarPages = 10
   
REM -------------------------------------------------------------------------
REM Pega os parametros passados
REM -------------------------------------------------------------------------
  Public sparBrowseOption
  Public sparBrowseTarget
  Public sparBrowseActualPage
  Public sparBrowseEncripted

  sparBrowseOption     = lCase(Request.QueryString(conBrowseOption))
  sparBrowseTarget     = lCase(Request.QueryString(conBrowseTarget))
  sparBrowseActualPage = lCase(Request.QueryString(conBrowseActualPage))
  sparBrowseEncripted  = Request.QueryString(conBrowseEncripted)

  BrowseSetOptionError False

  If sparBrowseEncripted > "" Then
    BrowseSetOptionError False

    sparBrowseEncripted = UncriptNumber(sparBrowseEncripted)

    If IsUncriptNumberOK Then
      SetPiece sparBrowseEncripted
  
      sparBrowseActualPage = Int(GetPiece(4))
      sparBrowseOption     = Int(GetPiece(1))
      
      If Len(sparBrowseEncripted) > 5 Then
        sparBrowseTarget     = Int(GetPiece(4))
        
      End If    
    Else
      BrowseSetOptionError True
    
    End If
  End If
  
  If sparBrowseOption = conBrowsePageSize Then
    sparBrowseTarget = CInt(Request.Form("str" & conBrowsePageSize))
    
  End If

REM -------------------------------------------------------------------------
REM Constantes Locais
REM -------------------------------------------------------------------------
  Public sintBrowseLineCounter

  sintBrowseLineCounter = 0

REM -------------------------------------------------------------------------
REM Processa a navegacao do Browse
REM -------------------------------------------------------------------------
Public Sub BrowseBegin

  If Not IsBrowseOptionError Then
    If sstrThisQueryString > ""  And Not sparBrowseOption > "" And Not sparBrowseActualPage > "" Then
       BrowseSetQueryString sstrThisQueryString
        
    End If
  
    If IsBrowse Then
      Select Case sparBrowseOption
        Case conBrowsePage
          BrowseSetPage CInt(sparBrowseTarget)
  
        Case conBrowsePageSize
          BrowseSetPageSize CInt(sparBrowseTarget)
    
          BrowseSetPage CInt(sparBrowseActualPage)
          
        Case Else
      End Select
  
  REM    BrowseGetRS
  
      If Not sobjRS.Bof And Not sobjRS.Eof Then
        sobjRS.PageSize = BrowseGetPageSize
  
        If BrowseGetPage > sobjRS.PageCount Then
          BrowseSetPage sobjRS.PageCount
       
        ElseIf BrowseGetPage = "" Or BrowseGetPage < 1 Then
          BrowseSetPage 1
          
        End If
  
        sobjRS.AbsolutePage = BrowseGetPage
  
        BrowseSetStatus True
  
      Else
  
        BrowseSetStatus False
          
      End If
  
  REM    BrowseSetRS
  
    End If
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseProcess

REM -------------------------------------------------------------------------
REM Monta o AHRef com o dado criptografado
REM -------------------------------------------------------------------------
Private Sub BrowseAHRef(ByVal strAHRef, ByVal strEncript, ByVal strNumber, ByVal strText)
  ShowHTMLCR strAHRef & EncriptNumber(strEncript & LeadingZeroes(strNumber, 4)) & ">" & strText & "</A>"

End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseAHRef

REM -------------------------------------------------------------------------
REM Mostra a barra de navegacao
REM -------------------------------------------------------------------------
Public Sub BrowseBar
  Dim intLastPage, intFirstPage
  Dim intLessPages, intMorePages
  Dim strAHRef, strQuery
  Dim i
  Dim strEncript

  If Not IsBrowseOptionError Then
    ShowHTMLCR "<FONT SIZE=1>"
    
    If IsBrowse And sobjRS.PageCount > 1 Then
  
      intLastPage  = sobjRS.PageCount
  
      intFirstPage = Int((BrowseGetPage - 1) / BrowseGetBarPages) * BrowseGetBarPages + 1
  
      intLastPage = intFirstPage + BrowseGetBarPages - 1
  
      If intLastPage >= sobjRS.PageCount Then
        intLastPage = sobjRS.PageCount - intLastPage + BrowseGetBarPages + intFirstPage - 1
        
      End If
  
      REM ---------------------------------------------------------------------
      If BrowseGetPage + BrowseGetBarPages > sobjRS.PageCount Then
        intMorePages = sobjRS.PageCount - BrowseGetPage
        
      Else
        intMorePages = BrowseGetBarPages
        
      End If 
  
      REM ---------------------------------------------------------------------
      If BrowseGetPage - BrowseGetBarPages <= 0 Then
        intLessPages = intFirstPage
        
      Else
        intLessPages = BrowseGetBarPages
        
      End If 
    End If

    Dim strSTail
  
    strSTail = ""
  
    If BrowseGetLineCounter > 1 Then
      strSTail = "s"
      
    End If
    
    ShowHTMLCR "<BR>" & BrowseGetLineCounter & " Linha" & strSTail & " apresentada" & strSTail
      
    If BrowseGetPage < sobjRS.PageCount Then
      ShowHTMLCR "de " & sobjRS.RecordCount & " linhas"
    
    End If

    strSTail = ""
  
    If sobjRS.PageCount > 1 Then
      strSTail = "s"
      
    End If
    
    If sobjRS.PageCount > 1 Then
      ShowHTMLCR "<BR>na página " & BrowseGetPage
      ShowHTMLCR " de " & sobjRS.PageCount & " página" & strSTail & "<BR>"
      
    End If

    strQuery = sstrThisScriptName & "?"
  
    If isBrowseQueryString > "" Then    
      strQuery = strQuery & BrowseGetQueryString & "&"
      
    End If
  
    strQuery = strQuery & conBrowseEncripted & "="
  
    strEncript = LeadingZeroes(BrowseGetPage, 4)
                                                        
    ShowHTMLCR "<form method=POST action=" & strQuery & EncriptNumber(strEncript & conBrowsePageSize) & ">"
    ShowHTML "<b>Tamanho da página:</b> <input type=text name=str" & conBrowsePageSize & " size=3 value=" & BrowseGetPageSize & ">"
    ShowHTML "<input type=submit Name=cmdbrowse value=Atualizar><BR>"
    ShowHTMLCR "</form>"
    

    If IsBrowse And sobjRS.PageCount > 1 Then
      strEncript = strEncript & conBrowsePage

      strAHRef = "<A HREF=" & strQuery
   
      If BrowseGetPage > 1 Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage - 1, "[Anterior]"
        
      Else
        ShowHTMLCR "[Anterior]"
        
      End If
  
      If BrowseGetPage < sobjRS.PageCount Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage + 1, "[Próxima]"
        
      Else
        ShowHTMLCR "[Próxima]"
        
      End If
  
      ShowHTMLCR "&nbsp;"
  
      If intFirstPage > 1 Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage - intLessPages, "[Menos " & intLessPages & "]"
  
      Else
        ShowHTMLCR "&nbsp;[Menos 10]"
        
      End If
  
      ShowHTMLCR "&nbsp;"
  
      For i = intFirstPage To BrowseGetPage - 1
        BrowseAHRef strAHRef, strEncript, i, i
  
      Next 
  
      ShowHTMLCR "<FONT COLOR=Red><B>" & BrowseGetPage & "</B></FONT>"
    
      For i = BrowseGetPage + 1 To intLastPage
        BrowseAHRef strAHRef, strEncript, i, i
  
      Next 
  
      ShowHTMLCR "&nbsp;"
      
      If intLastPage < sobjRS.PageCount Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage + intMorePages, "[Mais " & intMorePages & "]"
  
        Else
          ShowHTMLCR "&nbsp;[Mais 10]"
        
      End If
  
      ShowHTMLCR "&nbsp;"
    
      If BrowseGetPage > 1 Then
        BrowseAHRef strAHRef, strEncript, 1, "[Primeira]"
  
      Else
        ShowHTMLCR "[Primeira]"
       
      End If
    
      If BrowseGetPage < sobjRS.PageCount Then
        BrowseAHRef strAHRef, strEncript, sobjRS.PageCount, "[Última]"
  
      Else    
        ShowHTMLCR "[Última]"
        
      End If
    
    End If
  
    ShowHTMLCR "</FONT>"
    
  Else
    ShowMessageError "ZTIBrowser ERROR: Parametros Invalidos ou Adulterados!"

    strQuery = sstrThisScriptName & "?"
  
    If isBrowseQueryString > "" Then    
      strQuery = strQuery & BrowseGetQueryString & "&"
        
    End If

    strQuery = strQuery & conBrowseEncripted & "="

    strEncript = LeadingZeroes(BrowseGetPage, 4)

    strEncript = strEncript & conBrowsePage
  
    strAHRef = "<A HREF=" & strQuery

    ShowHTMLCR "<BR>"

    BrowseAHRef strAHRef, strEncript, 1, "[Primeira Página]"
  
  End If
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseBar

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Sub BrowseSetActual(ByVal strActual)
                                 
  Session("BrowseActual") = strActual
  
  Session("BrowseLast") = strActual

  If Not IsBrowse > "" Then
    BrowseSetStatus False
    
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetActual

REM -------------------------------------------------------------------------
REM Define quantidade de paginas na barra 
REM -------------------------------------------------------------------------
Public Sub BrowseSetBarPages(ByVal intBarPages)
 
  Session("BrowseBarPages") = intBarPages
 
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetBarPages

REM -------------------------------------------------------------------------
REM Define quantidade de paginas na barra 
REM -------------------------------------------------------------------------
Public Function BrowseGetBarPages

  If Not CInt(Session("BrowseBarPages")) > 1 Then
    BrowseSetBarPages conBrowseBarPages
    
  End If
 
  BrowseGetBarPages = Session("BrowseBarPages") 
 
End Function
REM -------------------------------------------------------------------------
REM End Function BrowseGetBarPages

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Function BrowseGetActual

  BrowseGetActual = Session("BrowseActual")
  
End Function
REM -------------------------------------------------------------------------
REM End Function BrowseGetActual

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Sub BrowseSetLast

  BrowseSetActual Session("BrowseLast")
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetLast

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Function IsBrowse

  If Not BrowseGetActual > "" Then
      BrowseSetLast
      
  End If

  IsBrowse = Session("BrowseActive" & Session("BrowseActual"))
  
End Function
REM -------------------------------------------------------------------------
REM End Function IsBrowse

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Sub BrowseSetStatus(ByVal blnStatus)

  Session("BrowseActive" & Session("BrowseActual")) = blnStatus
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetStatus

REM -------------------------------------------------------------------------
REM Define o tamanho da pagina
REM -------------------------------------------------------------------------
Public Sub BrowseSetPageSize(ByVal intPageSize)

  Session("BrowsePageSize" & Session("BrowseActual")) = intPageSize
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetLast

REM -------------------------------------------------------------------------
REM Retorna o tamanho da pagina
REM -------------------------------------------------------------------------
Public Function BrowseGetPageSize

  BrowseGetPageSize = CInt(Session("BrowsePageSize" & Session("BrowseActual")))
  
  If Not BrowseGetPageSize > 0 Or Not IsNumeric(BrowseGetPageSize) Then
    BrowseGetPageSize = conBrowseDefaulPS
    
  End If
  
End Function
REM -------------------------------------------------------------------------
REM End Function BrowseGetPageSize

REM -------------------------------------------------------------------------
REM Verifica se terminou a apresentacao da pagina
REM -------------------------------------------------------------------------
Public Function IsBrowsePage

  IsBrowsePage = True

  If BrowseGetLineCounter >= BrowseGetPageSize Then
    IsBrowsePage = False

  End If    
  
End Function
REM -------------------------------------------------------------------------
REM End Function IsBrowsePage

REM -------------------------------------------------------------------------
REM Conta as linhas de apresentacao
REM -------------------------------------------------------------------------
Public Sub BrowseLineCounter

  sintBrowseLineCounter = sintBrowseLineCounter + 1
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseLineCounter

REM -------------------------------------------------------------------------
REM Conta as linhas de apresentacao
REM -------------------------------------------------------------------------
Public Function BrowseGetLineCounter

  BrowseGetLineCounter = sintBrowseLineCounter
  
End Function
REM -------------------------------------------------------------------------
REM End Sub BrowseGetLineCounter

REM -------------------------------------------------------------------------
REM Retorna o RS
REM -------------------------------------------------------------------------
Public Sub BrowseGetRS

  Set sobjRS = Session("BrowseRS" & Session("BrowseActual"))
  
End Sub
REM -------------------------------------------------------------------------
REM End Function BrowseGetRS

REM -------------------------------------------------------------------------
REM Define o RS
REM -------------------------------------------------------------------------
Public Sub BrowseSetRS

  Set Session("BrowseRS" & Session("BrowseActual")) = sobjRS

  BrowseSetStatus True
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetRS

REM -------------------------------------------------------------------------
REM Apaga todas as referencias do Browse demarcado
REM -------------------------------------------------------------------------
Public Sub BrowseClearActual
               
  BrowseSetStatus False
  BrowseSetActual ""
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseClearActual

REM -------------------------------------------------------------------------
REM Mostra o Cabecalho
REM -------------------------------------------------------------------------
Public Sub BrowseHeader

End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseHeader

REM -------------------------------------------------------------------------
REM Mostra o Rodape
REM -------------------------------------------------------------------------
Public Sub BrowseEnd

End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseFooter

REM -------------------------------------------------------------------------
REM Retorna a Pagina Atual
REM -------------------------------------------------------------------------
Public Function BrowseGetPage

  BrowseGetPage = CInt(Session("BrowsePage" & Session("BrowseActual")))
  
  If BrowseGetPage < 1 Then
    BrowseGetPage = 1
    
  End If
  
End Function
REM -------------------------------------------------------------------------
REM End Function BrowseGetRS

REM -------------------------------------------------------------------------
REM Define a Pagina Atual
REM -------------------------------------------------------------------------
Public Sub BrowseSetPage(ByVal intPage)

  Session("BrowsePage" & Session("BrowseActual")) = intPage
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetPage
  
REM -------------------------------------------------------------------------
REM Retorna o estado do Query String
REM -------------------------------------------------------------------------
Public Function IsBrowseQueryString

  IsBrowseQueryString = Session("BrowseQueryStringStatus" & Session("BrowseActual"))
  
  If IsBrowseQueryString > "" Then
    IsBrowseQueryString = False
    
    Session("BrowseQueryStringStatus" & Session("BrowseActual")) = False
    
  End If
  
End Function
REM -------------------------------------------------------------------------
REM End Function IsBrowseQueryString

REM -------------------------------------------------------------------------
REM Retorna o estado do Query String
REM -------------------------------------------------------------------------
Public Function BrowseGetQueryString

  BrowseGetQueryString = Session("BrowseQueryString" & Session("BrowseActual"))
  
End Function
REM -------------------------------------------------------------------------
REM End Function IsBrowseQueryString

REM -------------------------------------------------------------------------
REM Pega o Query String
REM -------------------------------------------------------------------------
Public Sub BrowseSetQueryString(ByVal strQueryString)

  Session("BrowseQueryString" & Session("BrowseActual")) = strQueryString
  
  Session("BrowseQueryStringStatus" & Session("BrowseActual")) = True
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetQueryString

REM -------------------------------------------------------------------------
REM Retorna o estado das Opcoes do Browse
REM -------------------------------------------------------------------------
Public Function IsBrowseOptionError

  IsBrowseOptionError = Session("BrowseOptionError" & Session("BrowseActual"))
  
  If Not IsBrowseOptionError > "" Then
    IsBrowseOptionError = False
    
    Session("BrowseOptionError" & Session("BrowseActual")) = False
    
  End If
  
End Function
REM -------------------------------------------------------------------------
REM End Function IsBrowseOptionError

REM -------------------------------------------------------------------------
REM Define o estado das Opcoes do Browse
REM -------------------------------------------------------------------------
Public Sub BrowseSetOptionError(ByVal blnOptionError)

  Session("BrowseOptionError" & Session("BrowseActual")) = blnOptionError
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSetOptionError
  
REM -------------------------------------------------------------------------
REM Fim do Browse.inc
%>