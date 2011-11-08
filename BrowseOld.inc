<SCRIPT LANGUAGE=VBScript RUNAT=Server> 
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

Const conBrowsePage     = "BP"
Const conBrowsePageSize = "BS"

Const conBrowseDefault = 20

REM -------------------------------------------------------------------------
REM Mostra os botoes de browse
REM -------------------------------------------------------------------------
Sub BrowseSet(ByRef objRS)
  Dim IntPage, intPageSize
  Dim strPage, strPageSize

  If Not objRS.Bof And Not objRS.Eof Then
    If IsNumeric(Session("BrowsePage")) Then
      intPage = Session("BrowsePage")
      
    Else
      intPage     = 1

      Session("BrowsePage") = 1
      
    End If

    If IsNumeric(Session("BrowsePageSize")) Then
      intPageSize = Session("BrowsePageSize")
      
    Else
      intPageSize = conBrowseDefault
      
      Session("BrowsePagePageSize") = conBrowseDefault
      
    End If

    strPageSize = Trim(Request.Form(conBrowsePageSize))
    strPage     = Trim(Request.QueryString(conBrowsePage))
  
    If Not strPageSize > "" Then
      strPageSize = Trim(Request.QueryString(conBrowsePageSize))
      
    End If

    If strPageSize > "" Then
      intPageSize = Int(strPageSize)
      
    End If

    If strPage > "" Then
      intPage = Int(strPage)
      
    End If
      
    If intPage > 0 Then
      Session("BrowsePage") = Int(strPage)
      
    End If

    If intPageSize > 0 Then
      Session("BrowsePageSize") = intPageSize
      
    End If

    objRS.PageSize = intPageSize
  
    objRS.AbsolutePage = intPage
  
    Session("Browse") = True

  Else

    Session("Browse") = False
        
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSet

REM -------------------------------------------------------------------------
REM Mostra os botoes de browse
REM -------------------------------------------------------------------------
Public Function BrowseCounter

  BrowseCounter = True

  If Session("Browse") Then
    If Not Session("BrowsePageSize") > "" Then
      Session("BrowseCounter") = 0
      
    End If

    Session("BrowseCounter") = Session("BrowseCounter") + 1
    
    If Session("BrowseCounter") > Session("BrowsePageSize") Then
      BrowseCounter = False
      
    End If                 
  End If
  
End Function
REM -------------------------------------------------------------------------
REM End Sub BrowseCounter

REM -------------------------------------------------------------------------
REM Mostra os botoes de browse
REM -------------------------------------------------------------------------
Public Sub BrowseBar(ByVal objRS)
  Dim i
  Dim intFirstPage, intLastPage
  Dim intPageSize, intPage, intRecordCount 
  
  intPageSize    = Session("BrowsePageSize")
  intPage        = Session("BrowsePage")
  intRecordCount = Session("BrowseCounter")

  If Session("Browse") Then
    ShowHTML "<p><b>(" & intRecordCount - 1 & " Linha"
    
    If intRecordCount > 1 Then ShowHTML "s"
    
    ShowHTML " apresentada"
    
    If intRecordCount > 1 Then ShowHTML "s"
    
    ShowHTMLCR ")</b>"                           

    intLastPage  = objRS.PageCount
  
    intFirstPage = Int(intPage / intPageSize) * intPageSize + 1
    
    If intLastPage - intFirstPage > intPageSize Then
      intLastPage = intFirstPage + intPageSize - 1
      
    End If
    
    If intLastPage > objRS.PageCount Then
      intLastPage = objRS.PageCount
      
    End If
    
    Dim strAHRef
    
    strAHRef = "<A HREF=" & sstrScriptName & "?"
    
    If sstrThisQueryString > "" Then
      strAHRef = strAHRef & sstrThisQueryString & "&"
      
    End If
    
    ShowHTMLCR "<br>P&aacute;gina Atual: " & intPage & "<P>"
    
    ShowHTMLCR "<form method=POST action=" & sstrScriptName & Request.ServerVariables("QUERY_STRING") & ">"
    ShowHTML "<b>Tamanho da página:</b> <input type=text name=" & conBrowsePageSize & " size=2 value=" & intPageSize & ">"
    ShowHTML "<input type=submit Name=cmdbrowse value=Atualiza><BR>"
  
    If intPage > 1 Then
      ShowHTML strAHRef &  conBrowsePageSize & "=" & intPage + 1 & ">PgUp</A>"
      
    End If
  
    If intPage < objRS.PageCount Then
      ShowHTML strAHRef &  conBrowsePageSize & "=" & intPage - 1 & ">PgDn</A>"
      
    End If
  
    ShowHTML "&nbsp;"
  
    For i = intFirstPage To intLastPage
      ShowHTML strAHRef &  conBrowsePageSize & "=" & i & ">" & i & "</A>"
      
    Next 
  
    If intLastPage < objRS.PageCount Then
      ShowHTML strAHRef &  conBrowsePageSize & "=" & intLastPage + 1 & ">" & intLastPage + 1 & "</A>"
      
    End If
  
    If intFirstPage > 0 Then
      ShowHTML strAHRef &  conBrowsePageSize & "=1>1</A>"
     
    End If
  
    If intLastPage < objRS.PageCount Then
      ShowHTML strAHRef &  conBrowsePageSize & "=" & objRS.PageCount + 1 & ">" & objRS.PageCount & "</A>"
      
    End If
  
    ShowHTMLCR "<br><input type=submit Name=cmdNewQuery value=""Nova Consulta"">"
  
    ShowHTMLCR "</form>"
    
  End If
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseBar

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Function IsBrowse
End Function
REM -------------------------------------------------------------------------
REM End Function IsBrowse

REM -------------------------------------------------------------------------
REM Mostra os botoes de browse
REM -------------------------------------------------------------------------
Public Function BrowseCounter
End Sub
REM -------------------------------------------------------------------------
REM End Sub BrowseSet


REM -------------------------------------------------------------------------
REM Fim do Browse.inc
</SCRIPT>