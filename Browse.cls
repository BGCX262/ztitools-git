VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Browse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

REM -------------------------------------------------------------------------
REM  Browse.cls
REM -------------------------------------------------------------------------
REM  Descricao   : Include do sistema de Browse
REM  Cria‡ao     : 9/21/98 5:59AM
REM  Local       : Brasilia/DF
REM  Elaborado   : Ruben Zevallos Jr. <zevallos@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1998, 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

'    Description = "ZTITools.Browse - Sistema de Browse - 1/31/99 10:47AM"
'    ProgId = "ZTITools.Browse"
'    Version = "1.01"
'    ClassID = "{E0408FF0-B920-11d2-B163-00A0C923ECAE}"

Option Explicit

REM -------------------------------------------------------------------------
REM Constantes de Navegacao
REM -------------------------------------------------------------------------
  Private Const conBrowseOption     = "BO"
  Private Const conBrowseTarget     = "BT"
  Private Const conBrowseActualPage = "BA"
  Private Const conBrowseEncripted  = "BX"

  Private Const conBrowsePage     = 1
  Private Const conBrowsePageSize = 2

  Private Const conBrowseDefaulPS = 10
  Private Const conBrowseBarPages = 10
   
REM -------------------------------------------------------------------------
REM Pega os parametros passados
REM -------------------------------------------------------------------------
  Private sparBrowseOption
  Private sparBrowseTarget
  Private sparBrowseActualPage
  Private sparBrowseEncripted

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

Private ScriptContext As ScriptingContext

Private sstrComments As String, _
    sstrTitle As String, _
    sstrVersion As String, _
    sstrCompanyName As String, _
    sstrProductName As String, _
    sstrLegalCopyright As String

Private Request As Request, _
    Response As Response, _
    Server As Server, _
    Session As Session, _
    Application As Application
    
    
Public Sub OnStartPage(ScriptContext As ScriptingContext)
  
  Set ScriptContext = ScriptContext
  
  Set Request = ScriptContext.Request
  Set Response = ScriptContext.Response
  Set Server = ScriptContext.Server
  Set Session = ScriptContext.Session
  Set Application = ScriptContext.Application
  
End Sub

Private Sub Class_Initialize()

    sstrComments = App.Comments
    sstrTitle = App.Title
    sstrVersion = App.Major & "." & _
                App.Minor & "." & _
                App.Revision
    sstrCompanyName = App.CompanyName
    sstrProductName = App.ProductName

End Sub

Public Property Get Comments() As String
    
    Comments = sstrComments
    
End Property

Public Property Get Title() As String
    
    Title = sstrTitle
    
End Property

Public Property Get Version() As String
    
    Version = sstrVersion
    
End Property
Public Property Get CompanyName() As String
    
    CompanyName = sstrCompanyName
    
End Property

Public Property Get ProductName() As String
    
    ProductName = sstrProductName
    
End Property

Public Property Get LegalCopyright() As String
    
    LegalCopyright = sstrLegalCopyright
    
End Property

REM -------------------------------------------------------------------------
REM Retorna a Pagina Atual
REM -------------------------------------------------------------------------
Public Property Get Page

  Page = CInt(Session("BrowsePage" & Session("BrowseActual")))
  
  If BrowseGetPage < 1 Then
    Page = 1
    
  End If
  
End Property

Public Property Let Page(ByVal intPage)

  Session("BrowsePage" & Session("BrowseActual")) = intPage
  
End Property
  
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
REM Retorna o estado do Query String
REM -------------------------------------------------------------------------
Public Function BrowseGetQueryString

  BrowseGetQueryString = Session("BrowseQueryString" & Session("BrowseActual"))
  
End Function

REM -------------------------------------------------------------------------
REM Pega o Query String
REM -------------------------------------------------------------------------
Public Sub BrowseSetQueryString(ByVal strQueryString)

  Session("BrowseQueryString" & Session("BrowseActual")) = strQueryString
  
  Session("BrowseQueryStringStatus" & Session("BrowseActual")) = True
  
End Sub

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
REM Define o estado das Opcoes do Browse
REM -------------------------------------------------------------------------
Public Sub BrowseSetOptionError(ByVal blnOptionError)

  Session("BrowseOptionError" & Session("BrowseActual")) = blnOptionError
  
End Sub

REM -------------------------------------------------------------------------
REM Processa a navegacao do Browse
REM -------------------------------------------------------------------------
Public Sub Begin

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
REM Monta o AHRef com o dado criptografado
REM -------------------------------------------------------------------------
Private Sub AHRef(ByVal strAHRef As String, _
                        ByVal strEncript As String, _
                        ByVal strNumber As String, _
                        ByVal strText As String)
                        
  Response.Write strAHRef & EncriptNumber(strEncript & LeadingZeroes(strNumber, 4)) & ">" & strText & "</A>"

End Sub

REM -------------------------------------------------------------------------
REM Mostra a barra de navegacao
REM -------------------------------------------------------------------------
Public Sub Bar
  Dim intLastPage, intFirstPage
  Dim intLessPages, intMorePages
  Dim strAHRef, strQuery
  Dim i
  Dim strEncript

  If Not IsBrowseOptionError Then
    Response.Write "<FONT SIZE=1>"
    
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
    
    Response.Write "<BR>" & BrowseGetLineCounter & " Linha" & strSTail & " apresentada" & strSTail
      
    If BrowseGetPage < sobjRS.PageCount Then
      Response.Write "de " & sobjRS.RecordCount & " linhas"
    
    End If

    strSTail = ""
  
    If sobjRS.PageCount > 1 Then
      strSTail = "s"
      
    End If
    
    If sobjRS.PageCount > 1 Then
      Response.Write "<BR>na página " & BrowseGetPage
      Response.Write " de " & sobjRS.PageCount & " página" & strSTail & "<BR>"
      
    End If

    strQuery = sstrThisScriptName & "?"
  
    If isBrowseQueryString > "" Then    
      strQuery = strQuery & BrowseGetQueryString & "&"
      
    End If
  
    strQuery = strQuery & conBrowseEncripted & "="
  
    strEncript = LeadingZeroes(BrowseGetPage, 4)
                                                        
    Response.Write "<form method=POST action=" & strQuery & EncriptNumber(strEncript & conBrowsePageSize) & ">"
    ShowHTML "<b>Tamanho da página:</b> <input type=text name=str" & conBrowsePageSize & " size=3 value=" & BrowseGetPageSize & ">"
    ShowHTML "<input type=submit Name=cmdbrowse value=Atualizar><BR>"
    Response.Write "</form>"
    

    If IsBrowse And sobjRS.PageCount > 1 Then
      strEncript = strEncript & conBrowsePage

      strAHRef = "<A HREF=" & strQuery
   
      If BrowseGetPage > 1 Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage - 1, "[Anterior]"
        
      Else
        Response.Write "[Anterior]"
        
      End If
  
      If BrowseGetPage < sobjRS.PageCount Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage + 1, "[Próxima]"
        
      Else
        Response.Write "[Próxima]"
        
      End If
  
      Response.Write "&nbsp;"
  
      If intFirstPage > 1 Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage - intLessPages, "[Menos " & intLessPages & "]"
  
        Else
          Response.Write "&nbsp;[Menos 10]"
        
      End If
  
      Response.Write "&nbsp;"
  
      For i = intFirstPage To BrowseGetPage - 1
        BrowseAHRef strAHRef, strEncript, i, i
  
      Next 
  
      Response.Write "<FONT COLOR=Red><B>" & BrowseGetPage & "</B></FONT>"
    
      For i = BrowseGetPage + 1 To intLastPage
        BrowseAHRef strAHRef, strEncript, i, i
  
      Next 
  
      Response.Write "&nbsp;"
      
      If intLastPage < sobjRS.PageCount Then
        BrowseAHRef strAHRef, strEncript, BrowseGetPage + intMorePages, "[Mais " & intMorePages & "]"
  
        Else
          Response.Write "&nbsp;[Mais 10]"
        
      End If
  
      Response.Write "&nbsp;"
    
      If BrowseGetPage > 1 Then
        BrowseAHRef strAHRef, strEncript, 1, "[Primeira]"
  
      Else
        Response.Write "[Primeira]"
       
      End If
    
      If BrowseGetPage < sobjRS.PageCount Then
        BrowseAHRef strAHRef, strEncript, sobjRS.PageCount, "[Última]"
  
      Else    
        Response.Write "[Última]"
        
      End If
    
    End If
  
    Response.Write "</FONT>"
    
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

    Response.Write "<BR>"

    BrowseAHRef strAHRef, strEncript, 1, "[Primeira Página]"
  
  End If
End Sub

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
REM Define quantidade de paginas na barra 
REM -------------------------------------------------------------------------
Public Sub BrowseSetBarPages(ByVal intBarPages)
 
  Session("BrowseBarPages") = intBarPages
 
End Sub

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
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Function BrowseGetActual

  BrowseGetActual = Session("BrowseActual")
  
End Function

REM -------------------------------------------------------------------------
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Sub BrowseSetLast

  BrowseSetActual Session("BrowseLast")
  
End Sub

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
REM Retorna se o Browse esta ativo
REM -------------------------------------------------------------------------
Public Sub BrowseSetStatus(ByVal blnStatus)

  Session("BrowseActive" & Session("BrowseActual")) = blnStatus
  
End Sub

REM -------------------------------------------------------------------------
REM Define o tamanho da pagina
REM -------------------------------------------------------------------------
Public Sub BrowseSetPageSize(ByVal intPageSize)

  Session("BrowsePageSize" & Session("BrowseActual")) = intPageSize
  
End Sub

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
REM Verifica se terminou a apresentacao da pagina
REM -------------------------------------------------------------------------
Public Function IsBrowsePage

  IsBrowsePage = True

  If BrowseGetLineCounter >= BrowseGetPageSize Then
    IsBrowsePage = False

  End If    
  
End Function

REM -------------------------------------------------------------------------
REM Conta as linhas de apresentacao
REM -------------------------------------------------------------------------
Public Sub BrowseLineCounter

  sintBrowseLineCounter = sintBrowseLineCounter + 1
  
End Sub

REM -------------------------------------------------------------------------
REM Conta as linhas de apresentacao
REM -------------------------------------------------------------------------
Public Function BrowseGetLineCounter

  BrowseGetLineCounter = sintBrowseLineCounter
  
End Function

REM -------------------------------------------------------------------------
REM Retorna o RS
REM -------------------------------------------------------------------------
Public Sub BrowseGetRS

  Set sobjRS = Session("BrowseRS" & Session("BrowseActual"))
  
End Sub

REM -------------------------------------------------------------------------
REM Define o RS
REM -------------------------------------------------------------------------
Public Sub BrowseSetRS

  Set Session("BrowseRS" & Session("BrowseActual")) = sobjRS

  BrowseSetStatus True
  
End Sub

REM -------------------------------------------------------------------------
REM Apaga todas as referencias do Browse demarcado
REM -------------------------------------------------------------------------
Public Sub BrowseClearActual
               
  BrowseSetStatus False
  BrowseSetActual ""
  
End Sub

REM -------------------------------------------------------------------------
REM Mostra o Cabecalho
REM -------------------------------------------------------------------------
Public Sub BrowseHeader

End Sub

REM -------------------------------------------------------------------------
REM Mostra o Rodape
REM -------------------------------------------------------------------------
Public Sub BrowseEnd

End Sub
