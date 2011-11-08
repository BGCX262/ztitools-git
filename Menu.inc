<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Edit.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Edicao
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

REM -------------------------------------------------------------------------
REM Inicio do menu
REM -------------------------------------------------------------------------
Public Sub MenuBegin

  If sblnVDHTML Then
    ShowHTMLCR "<script language=""JavaScript"">"
    ShowHTMLCR "<" + "!--"
    
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MenuBegin

REM -------------------------------------------------------------------------
REM Inicio do menu
REM -------------------------------------------------------------------------
Public Sub MenuEnd

  If sblnVDHTML Then
    ShowHTMLCR "//--" + ">"
    ShowHTMLCR "</script>"
    
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MenuEnd

REM -------------------------------------------------------------------------
REM Inicializacao dos botoes
REM -------------------------------------------------------------------------
Public Sub ButtonInit(ByVal strIMGExt, ByVal strIMGUp, ByVal strIMGOver, ByVal strIMGOut, ByVal strIMGDown)
  Dim strResult

  Session("MenuIMGExt")  = strIMGExt
  Session("MenuIMGUp")   = strIMGUp
  
  If sblnVDHTML Then
    strResult = ""
  
    If strIMGExt > "" Then strResult = """" & strIMGExt & """"
    If strIMGUp > "" Then strResult = strResult & ", """ & strIMGUp & """"
    If strIMGOver > "" Then strResult = strResult & ", """ & strIMGOver & """"
    If strIMGOut > "" Then strResult = strResult & ", """ & strIMGOut & """"
    If strIMGDown > "" Then strResult = strResult & ", """ & strIMGDown & """"

    ShowHTMLCR "InitButton(" & strResult & ");"
    
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ButtonInit

Dim sstrIMGSrc
Dim sstrIMGAlt
Dim sstrAHREF
Dim sstrATarget
Dim sintIMGWidth 
Dim sintIMGHeight
Dim sstrIMGRoot

REM -------------------------------------------------------------------------
REM Inicializacao dos botoes
REM -------------------------------------------------------------------------
Sub ButtonAdd(ByVal strIMGSrc, ByVal strIMGAlt, ByVal strAHREF, ByVal strATarget, ByVal intIMGWidth, ByVal intIMGHeight, ByVal strIMGRoot)
  Dim strResult
  
  strResult = ""

  If sblnVDHTML Then
    If strIMGSrc > "" Then strResult = """" & strIMGSrc & """"
    If strIMGAlt > "" Then strResult = strResult & ", """ & strIMGAlt & """"
    If strAHREF > "" Then strResult = strResult & ", """ & strAHREF & """"
    If strATarget > "" Then
      If LCase(strATarget) = "null" Then
        strResult = strResult & ", ""null"""
        
      Else
        strResult = strResult & ", """ & strATarget & """"
        
      End If      
    End If

    If intIMGWidth > "" Then strResult = strResult & ", "  & intIMGWidth 
    If intIMGHeight > "" Then strResult = strResult & ", "  & intIMGHeight
    If strIMGRoot > "" Then strResult = strResult & ", """ & strIMGRoot & """"

    ShowHTML "AddNewButton(" & strResult & ");"

  Else
    If strIMGSrc > "" Then
      sstrIMGSrc = strIMGSrc
      
    Else
      strIMGSrc = sstrIMGSrc

    End If
      
    If strIMGAlt > "" Then
      sstrIMGAlt = strIMGAlt
      
    Else
      strIMGAlt = sstrIMGAlt
      
    End If

    If strAHREF > "" Then
      sstrAHREF = strAHREF
      
    Else
      strAHREF = sstrAHREF
      
    End If

    If strATarget > "" Then
      sstrATarget = strATarget

    Else
      strATarget = sstrATarget
      
    End If

    If intIMGWidth > "" Then
      sintIMGWidth = intIMGWidth
      
    Else
      intIMGWidth = sintIMGWidth
      
    End If

    If intIMGHeight > "" Then
      sintIMGHeight = intIMGHeight
      
    Else
      intIMGHeight = sintIMGHeight
      
    End If

    If strIMGRoot > "" Then
      sstrIMGRoot = strIMGRoot
      
    Else
      strIMGRoot = sstrIMGRoot
      
    End If

    If strIMGSrc > "" Then strIMGSrc = " src=""" & strIMGRoot & strIMGSrc & Session("MenuIMGUp") & Session("MenuIMGExt") & """"
    If strIMGAlt > "" Then strIMGAlt = " alt=""" & strIMGAlt & """"
    If strAHREF > "" Then strAHREF = " HREF=""" & strAHREF & """"
    If strATarget > "" Then strATarget = " TARGET=""" & strATarget & """"
    If intIMGWidth > "" Then intIMGWidth = " WIDTH="  & intIMGWidth 
    If intIMGHeight > "" Then intIMGHeight = " HEIGHT="  & intIMGHeight

    strResult = "<A" & strAHREF & strATarget & ">" 
    strResult = strResult & "<IMG" & strIMGSrc & intIMGWidth & intIMGHeight & " BORDER=0>"
    strResult = strResult & "</A>"

    ShowHTML strResult
    
  End If

  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ButtonAdd

REM -------------------------------------------------------------------------
REM Inicializacao dos botoes
REM -------------------------------------------------------------------------
Sub ButtonHTML(strText)

  If sblnVDHTML Then
    ShowHTML "HTMLButton("""& strText & """);"

  Else
    ShowHTML strText
  
  End If
  
End Sub
REM -------------------------------------------------------------------------
REM Final da Sub ButtonHTML

REM -------------------------------------------------------------------------
REM Fim do Edit.inc
%>
