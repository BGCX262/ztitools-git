<%@ LANGUAGE="VBSCRIPT" %>
<%
REM -------------------------------------------------------------------------
REM  Code.asp
REM -------------------------------------------------------------------------
REM  Descricao   : Visualizador do Codigo Fonte de Programas ASP
REM  Cria‡ao     : 12:00h 20/01/97
REM  Local       : Brasilia/DF
REM  Elaborado   : Microsoft Coporation
REM -------------------------------------------------------------------------
%>
<!--#INCLUDE VIRTUAL="/ZTITools/All.inc"--> 
<!--#INCLUDE VIRTUAL="/ZTITools/Zevallos.inc"--> 
<%

MainBody

REM -------------------------------------------------------------------------
REM -------------------------------------------------------------------------
FUNCTION fValidPath (ByVal strPath)
  Dim iPos
  
  strPath = LCase(strPath)
  
  iPos = InStr(1, strPath, "/", 1)
  
  If iPos Then
    fValidPath = 1
  Else
    fValidPath = 0
  End If

  iPos = InStr(2, strPath, "/", 1)

  If (lcase(Right(strPath, 4)) = ".asa" Or Left(strPath, iPos) = "/ztitools/") And strPath <> "/ztitools/code/code.asp" Then
    fValidPath = 0
    
  End If
 
END FUNCTION

REM -------------------------------------------------------------------------
REM Returns the minimum number greater than 0
REM If both are 0, returns -1
REM -------------------------------------------------------------------------
FUNCTION fMin (iNum1, iNum2)
 If iNum1 = 0 AND iNum2 = 0 Then
   fMin = -1
 ElseIf iNum2 = 0 Then
   fMin = iNum1
 ElseIf iNum1 = 0 Then
   fMin = iNum2
 ElseIf iNum1 < iNum2 Then
   fMin = iNum1
 Else 
   fMin = iNum2
 End If
END FUNCTION
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
FUNCTION fCheckLine (ByVal strLine)
 Dim iTemp, iPos

 fCheckLine = 0
 iTemp = 0

 iPos = InStr(strLine, "<" & "%")
 If fMin(iTemp, iPos) = iPos Then
   iTemp = iPos
   fCheckLine = 1
 End If
 
 iPos = InStr(strLine, "%" & ">")
 If fMin(iTemp, iPos) = iPos Then
   iTemp = iPos
   fCheckLine = 2
 End If

 iPos = InStr(1, strLine, "<" & "SCRIPT>", 1)
 
 If fMin(iTemp, iPos) = iPos Then
   iTemp = iPos
   fCheckLine = 3
 End If

 iPos = InStr(1, strLine, "<" & "/SCRIPT>", 1)
 
 If fMin(iTemp, iPos) = iPos Then
   iTemp = iPos
   fCheckLine = 4
 End If

END FUNCTION
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
SUB PrintHTML (ByVal strLine)
 Dim iSpaces, i, iPos
 
iSpaces = Len(strLine) - Len(LTrim(strLine))
i = 1
While Mid(Strline, i, 1) = Chr(9)
	iSpaces = iSpaces + 5
	i = i + 1
Wend

 If iSpaces > 0 Then
   For i = 1 to iSpaces
     Response.Write("&nbsp;")
   Next
 End If

 iPos = InStr(strLine, "<")

 If iPos Then
   Response.Write(Left(strLine, iPos - 1))
   Response.Write("&lt;")
   strLine = Right(strLine, Len(strLine) - iPos)
   PrintHTML(strLine)
   
 Else
   Response.Write(strLine)
   
 End If
END SUB
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------
SUB PrintLine (ByVal strLine, iFlag)
 Dim iPos
 
 Select Case iFlag
   Case 0
     PrintHTML(strLine)
     
   Case 1
     iPos = InStr(strLine, "<" & "%")
     PrintHTML(Left(strLine, iPos - 1))
     Response.Write("<FONT COLOR=#ff0000>")
     Response.Write("&lt;%")
     strLine = Right(strLine, Len(strLine) - (iPos + 1))
     PrintLine strLine, fCheckLine(strLine)
   Case 2
     iPos = InStr(strLine, "%" & ">")
     PrintHTML(Left(strLine, iPos -1))
     Response.Write("%&gt;")
     Response.Write("</FONT>")
     strLine = Right(strLine, Len(strLine) - (iPos + 1))
     PrintLine strLine, fCheckLine(strLine)
   Case 3
     iPos = InStr(1, strLine, "<" & "SCRIPT", 1)
     PrintHTML(Left(strLine, iPos - 1))
     Response.Write("<FONT COLOR=#0000ff>")
     Response.Write("&lt;SCRIPT")
     strLine = Right(strLine, Len(strLine) - (iPos + 6))
     PrintLine strLine, fCheckLine(strLine)
   Case 4
     iPos = InStr(1, strLine, "<" & "/SCRIPT>", 1)
     PrintHTML Left(strLine, iPos - 1)
     Response.Write("&lt;/SCRIPT&gt;")
     Response.Write("</FONT>")
     strLine = Right(strLine, Len(strLine) - (iPos + 8))
     PrintLine strLine, fCheckLine(strLine)
     
   Case Else
     Response.Write("FUNCTION ERROR -- PLEASE CONTACT ADMIN.")
     
 End Select
END SUB
REM -------------------------------------------------------------------------

REM -------------------------------------------------------------------------

Sub MainBody
Dim strVirtualPath
Dim strFilename
Dim objFS, oInStream, strOutput

  HTMLBegin
  HeadAll "Visualiza o Codigo do Programa ASP - Active Server Page"
  
  BodyBegin
  
  ShowPageTitle "/ztitools/img/Assinatura75.Gif", "Logo Zevallos", "Tecnologia em Informa&ccedil;&atilde;o"
  
  ShowNewsTitle "<p><center>Visualiza o Código Fonte de Programa ASP</center>"
  
  strVirtualPath = Request("sourcefile")
  
  
  ShowHTMLCR "<BR>Retorna para <a href=" & strVirtualPath & ">" & strVirtualPath & "</A>"
  ShowHTMLCR "<BR>Retorna para <a href=/ TARGET=_top>Pagina Principal do Site</A><hr>"
  
  ShowHTMLCR "<PRE>"
  
  If fValidPath(strVirtualPath) Then
  	strFilename = Server.MapPath(strVirtualPath)
  
  	Set objFS = Server.CreateObject("Scripting.FileSystemObject")
  
  	Set oInStream = objFS.OpenTextFile(strFilename, 1, FALSE)
  
      While NOT oInStream.AtEndOfStream
        strOutput = oInStream.ReadLine
        PrintLine strOutput, fCheckLine(strOutput)
        Response.Write("<BR>")
      Wend
  
    Else
      Response.Write("<H1>Visuação do Código Fonte -- Acesso Negado!</H1>")
  
    End If  
  
  ShowHTMLCR "</PRE>"
  
  PageFooterDefault
  
  BodyEnd
  HTMLEnd
End Sub
%>