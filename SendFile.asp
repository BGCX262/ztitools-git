<%@ LANGUAGE="VBSCRIPT" %>
<%
REM -------------------------------------------------------------------------
REM  /Pedidos.ASP
REM -------------------------------------------------------------------------
REM  Descricao   : Manipula os pedidos de compra
REM  Criacao     : 11:00h 14/03/99
REM  Local       : Brasilia/DF
REM  Elaborado   : Eduardo Alves <edualves@zevallos.com.br>
REM  Versao      : 1.0.0
REM  Copyright   : 1999 by Zevallos(r) Tecnologia em Informacao
REM -------------------------------------------------------------------------

%>
<!--#INCLUDE VIRTUAL="/ZTITools/All.inc"-->
<!--#INCLUDE VIRTUAL="/ZTITools/EditForm/Current.inc"-->
<%

MainBody

Private Sub Escolhe
  ShowHTMLCR "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
  ShowHTMLCR "<TITLE>Envio de imagem</TITLE>"
  ShowHTMLCR "</HEAD>"
  ShowHTMLCR "<BODY BGCOLOR=#ffffff>"
  ShowHTMLCR "<H2><CENTER>Envio de imagem</CENTER></H2>"
  
  ShowHTML "<FORM ENCTYPE=""MULTIPART/FORM-DATA"" METHOD=""POST"" ACTION="""
  ShowHTML "SendFile.asp?O=2&"
  ShowHTMLCR "Folder=" & Server.URLEncode(sstrThisSiteRootDir & "img\produtos\") & "&Field=" & Request.QueryString("Field")& "&FileName=" & Request.QueryString("FileName") & """>"
  ShowHTMLCR "Escolha o arquivo:"
  ShowHTMLCR "<INPUT TYPE=""FILE"" NAME=""FILE1"">"
  ShowHTMLCR "<INPUT TYPE=""SUBMIT"" NAME=""SUB1"" VALUE=""   OK   "">"
  ShowHTMLCR "</FORM>"
REM  ShowHTMLCR "Folder=" & Request.QueryString("Folder") & "&Field=" & Request.QueryString("Field")& "&FileName=" & Request.QueryString("FileName") & """>"
  ShowHTMLCR "</BODY>"

End Sub

Private Sub Manda
Dim strExten, intPos, upl

  Set upl = Server.CreateObject("SoftArtisans.FileUp")
  upl.Path = Request.QueryString("Folder")

  If upl.IsEmpty Then
    ShowMessageError "O arquivo escolhido está vazio."
    Escolhe

  ElseIf upl.ContentDisposition <> "form-data" Then
    ShowMessageError "O arquivo não pode ser enviado"
    Escolhe

  Else
  	on error resume next 
  	intPos = InStr(ZTIReverse(upl.UserFilename), ".")
  	If intPos > 0 Then
  	  strExten = Right(upl.UserFilename, intPos)

  	End If  
  	upl.SaveAs Request.QueryString("FileName") & strExten
  
  	If Err <> 0 Then
      ShowMessageError "Um erro ocorreu ao tentar salvar o arquivo"
      Escolhe

    Else
      If EditFindField(Request.QueryString("Field")) Then
        Session(EditCurrentField & "value") = Request.QueryString("FileName") & strExten

      End If
      ShowHTMLCR "<SCRIPT>"
      ShowHTMLCR "  parent.window.close();"
      ShowHTMLCR "</SCRIPT>"

    End If
  
  End If  

End Sub

REM -------------------------------------------------------------------------
REM Corpo Principal do sistema
REM -------------------------------------------------------------------------
Private Sub MainBody

  ShowHTMLCR "<HTML>"
  Select Case Request.QueryString("O")
    Case "1"
      Escolhe

    Case "2"
      Manda

  End Select
  ShowHTMLCR "</HTML>"

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub MainBody

REM -------------------------------------------------------------------------
REM Fim do ZTIEditForm.asp
%>