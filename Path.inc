<%
REM -------------------------------------------------------------------------
REM  /ZTITools/Path.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de Path
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

REM -------------------------------------------------------------------------
REM Retorna a barra do path 
REM -------------------------------------------------------------------------
Public Function PathGetSlash(ByVal strFullPath)
  Dim strResult
  
    strResult  = ""
  
    If InStr(strFullPath, "\") Then
      strResult = "\"
    
    ElseIf InStr(strFullPath, "/") Then
      strResult = "/"
    
    End If
    
  PathGetSlash = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PathGetSlash

REM -------------------------------------------------------------------------
REM Estrai o nome do arquivo de um Full Path
REM -------------------------------------------------------------------------
Public Function PathInvertSlash(ByVal strFullPath)
  Dim strResult, strSlash, strInvSlash
  Dim intSlashPos, intLen

  strResult = ""
  
  If strFullPath > "" Then  
    intSlashPos = 0
    strResult   = strFullPath
    intLen      = Len(strFullPath)
    strSlash    = PathGetSlash(strFullPath)
    
    If strSlash = "\" Then
      strInvSlash = "/"
      
    Else
      strInvSlash = "\"
      
    End If
      
    Do
      intSlashPos = InStr(intSlashPos + 1, strResult, strSlash)

      If intSlashPos > 0 Then      
        strResult = Left(strResult, intSlashPos - 1) & strInvSlash & Right(strResult, intLen - intSlashPos)
        
      End If
  
    Loop While intSlashPos > 0

  End If      
  
  PathInvertSlash = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PathInvertSlash

REM -------------------------------------------------------------------------
REM Converte o path para barras /
REM -------------------------------------------------------------------------
Public Function PathURL(ByVal strFullPath)
  Dim strResult

  strResult = strFullPath

  If strFullPath > "" Then
    If PathGetSlash(strFullPath) = "\" Then
      strResult = PathInvertSlash(strFullPath)

      If PathDrive(strFullPath) > "" Then
        strResult = Mid(strResult, 3)
        
      End If
    End If
  End If
  
  PathURL = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PathURL

REM -------------------------------------------------------------------------
REM Converte o path para barras \
REM -------------------------------------------------------------------------
Public Function PathNormal(ByVal strFullPath)
  Dim strResult
  
  strResult = strFullPath

  If strFullPath > "" Then
    If PathGetSlash(strFullPath) = "/" Then
      strResult = PathInvertSlash(strFullPath)
      
    End If
  End If
  
  PathNormal = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PathNormal

REM -------------------------------------------------------------------------
REM Converte o path para barras \
REM -------------------------------------------------------------------------
Public Function PathDelRightSlash(ByVal strFullPath)
  Dim strResult

  strResult = strFullPath
    
  If strFullPath > "" Then
    If Right(strResult, 1) = PathGetSlash(strResult) Then
      strResult = Left(strResult, Len(strResult) -1 )
      
    End If
  End If
    
  PathDelRightSlash = strResult
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function PathToURL

REM -------------------------------------------------------------------------
REM Estrai o nome do arquivo de um Full Path
REM -------------------------------------------------------------------------
Public Function PathFileName(ByVal strFullPath)
  Dim strResult, strSlash
  
  strResult = ""

  If strFullPath > "" Then  
    If InStr(strFullPath, ".") > 0 Then
      strResult = strFullPath
  
      strSlash = PathGetSlash(strFullPath)
      
      If strSlash > "" Then
        strResult = Mid(strFullPath, InstrRev(strFullPath, strSlash) + 1)
        
      End If    
    End If
  End If
    
  PathFileName = strResult
    
End Function
REM -------------------------------------------------------------------------
REM Final da Function PathFileName

REM -------------------------------------------------------------------------
REM Estrai o Path de um Full Path
REM -------------------------------------------------------------------------
Public Function Path(ByVal strFullPath)
  Dim strResult

  strResult = ""
  
  If strFullPath > "" Then    
    strResult = strFullPath
    
    If InStr(strFullPath, ".") > 0 Then
      strResult = Left(strFullPath, Len(strFullPath) - Len(PathFileName(strFullPath)))
      
    End If
  End If
   
  Path = strResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function Path

REM -------------------------------------------------------------------------
REM Estrai o a letra do Drive de um Full Path
REM -------------------------------------------------------------------------
Public Function PathOnly(ByVal strFullPath)
  Dim strResult, intColon
  
  strResult = ""

  If strFullPath > "" Then  
    strFullPath = Path(strFullPath)
    
    intColon = InStr(strFullPath, ":")
    
    If intColon > 0 Then
      strResult = Mid(strFullPath, intColon + 1)
      
    End If
  End If
  
  PathOnly = strResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function PathOnly

REM -------------------------------------------------------------------------
REM Verifica se e uma URL
REM -------------------------------------------------------------------------
Public Function IsURL(ByVal strURL)

  strURL = UCase(Trim(strURL))

	IsURL = True

	If Left(strURL, 6) = "HTTP:/" Then
	ElseIf Left(strURL, 6) = "FILE:/" Then
	ElseIf Left(strURL, 8) = "MAILTO:/" Then
	ElseIf Left(strURL, 5) = "FTP:/" Then
	ElseIf Left(strURL, 8) = "GOPHER:/" Then
	ElseIf Left(strURL, 6) = "NEWS:/" Then
	ElseIf Left(strURL, 7) = "HTTPS:/" Then
	ElseIf Left(strURL, 8) = "TELNET:/" Then
	ElseIf Left(strURL, 6) = "NNTP:/" Then
	Else
		IsURL = False
		
	End IF
	
End Function
REM -------------------------------------------------------------------------
REM Final da Function IsURL

REM -------------------------------------------------------------------------
REM Estrai o a letra do Drive de um Full Path
REM -------------------------------------------------------------------------
Public Function PathDrive(ByVal strFullPath)
  Dim strResult
  
  strResult = ""

  If strFullPath > "" Then  
    If InStr(strFullPath, ":") > 0 Then
      strResult = Left(strFullPath, 2)
      
    End If
  End If
  
  PathDrive = strResult

End Function
REM -------------------------------------------------------------------------
REM Final da Function PathDrive

REM -------------------------------------------------------------------------
REM Demo do Path
REM -------------------------------------------------------------------------
Public Sub PathDemo
  Dim strFullPath
  
  strFullPath = Request.ServerVariables("PATH_TRANSLATED")

  Response.Write "<BR>" & strFullPath
  Response.Write "<BR>" & PathDrive(strFullPath)
  Response.Write "<BR>" & PathFileName(strFullPath)
  Response.Write "<BR>" & Path(strFullPath)
  Response.Write "<BR>" & PathOnly(strFullPath)
  Response.Write "<BR>" & PathGetSlash(strFullPath)
  Response.Write "<BR>" & PathInvertSlash(strFullPath)
  Response.Write "<BR>PathURL=" & PathURL(strFullPath)
  Response.Write "<BR>PathNormal=" & PathNormal(strFullPath)
  Response.Write "<BR>PathDelRightSlash=" & PathDelRightSlash(Path(strFullPath))

  strFullPath = Request.ServerVariables("PATH_INFO")

  Response.Write "<BR>FullPath=" & strFullPath
  Response.Write "<BR>PathDrive=" & PathDrive(strFullPath)
  Response.Write "<BR>PathFileName=" & PathFileName(strFullPath)
  Response.Write "<BR>Path=" & Path(strFullPath)
  Response.Write "<BR>PathOnly=" & PathOnly(strFullPath)
  Response.Write "<BR>PathGetSlash=" & PathGetSlash(strFullPath)
  Response.Write "<BR>PathInvertSlash=" & PathInvertSlash(strFullPath)
  Response.Write "<BR>PathURL=" & PathURL(strFullPath)
  Response.Write "<BR>PathNormal=" & PathNormal(strFullPath)
  Response.Write "<BR>PathDelRightSlash=" & PathDelRightSlash(Path(strFullPath))

End Sub
REM -------------------------------------------------------------------------
REM Final da Sub PathDemo

REM -------------------------------------------------------------------------
REM Fim do Path.inc
%>