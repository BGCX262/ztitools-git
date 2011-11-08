<%
REM -------------------------------------------------------------------------
REM  /ZTITools/File.inc
REM -------------------------------------------------------------------------
REM  Descricao   : Include de File
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
REM Verifica se o arquivo existe
REM -------------------------------------------------------------------------
Public Function IsFileExist(ByVal strFullPath)
  Dim objFS

  SET objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If
  
  IsFileExist = objFS.FileExists(strFullPath)


  SET objFS = Nothing
  
End Function
REM -------------------------------------------------------------------------
REM Final da Function isFileExist

REM -------------------------------------------------------------------------
REM Copia arquivos
REM -------------------------------------------------------------------------
Function FileCopy(ByVal strSource, ByVal strDestination, ByVal blnOverwrite)
  Dim objFS

  SET objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If
  
  IsFileExist = objFS.CopyFile(strSource, strDestination, blnOverwrite)


  SET objFSFE = Nothing
  
  
End Function
REM -------------------------------------------------------------------------
REM Final da FileCopy

REM -------------------------------------------------------------------------
REM Copia arquivos e sobrescreve
REM -------------------------------------------------------------------------
Function FileCopyOverwrite(ByVal strSource, ByVal strDestination)

  FielCopyOverwrite = FileCopy(strSource, strDestination, True)
  
End Function
REM -------------------------------------------------------------------------
REM Final da FileCopyOverwrite

REM -------------------------------------------------------------------------
REM Extrai o nome do arquivo do Path
REM -------------------------------------------------------------------------
Function ExtractFileName(sFileName)
	Dim nIdx
	On Error Resume Next
	For nIdx = Len(sFileName) To 1 Step -1
	If Mid(sFileName, nIdx, 1) = "\" Then
		ExtractFileName = Mid(sFileName, nIdx + 1)
		Exit Function
	End If
	Next
	ExtractFileName = sFileName
End Function
REM -------------------------------------------------------------------------
REM Final da FileCopyOverwrite

REM -------------------------------------------------------------------------
REM Retorna a data do arquivo
REM -------------------------------------------------------------------------
Function FileDateCreated(ByVal strFullPath)
  Dim objFS, objFile

  SET objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If

  SET objFile = objFS.GetFile(strFullPath)
 
  FileDateCreated = objFile.DateCreated

  SET objFile = Nothing

  SET objFS = Nothing

End Function
REM -------------------------------------------------------------------------
REM Final da Function FileDateCreated

REM -------------------------------------------------------------------------
REM Retorna a data de atualizacao do arquivo
REM -------------------------------------------------------------------------
Function FileDateLastModified(ByVal strFullPath)
  Dim objFS, objFile

  SET objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If

  SET objFile = objFS.GetFile(strFullPath)
 
  FileDateLastModified = objFile.DateLastModified

  SET objFile = Nothing

  SET objFS = Nothing

End Function
REM -------------------------------------------------------------------------
REM Final da Function FileDateCreated

REM -------------------------------------------------------------------------
REM Retorna a data do arquivo
REM -------------------------------------------------------------------------
Function FileSize(ByVal strFullPath)
  Dim objFS, objFile

  SET objFS = Server.CreateObject("Scripting.FileSystemObject")

  If objFS.GetDriveName(strFullPath) = "" Then
    strFullPath = Server.MapPath(strFullPath)
  
  End If

  SET objFile = objFS.GetFile(strFullPath)
 
  FileSize = objFile.Size

  SET objFile = Nothing

  SET objFS = Nothing

End Function
REM -------------------------------------------------------------------------
REM Final da Function FileSize

REM -------------------------------------------------------------------------
REM Fim do File.inc
%>