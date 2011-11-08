<%
	'-----------------------------------------------------
	'-- This file contains various ASP helper functions --
	'-----------------------------------------------------

	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: check the filename and
	':: Return an absolute path
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function CheckFileName(sFileName)
		if mid(sFileName, 1, 2)<>"\\" and mid(sFileName, 2, 1)<>":" then 
			CheckFileName = g_sServerMapPath & iff(mid(sFileName, 1, 1)="\", "", "\") & sFileName
		else
			CheckFileName = sFileName
		end if
	End Function

	'------------------------------------------
	'-- Return the extension from a filename --
	'------------------------------------------
	Function RemoveExtension(sFilename)
		RemoveExtension = Mid(sFilename, 1, InStrRev(sFilename, ".")-1)
	End Function
	
	
	Function RegistryKey(sKey)
		Dim oShell
		Set oShell = server.CreateObject("WScript.Shell") 
			RegistryKey = oShell.RegRead(sKey)
		Set oShell = nothing
	End Function
	
	
	'-- Repeat a string ----------------------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Function repeat(num, str)
		Dim i
		For i = 1 to num
			repeat = repeat & str
		Next
	End Function


	'-- Open a text file and return its content, use the global FSO object
	'-----------------------------------------------------------------------------------------------------------------------
	Function LoadFile(filename)
		if mid(filename, 1,2)<>"\\" and mid(filename, 2, 1)<>":" then filename = g_sServerMapPath & "\" & filename
		Dim tmp : tmp = Application(filename)		
		
		If Len(tmp)=0 oR Request("reloadfiles")="1" or true then
			'debug "File loaded from  Filesystem: " & filename
			if g_oFSO.FileExists(filename) then
				Dim oFile
				Set oFile = g_oFSO.OpenTextFile(filename)
				LoadFile = oFile.ReadAll
				oFile.Close
				Set oFile = Nothing
				Application(filename)	= LoadFile
			else
				LogIt "Global.asp", "LoadFile", ERROR, "File does not exists", filename
			end if
		Else
			'debug "File loaded from cache: " & filename
			LoadFile = tmp
		End If
	End Function
	
	
	'-- delete a file, by renaming it ----------------------------------------------------
	' Inout: 
	'		filname: the file to delete
	'-------------------------------------------------------------------------------------
	Sub DeleteFile(filename)
		If  g_oFSO.FileExists(filename) then
			g_oFSO.MoveFile filename ,  filename & ".deleted"
		End If
	end Sub
	
	
	'-- Check if a folder exists ---------------------------------------------------------
	' Input: 
	'		folder: the file to delete
	'-------------------------------------------------------------------------------------
	Function CheckFolder(folder)
		If  g_oFSO.FolderExists(folder) then
			CheckFolder = true
		Else
			CheckFolder = false
		End If
	end Function
	
	
	'-- Create a folder ----------------------------------------------------------------------------------------------------
	' Input:
	'			sFullname: the path to create
	' Output:
	'			True in case of success
	'-----------------------------------------------------------------------------------------------------------------------
	Private Function CreateFolder(sFullName)
		
		' good char and  remove last slash
		sFullName = replace(sFullName, "/", "\")
		if right(sFullName, 1) = "\" then sFullName = left(sFullName, len(sFullName)-1)
		
		
		if NOT g_oFSO.FolderExists(sFullName) then
			on error resume next
			g_oFSO.CreateFolder(sFullName)
			If Err.Number<>0 then
				LogIT "utilities.asp", "CreateFolder", ERROR, Err.Description, sFullName
				Err.Clear
				on error goto 0	
				CreateFolder = false
			else
				CreateFolder = true			
			End if
			on error goto 0
		Else
			CreateFolder = true	
		End if
		
	End Function
	
	
	
	
	'-- Delete a folder ----------------------------------------------------------------------------------------------------
	' Input:
	'			sFullname: the path to delete
	' Output:
	'			True in case of success
	'-----------------------------------------------------------------------------------------------------------------------
	Private Function DeleteFolder(sFullName)
		
		' good char.
		sFullName = replace(sFullName, "/", "\")
		
		' remove last slash
		if right(sFullName, 1) = "\" then sFullName = left(sFullName,  len(sFullName)-1)
		
		if g_oFSO.FolderExists(sFullName) then
			on error resume next
			g_oFSO.DeleteFolder(sFullName)
			If Err.Number<>0 then
				LogIT "utilities.asp", "DeleteFolder", ERROR, Err.Description, sFullName
				Err.Clear
				on error goto 0	
				DeleteFolder = false
			else
				DeleteFolder = true			
			End if
			on error goto 0
		Else
			DeleteFolder = true	
		End if
		
	End Function
	
	
	Function TemplateFileContent(sFileName)
		TemplateFileContent = LoadFile(sFileName)		
	End Function


	Sub Die(text)
		Response.Write text
		Response.End
	End Sub
	
	
	'-- return a Request Filtered element that is usable in a sql query
	Private Function GetParam(sName)
		GetParam = iff(len(Request.form(sName))>0, Request.form(sName), Request.QueryString(sName)) 'replace(Request.Form(sName), "'", "''")
	End Function
		
	
	' The IFF Function is a helper one: if test is True, t is returend, if test is False, f is returned.
	' This function is the ASP implementation of the ? : C(++) operator.
	Function IFF(test, t, f)
		If test Then
			IFF = t
		Else
			IFF = f
		End If
	End Function
	
	
	'-- return a date formated in YYYYMMDDHHNN
	Function YYYYMMDDHHNN(mydate)
		YYYYMMDDHHNN = cstr(Year(mydate) & Right("0" & Month(mydate), 2) & Right("0" & Day(mydate), 2) & Right("0" & Hour(mydate), 2) & Right("0" & Minute(mydate), 2))
	End Function
	
		
	Private function GP(oNode, sParam)
		If oNode.selectNodes("add[@key='" & sParam & "']/@value").length>0 Then
			GP = oNode.selectSingleNode("add[@key='" & sParam & "']/@value").value
			if len(GP) = 0 Then
				GP =  oNode.selectSingleNode("add[@key='" & sParam & "']").text
			End If
		Else
			GP = ""
		End If
	End Function
		
	
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	' Return a random GUID string
	'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
	Private Function GetGuid()
		
		on error resume next
		GetGuid = CreateGUID()
		if err=0 then
			on error goto 0
			exit function
		End If
		
		err.Clear
		on error goto 0
		
		
		Dim intLen : intLen = 12
		Dim strTemp 
		Dim strChar
		Dim i
		i = 0
		randomize
		do While i < intLen  
			strChar = int(rnd * 74) + 48
			if 	(strChar >= 48 and strChar <= 57) or _
				(strChar >= 65 and strChar <= 90) or _
				(strChar >= 97 and strChar <= 122) Then
				i = i + 1
				strTemp = strTemp + chr(strChar)
			End if
		Loop
		GetGuid = strTemp
	End Function
		
	
	'--------------------------
	'-- Create a 'true' GUID --
	'--------------------------
	Function CreateGUID()
		 Dim oTypeLib, sGUID
		  
		Set oTypeLib = Server.CreateObject("Scriptlet.Typelib")
			sGUID = oTypeLib.GUID
			sGUID = LCase(sGUID)
			sGUID = Mid(sGUID, 1, Len(sGUID) - 2)
		Set oTypeLib = Nothing

		CreateGUID = Trim(sGUID)
	End Function  
	
	
	'+-----------------------------------------------------------------+
	'| The Include Function
	'+-----------------------------------------------------------------+
	Function Include(vbsFile)
		Dim fso, ts, buf
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set ts = fso.OpenTextFile(vbsFile)
		buf = ts.ReadAll()
		ts.Close

		ExecuteGlobal buf
	End Function
	
	
	'---------------------------------------------------------------
	'-- This function is used to wrap a function call, with cache --
	'---------------------------------------------------------------
	Function CacheFunctionCall(p_function, cachename, cachetimeout)
		Dim reload : reload = true
		dim tmp : tmp = Application(cachename)
		dim cachedate : cachedate = iff(len(tmp)>=12, mid(tmp, 1, 12), "")
				
		If len(tmp)>0 and len(cachename)>0 and len(cachetimeout)>0 and len(cachedate)=12 then
			if int(YYYYMMDDHHNN(now)-cachedate) <= int(cachetimeout) then reload = false
		End if
		
	'	reload = true
					
		'-- Si reload est a true, on execute la function
		if reload then
			tmp = cstr(YYYYMMDDHHNN(now))
			Call execute("tmp = tmp & " & p_function)
			Application(cachename) = tmp
		end if
		
		'return data
		CacheFunctionCall = mid(tmp, 13)
	End function
	
	
	'-- execute un gethttp, avec gestion de caache dans des var. applications
	Function GetHttp(url, cachename, cachetimeout)
		Dim reload : reload = true
		dim tmp : tmp = Application(cachename)
		dim cachedate : cachedate = mid(tmp, 1, 12)
				
		If len(cachename)>0 and len(cachetimeout)>0 and len(cachedate)>0 then
			if int(YYYYMMDDHHNN(now)-cachedate) <= int(cachetimeout) then reload = false
		End if
						
		'-- on fait le get		
		if reload then
			Dim oWht
			Set oWht = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
			oWht.open "GET",  url	
			oWht.send
			
			if oWht.Status=200 then
				tmp = cstr(YYYYMMDDHHNN(now)) & oWht.ResponseText
				Application(cachename) = tmp
			else
				die "GetHttp::" & oWht.Status
			end if
			
			set oWht = Nothing
		end if
		
		'return data
		GetHttp = mid(tmp, 13)
	End function
	
	
	'-- execute un gethttp, avec gestion de caache dans des var. applications
	Function XmlXsl(xml, xsl, cachename, cachetimeout)
		Dim reload : reload = true
		dim tmp : tmp = Application(cachename)
		dim cachedate : cachedate = mid(tmp, 1, 12)
					
		If len(cachename)>0 and len(cachetimeout)>0 and len(cachedate)>0 then
			if int(YYYYMMDDHHNN(now)-cachedate) <= int(cachetimeout) then reload = false
		End if
						
		'-- Do the transformation
		if reload then
			tmp = cstr(YYYYMMDDHHNN(now)) & Transform(xml, xsl)
			Application(cachename) = tmp
		end if
					
		'return data
		XmlXsl = mid(tmp, 13)
	End function
	
	
	'---------------
	'-- Log error --
	'---------------
	Sub LogIt(filename, method, level, title, message)		
		'on error resume next
		
		'-- log the error into a file
		Dim oLog
		Set oLog = New LogFile
			oLog.TemplateFileName = DATA_FOLDER & LOGS_FOLDER & "%y-%m-%d.csv"
			oLog.FieldSeparator = GetSeparator
			oLog.Log(array(filename, method, level, title, message, g_sUrl, Request.Form))
		Set oLog = nothing
		
		'-- clear and remove the error trapping
		if err<>0 then
			err.Clear
		end if
		'on error goto 0		
	End Sub
	
	'---------------------------------------------------------------------------------------------
	'-- This function return the separator used by the Text Driver (for stats usage)
	'-- the 1st time, It read the value from the registry and store it in a application variable
	'---------------------------------------------------------------------------------------------
	Function GetSeparator
		
		If len(application(APPVAR_SEPARATOR))=0 then
			Dim tmp : tmp = trim(RegistryKey(appSettings("JET_REGISTRY_KEY") & "\Format"))	
			
			Select Case tmp
				case "CSVDelimited"
					GetSeparator = ","
				case else
					GetSeparator = TRIM(mid(tmp, InStr(1,tmp, "(") +1, InStr(1,tmp, ")")-InStr(1,tmp, "(")-1))
			End Select
			
			application(APPVAR_SEPARATOR) = GetSeparator
		Else
			GetSeparator = application(APPVAR_SEPARATOR)
		End if		
	End Function
	
	
	'--------------------------------------------------------
	'-- Print a fatal error message and stop the execution --
	'--------------------------------------------------------
	Sub FatalError (title, message)
		with Response
			.Write "<html>"
			.Write "<head><title>"&title&"</title></head>"
			.Write "<body style='font: messagebox;'><img src=engine/admin/media/error.png align=left><span style='color: 6584C0; font: 12pt verdana; display: inline; font-weight: bold;width: 400px;'>"&title&"</span><br><br>" & message & "</body>"
			.Write "</html>"
		end with
		Response.End
	End Sub
	
		
	'--------------------------------------------------
	'-- Reload the opener window and close the popup --
	'--------------------------------------------------
	Sub ReloadMainFrameAndClose
		Response.Write "<scr"&"ipt language='javascript'>window.opener.location.reload(true); self.close();</sc"&"ript>"
	End Sub
	
	
	'------------------------------------------------------------------
	'-- Ajust the size of the popup to the size of the table tblEdit --
	'------------------------------------------------------------------
	Sub ResizePopup
		With Response
			.Write "<script language='javascript'>" & vbCrLf
			.Write "	//ajust the size of the popup to the size of the table tblEdit" & vbCrLf
			.Write "	window.resizeTo(document.all.tblEdit.clientWidth + 10, document.all.tblEdit.clientHeight + 30); " & vbCrLf
				
			.Write "	//center the popup" & vbCrLf
			.Write "	var w = screen.availWidth;" & vbCrLf
			.Write "	var h = screen.availHeight;" & vbCrLf
			.Write "	var popW = document.all.tblEdit.clientWidth + 8;" & vbCrLf
			.Write "	var popH = document.all.tblEdit.clientHeight + 20;" & vbCrLf
			.Write "	var leftPos = (w-popW)/2;" & vbCrLf
			.Write "	var topPos = (h-popH)/2;" & vbCrLf
				
			.Write "	window.moveTo(leftPos, topPos);" & vbCrLf
			.Write "</script>"
		End With
	End Sub
	
	
	'-------------------------------------------------------------
	'-- Determine if a Value Exists in an Array without Looping --
	'-------------------------------------------------------------
	Public Function IsInArray(FindValue, arrSearch )
		If Not IsArray(arrSearch) Then Exit Function		
		IsInArray = InStr(1, vbNullChar & Join(arrSearch, vbNullChar) & vbNullChar, vbNullChar & FindValue & vbNullChar) > 0
	End Function

%>