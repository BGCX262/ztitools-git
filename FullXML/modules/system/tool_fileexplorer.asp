<!-- #include file="lib/freeASPUpload.asp" -->
<%
	
	'-- return the root folder for the filesystem explorer
	private function current_user_root_folder
		if g_oUser.Group<>"administrator" and g_oUser.Group<>"webmaster" Then
			current_user_root_folder = "/" & g_oUser.Login & "/" 
		else
			current_user_root_folder = "/"
		end if
	end function
	
	'-- return the current sub folder for the filesystem explorer
	private function current_sub_folder
		If instr(1, Request.QueryString("SubPath"), "..") then
				current_sub_folder = ""
			Else
				current_sub_folder = Request.QueryString("SubPath")
			End if
	end function
	
	
	'-- Display the FileExplorer -------------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Sub webform_System_FileExplorer
		Dim oUploadFolder, oSub, oFile		
				
		'-- Display the Explorer view
		Set oUploadFolder = g_oFSO.GetFolder(upload_path & current_user_root_folder & current_sub_folder)
		With Response
			
			.Write "<script language=javascript src=modules/system/fileexplorer.js></script>"
			
			'-- display the toolbar
			.Write "<table cellpadding=0 cellspacing=0 class=tb>"
			.Write "<tr><td>"
			if (g_oUser.Group="administrator" or g_oUser.Group="webmaster") then .Write "<a href="&g_sScriptName&"?webform=webform_system_fileexplorer_newfolder&subpath="& current_sub_folder & ">New folder</a>"
			if (g_oUser.Group="administrator" or g_oUser.Group="webmaster") then.Write "<a "& iff(len(current_sub_folder) >1, "", "disabled") & " href="&g_sScriptName&"?webform=webform_system_fileexplorer_deletefolder&subpath="& current_sub_folder & ">Delete</a>"
			.Write "<a href="&g_sScriptName&"?webform=webform_system_fileexplorer_upload&subpath="& current_sub_folder & ">Upload</a>"
			.Write "</table>"
			
			
			'-- Open the main table for the file explorer
			.Write "<table width=100% cellpadding=0 cellspacing=0 >"
			.Write "<tr valign=top>"
			
			'-- right cell is for the preview pane
			.Write "<form method=post action="&g_sUrl&">"
			.Write "<td id=webpreview width=152>"
				.Write "<img src=modules/system/media/Wvleft.png><br>"
				.Write "<b style='font: 14px arial; font-weight: bold;'>" & current_user_root_folder & current_sub_folder & "</b><br>"
				.Write "<img src=modules/system/media/Wvline.png>"
				.Write "<div id=fileinfos style='display: hidden'>"
				'.Write "<a href=>delete this file</a><br>"
				'.Write "<a id=select href='#'>select this file</a><br>"				
				.Write "</div><br>"
				.Write "<div id=preview></div>"
			.Write "</td>"
			.Write "</form>"
			
			
			'-- Left pane is for folders and files list
			.Write "<td id=fileexplorer>"
					
			.Write "<table class=datagrid cellpadding=0 cellspacing=0>"
			.Write "<thead><tr class=datagrid_column><th width=100% >" & String("system", "tool_fileexplorer", "filename") & "</th><th>" & String("system", "tool_fileexplorer", "filesize") & "</th><th nowrap=true>" & String("system", "tool_fileexplorer", "filetype") & "</th><th nowrap=true>" & String("system", "tool_fileexplorer", "filemodified") & "</th></tr></thead>"
			
			.Write "<tbody id=fileTblBody>"
			
			'-- If not in root, display a link to parent Folder
			If len(current_sub_folder)>1 then

				Dim ParentPath
				ParentPath = mid(current_sub_folder, 1, len(current_sub_folder)-1)
				ParentPath = left(ParentPath, instrrev(ParentPath, "/"))
				
				.Write "<tr class=file>"
				.Write "	<td><img src=" & appSettings("MODULES_FOLDER") & "/system/media/folder.png align=absmiddle> <a href="&g_sScriptName&"?webform=webform_system_fileexplorer&SubPath=" & ParentPath  & ">..</td>"   '<b>" & String("system", "tool_fileexplorer", "parentfolder") & "</b>
				.Write "	<td align=right></td><td></td><td></td>"
				.Write "</tr>"
			End If
			
			'-- Loop on subfolder
			For each oSub in oUploadFolder.SubFolders
				.Write "<tr class=file>"
				.Write "	<td><img src=" & appSettings("MODULES_FOLDER") & "/system/media/folder.png align=absmiddle> <a href="""&g_sScriptName&"?webform=webform_system_fileexplorer&SubPath=" & current_sub_folder  & oSub.Name & "/"">" & oSub.name & "</td>"
				.Write "	<td align=right>~</td><td nowrap=true>" & oSub.Type & "</td><td nowrap=true>" & oSub.DateLastModified & "</td>"
				.Write "</tr>"
			Next
			
			'-- Loop on files
			For each oFile in oUploadFolder.Files
				.Write "<tr class=file>"
				.Write "	<td><img src=" & appSettings("MODULES_FOLDER") & "/system/media/file.png align=absmiddle> <a href=# onfocus=""showOptions(0, '" & MEDIA_FOLDER & current_user_root_folder & current_sub_folder & oFile.name&"')"" >" & oFile.name & "</td>"
				.Write "	<td align=right nowrap=true>" & Round(oFile.Size/1024) & " KB</td><td nowrap=true>" & oFile.Type & "</td><td nowrap=true>" & oFile.DateLastModified & "</td>"
				.Write "</tr>"
			Next
			
			.Write "</tbody></table>"
			.Write "</td></tr></table><br>"
			
			
		End With
	End Sub
	
	
	'---------------------------------------------
	'-- Display the form to create a new folder --
	'---------------------------------------------
	Public Sub webform_system_fileexplorer_newfolder
		with Response			
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<caption>" & String("system", "tool_fileexplorer", "newfolder") & " : " & current_sub_folder & "</caption>"
			.Write "<form method='GET' action='" & g_sURL & "'>"
			.Write "<input name=webform type=hidden value='webform_system_fileexplorer'>"
			.Write "<input name=SubPath type=hidden value='" & current_sub_folder & "'>"
			.Write "<input name=process type=hidden value='do_create_folder'>"
		
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_fileexplorer", "newfolder") & "</th><td><input type=text name=foldername class=medium></td></tr>"
			.Write "<tr class=datagrid_buttonrow><td colspan=2 align=right><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value='"&String("system", "common", "back")&"' onclick='history.back(-1);'></td></tr>"
					
			.Write "</form></table><br>"
		end with
	End Sub
	
	
	'---------------------------------------------
	'-- Display a form to delete current folder --
	'---------------------------------------------
	Public Sub webform_system_fileexplorer_deletefolder
		if len(current_sub_folder)>1 then
			with Response			
				.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
				.Write "<caption>" & String("system", "tool_fileexplorer", "deletefolder") & ": " & current_sub_folder &  "</caption>"
				.Write "<form method='POST' action='" & g_sURL & "' name='frmDelete'>"
				.Write "<input name=webform type=hidden value='webform_system_fileexplorer'>"
				.Write "<input name=SubPath type=hidden value='" & current_sub_folder & "'>"
				.Write "<input name=process type=hidden value='do_delete_folder'>"
			
				.Write "<tr><td align=center>" & String("system", "common", "confirmdelete") & "</td></tr>"
				.Write "<tr class=datagrid_buttonrow><td colspan=2 align=right><input type=button value='" & String("system", "common", "ok") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms['frmDelete'].submit();}"">&nbsp;<input type=button value='"&String("system", "common", "back")&"' onclick='history.back(-1);'></td></tr>"
						
				.Write "</form></table><br>"
			end with
		End If
	End Sub
	
	
	'-- Display the Upload files Form --
	Private Sub webform_system_fileexplorer_upload
		with Response			
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<caption>" & String("system", "tool_fileexplorer", "uploadfiles") & ": "& current_sub_folder & "</caption>"
			.Write "<form name='frmSend' method='POST' enctype='multipart/form-data' action='" & g_sScriptName & "?process=do_upload_files&subpath="&current_sub_folder&"'>"
						
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_fileexplorer", "file") & "</th><td><input type=file name=attach1 class=large></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_fileexplorer", "file") & "</th><td><input type=file name=attach2 class=large></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_fileexplorer", "file") & "</th><td><input type=file name=attach3 class=large></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_fileexplorer", "file") & "</th><td><input type=file name=attach4 class=large></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_fileexplorer", "file") & "</th><td><input type=file name=attach5 class=large></td></tr>"
			.Write "<tr class=datagrid_buttonrow><td colspan=2 align=right><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value='"&String("system", "common", "back")&"' onclick='history.back(-1);'></td></tr>"
			.Write "</form></table><br>"
		end with
	End Sub
	
	
	
	'-------------------------
	'-- Create a new folder --
	'-------------------------
	Sub Do_Create_Folder 
			
		Dim foldername	: foldername = Request.QueryString("foldername")
		Call CreateFolder(upload_path & current_user_root_folder & current_sub_folder & foldername)
		
		Response.Redirect g_sScriptName & "?webform=webform_system_fileexplorer&SubPath=" & current_sub_folder
	End Sub
	
	
	'---------------------
	'-- Delete a folder --
	'---------------------
	Sub Do_Delete_Folder
		Dim SubPath		
		
		If instr(1, Request.QueryString("SubPath"), "..") then
			Die "I say fuck that ;-)"
		End if	
		
		'-- do the delete
		Call DeleteFolder(replace(upload_path & current_user_root_folder & current_sub_folder, "/", "\"))
		
		'-- redirect to the parent folder
		Dim ParentPath
		ParentPath = mid(current_sub_folder, 1, len(current_sub_folder)-1)
		ParentPath = left(ParentPath, instrrev(ParentPath, "/"))
		
		Response.Redirect g_sScriptName & "?webform=webform_system_fileexplorer&SubPath=" &  ParentPath
		
	End Sub
	
	
	'------------------
	'-- Upload files --
	'------------------
	Sub Do_Upload_Files
				
		'-- Call the upload class
		Dim oUpload
		Set oUpload = New FreeASPUpload
		oUpload.Save(upload_path & current_user_root_folder & current_sub_folder)
		Set oUpload = Nothing
		
		Response.Redirect g_sScriptName & "?webform=webform_system_fileexplorer&SubPath=" &  current_sub_folder
	End Sub
	
	
	'-- check the system for upload capabilities ---------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	function System_CheckRequierement()
		Dim fso, fileName, testFile, streamTest
		TestEnvironment = ""
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if not fso.FolderExists(upload_path) then
			TestEnvironment = "<B>Folder " & upload_path & " does not exist.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
			exit function
		end if
		fileName = upload_path & "\test.txt"
		on error resume next
		Set testFile = fso.CreateTextFile(fileName, true)
		If Err.Number<>0 then
			TestEnvironment = "<B>Folder " & upload_path & " does not have write permissions.</B><br>The value of your uploadsDirVar is incorrect. Open uploadTester.asp in an editor and change the value of uploadsDirVar to the pathname of a directory with write permissions."
			exit function
		end if
		Err.Clear
		testFile.Close
		fso.DeleteFile(fileName)
		If Err.Number<>0 then
			TestEnvironment = "<B>Folder " & upload_path & " does not have delete permissions</B>, although it does have write permissions.<br>Change the permissions for IUSR_<I>computername</I> on this folder."
			exit function
		end if
		Err.Clear
		Set streamTest = Server.CreateObject("ADODB.Stream")
		If Err.Number<>0 then
			TestEnvironment = "<B>The ADODB object <I>Stream</I> is not available in your server.</B><br>Check the Requirements page for information about upgrading your ADODB libraries."
			exit function
		end if
		Set streamTest = Nothing
	end function
	
	
	'-- Display a control to select picture --------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Function HtmlComponent_SelectImage(sName, sValue)
		Dim t
				
		t = t & "<input type=text class=medium name=" & sName & " value=""" & sValue & """>&nbsp;"
				
		If g_oFSO.FolderExists(upload_path) then
			t = t & "<select name=sel_" & sName & " style='width: 200px; height: 50px' onchange='"&sName&".value = sel_"&sName&".options[sel_"&sName&".selectedIndex].value;'>" & SelectImageFolder(upload_path & current_user_root_folder, sValue, 0)  & "</select>"
			't = t & "<a href=# onclick='window.open(sel_" & sName & ".options[sel_" & sName & ".options.selectedIndex].value, ""preview"", ""width=200, height=200, resizable=1"");'>" & String("system", "common", "preview") & "</a>"
			t = t & "<a href=# onclick='window.open(" & sName & ".value, ""preview"", ""width=200, height=200, resizable=1"");'>" & String("system", "contenttype_image", "preview") & "</a>"
		End If
		
		HtmlComponent_SelectImage = t
	End Function
	
	
	'-- used by HtmlComponent_SelectImage, recursively craw folders --------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Private Function SelectImageFolder(sPath, sValue, decay)
		Dim t
		Dim oUploadFolder, oSubFolder, oFile
		Set oUploadFolder = g_oFSO.GetFolder(sPath)
		
		t = t & "<option>" & repeat(decay, "&nbsp;") & "[" & ucase(oUploadFolder.name) & "]</option>"
		
		For each oSubFolder in oUploadFolder.SubFolders
			t = t & SelectImageFolder(oSubFolder.Path, sValue, decay + 4)
		Next
		
		For each oFile in oUploadFolder.Files
			Dim imgpath: imgpath = mid(oFile.path, len(g_sServerMapPath)+2)
			t = t & "<option value='" & imgpath &"'" & iff(imgpath=sValue, " selected", "") & ">" & repeat(decay+4, "&nbsp;") & oFile.Name & "</option>"
		Next
		
		SelectImageFolder = t
	End Function
%>