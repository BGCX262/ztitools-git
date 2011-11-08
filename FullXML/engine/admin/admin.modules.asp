<%	'=====================================================================================================
	' THIS FILE CONTAINS ALL THE MODULES RELATED FUNCTIONS
	'=====================================================================================================
	
	
	'-------------------------------------------
	'-- Display the list of detected modules  --
	'-------------------------------------------
	SUB webform_list_modules
		Dim oModulesFolder, oModuleFolder, modulename
		Set oModulesFolder = g_oFSO.GetFolder(MODULES_FOLDER)
	
	    With Response
	        .Write "<h3>" & String("system", "modules", "modules") & "</h3>"
	        .Write String("system", "modules", "modulelist")
	        .write "<form method=post action="&g_sUrl&" name=frmModules id=frmModules>"
			.Write "<input type=hidden name=process value=do_create_moduleslist>"
			
            For each oModuleFolder in oModulesFolder.SubFolders

				modulename = oModuleFolder.name
				If UCase(modulename)<>"CVS" then
					.Write "<input type=checkbox id="&modulename&" value=" & modulename & " name=modules " & iff(modulename="system", "checked disabled" , "") & iff(Application(APPVAR_DOM_MODULES).SelectNodes("/modules/module[@name='" & modulename & "']").length=1, " checked", "") & "> <label for="&modulename&">"&modulename&"</label><br>"
				End If

			Next

            .Write "<br>&nbsp;&nbsp;<a href='javascript:frmModules.submit()' >"&String("system", "modules", "update")&"</a>"
			.write "</form>"
	    End With
	
	End Sub
	
	
	'-------------------------------------------------------
	'-- Create the modules list depending on admin choice --
	'-------------------------------------------------------
	Sub Do_Create_ModulesList
 		
 		'-- Calculate the list of modules to install
 		Dim arrModules : arrModules = split(getParam("modules"), ", ")
 		redim preserve arrModules(UBound(arrModules)+1)
 		arrModules(UBound(arrModules)) = "system"
 		 		
 		'-- 1st step: create the modules.xml file
 		Call Do_Create_Modules_XMLFile(arrModules)
 		
 		'-- 2st step: create the modules.includes.asp
		Call Do_Create_Modules_IncludeFile()
		
		'-- 3rd step: reload the modules.xml in memory
		Call LoadModulesInMemory(true)
		
		'-- 4th step: redirect to the modules list
		Response.Redirect g_sScriptName & "?webform=webform_list_modules"
		
	End Sub
	
		
	Sub EditModule
		
	End Sub
		
	
	'-----------------------------------------------------
	'-- Build the modules.xml file from the File System --
	'-----------------------------------------------------
	FUNCTION AutoDetectModules		
		CALL Do_Create_ModulesList()		
	END FUNCTION
	
	
	'---------------------------------
	'-- Create the modules.xml file --
	'---------------------------------
	Sub Do_Create_Modules_XMLFile(p_arrModules)
		Dim oModuleFolder, oCultureFile
		Dim oModule
		
		LogIt "admin.modules.asp", "Do_Create_Modules_XMLFile", INFO, "BEGIN modules loading.", ""
				
		'-- Create the domdocument
		Dim oXML
		Set oXML = CreateDomDocument
				
		'-- Create the root node (modules)
		Dim oRoot : Set oRoot = oXML.CreateElement("modules")
		
		
		'-- Loop on each module folder
		Dim i
		For i=LBound(p_arrModules) to UBound(p_arrModules) 
			
			LogIt "admin.modules.asp", "Do_Create_Modules_XMLFile", INFO, "Module " & p_arrModules(i) & " detected.", ""
		
			Set oModuleFolder = g_oFSO.GetFolder(MODULES_FOLDER & p_arrModules(i))
			
			'-- try to load the module configuration file
			If oXML.Load (oModuleFolder.Path & "\module.xml") then
				
				Set oModule = oXML.DocumentElement.CloneNode(true)
				
				'-----------------------------------------------------------					
				'-- load each culture, and append them to the module node --
				if g_oFSO.FolderExists(oModuleFolder.path & "\Cultures") then
					For Each oCultureFile in g_oFSO.GetFolder(oModuleFolder.path & "\Cultures").Files
						if oXML.Load (oCultureFile.Path) then
							oModule.AppendChild(oXML.DocumentElement.CloneNode(true))
						else
							Response.Write "admin/modules.asp line 127: " & oXML.parseerror.reason & " [" & oCultureFile.Path & "]"
						end if
					Next
				End if
				
				
				'---------------------------------
				'-- append the permissions file --
				If oXML.Load (oModuleFolder.Path & "\permissions.xml") Then
					oModule.AppendChild(oXML.DocumentElement.CloneNode(true))
				End If
				
				'-- append the module node to the root
				oRoot.appendChild(oModule.CloneNode(true))
			Else
				LogIt "admin.modules.asp", "Do_Create_Modules_XMLFile", ERROR, "Fail to load module " & oModuleFolder.Name & ".", oModuleFolder.Path & "\module.xml"
			End If
			
		Next
		
		
		LogIt "admin.modules.asp", "Do_Create_Modules_XMLFile", INFO, "END modules loading.", ""
		
		
		'-- clean the document, cause it contains the last module.xml
		Dim ochild
		For each ochild in oXML.ChildNodes
			oXML.removeChild(ochild)
		Next 
		
		
		'-- Add the processing instruction
		Dim pi : Set pi = oXML.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
		oXML.appendChild(pi)
		
		'-- adding asp comment
		Dim oComment : Set oComment = oXML.CreateComment(" <% Response.End %"&">")
		oXML.appendChild oComment
		
		'-- add the root element
		oXML.appendChild(oRoot)
		
		'-- save the document
		oXML.Save modules_xml 
				
		'-- release object
		Set oXML = Nothing
	End Sub
	
	
	'-- 
	Sub Do_Create_Modules_IncludeFile()
		Dim oXML, oModule, oContentType, oTool, oFunction
		
		Set oXML = CreateDomDocument		
		if not oXML.Load (modules_xml) then
			LogIt "admin.modules.asp", "CreateModulesIncludeFile", ERROR, oXML.ParseError.Reason, modules_xml
			die "Can't load " & DATA_FOLDER & MODULES_FILE
		end if
		
		
		'-- 1st, create the contenttypes file							
		Dim ModulesFile
		Set ModulesFile = g_oFSO.CreateTextFile(MODULES_FOLDER & "modules.asp", true, false)
			
		'-- Loop on each module
		for each oModule in oXML.DocumentElement.SelectNodes("/modules/module")
						
			'-- the module configuration function file
			ModulesFile.WriteLine "<!-- #include file=""" & GetAttribute(oModule, "name", "") & "/_module.asp""  -->"
			
			'-- add the content types
			for each oContentType in oModule.SelectNodes("contenttypes/contenttype")
				ModulesFile.WriteLine "<!-- #include file=""" & GetAttribute(oModule, "name", "") & "/" & GetAttribute(oContentType, "filename", "") & """  -->"
			next
						
			'-- add the tools
			for each oTool in oModule.SelectNodes("tools/tool")
				ModulesFile.WriteLine "<!-- #include file=""" & GetAttribute(oModule, "name", "") & "/" & GetAttribute(oTool, "filename", "") & """  -->"
			next
			
			'-- add the functions
			for each oFunction in oModule.SelectNodes("functions/function")
				ModulesFile.WriteLine "<!-- #include file=""" & GetAttribute(oModule, "name", "") & "/" & GetAttribute(oFunction, "filename", "") & """  -->"	
			next
			
		next
						
		ModulesFile.Close
		set ModulesFile = Nothing
		Set oXML = Nothing		
	End Sub	
	
	
	
'	Sub CreateModulesXMLFile
'		Dim oModulesFolder, oModuleFolder, oCultureFile
'		Set oModulesFolder = g_oFSO.GetFolder(MODULES_FOLDER)
'		
'		LogIt "admin.modules.asp", "CreateModulesXMLFile", INFO, "BEGIN modules loading.", ""
'			
'		
'		'-- Create the domdocument
'		Dim oXML : Set oXML = CreateDomDocument
'				
'		'-- Create the root node (modules)
'		Dim oRoot : Set oRoot = oXML.CreateElement("modules")
'		
'		
'		'-- Loop on each module folder
'		For each oModuleFolder in oModulesFolder.SubFolders
'			
'			'
'			If UCase(oModuleFolder.name)<>"CVS" then
'			
'				LogIt "admin.modules.asp", "CreateModulesXMLFile", INFO, "Module " & oModuleFolder.Name & " detected.", ""
'			
'				'-- try to load the module configuration file
'				If oXML.Load (oModuleFolder.Path & "\module.xml") then
'					Dim oModule
'					Set oModule = oXML.DocumentElement.CloneNode(true)
'											
'					'-- load each culture, and append them to the module node
'					if g_oFSO.FolderExists(oModuleFolder.path & "\Cultures") then
'						For Each oCultureFile in g_oFSO.GetFolder(oModuleFolder.path & "\Cultures").Files
'							if oXML.Load (oCultureFile.Path) then
'								oModule.AppendChild(oXML.DocumentElement.CloneNode(true))
'							else
'								Response.Write "can't load : " & oCultureFile.Path
'							end if
'						Next
'					End if
'					
'					'-- append the module node to the root
'					oRoot.appendChild(oModule.CloneNode(true))
'				Else
'					LogIt "admin.modules.asp", "CreateModulesXMLFile", ERROR, "Fail to load module " & oModuleFolder.Name & ".", oModuleFolder.Path & "\module.xml"
'				End If
'				
'			End if
'		Next
'		
'		
'		LogIt "admin.modules.asp", "CreateModulesXMLFile", INFO, "END modules loading.", ""
'		
'		
'		'-- clean the document, cause it contains the last module.xml
'		Dim ochild
'		For each ochild in oXML.ChildNodes
'			oXML.removeChild(ochild)
'		Next 
'		
'		
'		'-- Add the processing instruction
'		Dim pi : Set pi = oXML.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
'		oXML.appendChild(pi)
'		
'		'-- adding asp comment
'		Dim oComment : Set oComment = oXML.CreateComment(" <% Response.End %"&">")
'		oXML.appendChild oComment
'		
'		'-- add the root element
'		oXML.appendChild(oRoot)
'		
'		'-- save the document
'		oXML.Save modules_xml 
'				
'		'-- release object
'		Set oXML = Nothing
'	End Sub
	
	
	'-- Load each available modules ----------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Sub LoadModulesInMemory(bForceUpdate)
		
		'Set Application(APPVAR_DOM_MODULES) = Nothing
		'Application(APPVAR_DOM_MODULES) = Empty
		
		If bForceUpdate OR (not isObject(Application(APPVAR_DOM_MODULES))) or Request.QueryString("force")="yes" Then
			Dim oXML
			
			'-- clean old datas
			Set Application(APPVAR_DOM_MODULES) = Nothing
			Application(APPVAR_DOM_MODULES) = Empty
		
			'-- create and load the new dom
			Set oXML = CreateFreeDomDocument
			If NOT oXML.load(modules_xml) Then
				LogIt "admin.modules.asp", "LoadModulesInMemory", ERROR, oXML.ParseError & " : " & oXML.ParseError.Reason, modules_xml
				
				if oXML.ParseError=-2146697210 then
					Do_Create_ModulesList
					If NOT oXML.load(modules_xml) Then
						LogIt "admin.modules.asp", "LoadModulesInMemory", FATAL, "Can't create the modules file.", modules_xml
						FatalError "Fatal Error", "A fatal error occured. [ref fx004 : Can't Create the modules file.]"
					End If
				end if
			
			End If	
			
			Set Application(APPVAR_DOM_MODULES) = oXML
		
		End If		
	End Sub
	
	
'	'---------------------------------
'	'-- Display the insert/edit form
'	Sub EditModule
'		
'		ModuleProcess
'		
'		Dim moduleID : moduleID = Request.QueryString("id")
'		Dim process : process = "do_insert_module"
'		Dim name, systemname, author, version, url
'		Dim oXML, oNodeList, oNode
'				
'		Set oXML = CreateDomDocument
'		
'		if not oXML.Load (g_sServerMapPath & DATA_FOLDER & MODULES_FILE) then
'			LogIt "admin.modules.asp", "EditModule", ERROR, "Can't load xml", g_sServerMapPath & DATA_FOLDER & MODULES_FILE
'		end if
'		
'		'-- If an id is passed, then we are editing the data, so load the old value
'		if len(moduleID)>0 Then
'			process = "do_update_module"
'			
'			Set oNodeList = oXML.SelectNodes("/modules/module[@id='" & moduleID & "']")	
'			
'			if oNodeList.length>0 then				
'				name = GetAttribute(oNodeList(0), "name", "")
'				systemname = GetAttribute(oNodeList(0), "systemname", "")
'				author = GetAttribute(oNodeList(0), "author", "")
'				version = GetAttribute(oNodeList(0), "version", "")
'				url = GetAttribute(oNodeList(0), "url", "")
'			end if		
'		End If
'		
'				
'		With Response
'			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
'			.Write "<form action=" & g_sURL & " method=post id=frmEdit name=frmEdit>"
'			.Write "<input type=hidden name=process value='" & process & "'>"
'			.Write "<input type=hidden name=moduleID value='" & moduleID & "'>"
'			.Write "<caption>" & getString("module") & "</caption>"
'			
'			.Write "<tr class=datagrid_editrow><th>" & getString("name") & "</th><td><input type=text class=large name=name value='" & name & "'></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & getString("systemname") & "</th><td><input type=text class=large name=systemname value=""" & systemname & """></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & getString("author") & "</th><td><input type=text class=large name=author value=""" & author & """></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & getString("version") & "</th><td><input type=text class=small name=version value=""" & version & """></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & getString("url") & "</th><td><input type=text class=large name=url value=""" & url & """></td></tr>"
'			
'			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & getString("ok") & "'>&nbsp;<input type=button value=""" & getString("cancel") & """ onclick=""document.location='" & g_sScriptName & "?webform=webform_list_modules';""></td></tr>"
'			if len(moduleID) then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & getString("delete") & "' onclick=""if (confirm('" & getString("confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_module';document.forms[0].submit();}""></td></tr>"
'			.Write "</form>"
'			.Write "</table>"			
'		End With
'		
'		Set oXML = Nothing
'		
'	End Sub
	
	
'	'-- process action
'	Sub ModuleProcess
'		
'		Select case Request.Form("process")
'			case "do_update_module"
'				Call UpdateNode (DATA_FOLDER & MODULES_FILE, "/modules/module[@id='" & getParam("moduleID") & "']", Array("name", "systemname", "author", "version", "url"), Array("name", "systemname", "author", "version", "url"))
'				CreateModulesIncludeFile
'				Response.Redirect g_sScriptName & "?webform=webform_list_modules"
'			
'			case "do_insert_module"
'				Call InsertNode (DATA_FOLDER & MODULES_FILE, "/modules" , "module", Array("name", "systemname", "author", "version", "url"), Array("name", "systemname", "author", "version", "url"), true, "")
'				CreateModulesIncludeFile
'				Response.Redirect g_sScriptName & "?webform=webform_list_modules"
'			
'			case "do_delete_module"
'				Call DeleteNode (DATA_FOLDER & MODULES_FILE, "/modules/module[@id='" & getParam("moduleID") & "']")
'				CreateModulesIncludeFile
'				Response.Redirect g_sScriptName & "?webform=webform_list_modules"
'		'
'		End Select
'	End Sub
	
	
	
%>