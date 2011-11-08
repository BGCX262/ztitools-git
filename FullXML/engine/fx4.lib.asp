<%
	'-----------------------------------------------------------------
	'-- This file contains various functions used by the Fx4 engine --
	'-- Author:		John Roland
	'-- Modified:	2003/11/10
	'-----------------------------------------------------------------
	
	'---------------------------------------------------------------------
	'-- Return the permission level on a content, for a specified group --
	'---------------------------------------------------------------------
	Public Function readUserPermission(p_sXPath, p_sUser, p_sGroup)
		Dim bExit 		: bExit = false
		Dim oNodeList, oNode
		
		'-- select the xpath  of the current object
		set oNodeList = g_oWebSiteXML.SelectNodes(p_sXPath)
		
		'-- if the object does not exist
		If oNodeList.Length=0 Then
			readUserPermission = GetDefautLevel(p_sGroup)
			Exit Function
		Else
			Set oNode = oNodeList.Item(0)
		End If
		
		
		'-- We loop back the tree for permission
		Do While Not bExit
			
			'-- Simple : a permission is set for the object
			If oNode.SelectNodes("permission[@user='" & p_sUser  & "']").length=1 Then
				readUserPermission = cint(getAttribute(oNode.SelectSingleNode("permission[@user='" & p_sUser  & "']"), "value", "CONST_ACCESS_LEVEL_VIEWER"))
				bExit = true
				
			'-- Complex : we loop back to parenNodes in search for a permission
			Else
				if not isNUll(oNode.parentNode) Then
					If oNode.parentNode.nodeName="website" OR oNode.parentNode.nodeName="menu" OR oNode.parentNode.nodeName="page" Then
						Set oNode = oNode.parentNode
						bExit = False
					Else
						bExit = True
					End If
				End if							
			End if			
		Loop
		
		
		'-- In case no permission were found, select the default one
		If len(readUserPermission)=0 then
			readUserPermission = readGroupPermission(p_sXPath, p_sGroup)
		End If
						
	End Function
	
	
	'---------------------------------------------------------------------
	'-- Return the permission level on a content, for a specified group --
	'---------------------------------------------------------------------
	Public Function readGroupPermission(p_sXPath, p_sGroup)
		Dim bExit 		: bExit = false
		Dim oNodeList, oNode
		
		'-- select the xpath  of the current object
		set oNodeList = g_oWebSiteXML.SelectNodes(p_sXPath)
		
		'-- if the object does not exist
		If oNodeList.Length=0 Then
			readGroupPermission = GetDefautLevel(p_sGroup)
			Exit Function
		Else
			Set oNode = oNodeList.Item(0)
		End If
		
		
		'-- We loop back the tree for permission
		Do While Not bExit
			
			'-- Simple : a permission is set for the object
			If oNode.SelectNodes("permission[@group='" & p_sGroup  & "']").length=1 Then
				readGroupPermission = cint(getAttribute(oNode.SelectSingleNode("permission[@group='" & p_sGroup  & "']"), "value", "CONST_ACCESS_LEVEL_VIEWER"))
				bExit = true
				
			'-- Complex : we loop back to parenNodes in search for a permission
			Else
				if not isNUll(oNode.parentNode) Then
					If oNode.parentNode.nodeName="website" OR oNode.parentNode.nodeName="menu" OR oNode.parentNode.nodeName="page" Then
						Set oNode = oNode.parentNode
						bExit = False
					Else
						bExit = True
					End If
				End if							
			End if			
		Loop
		
		
		'-- In case no permission were found, select the default one
		If len(readGroupPermission)=0 then
			readGroupPermission = GetDefautLevel(p_sGroup)
		End If
						
	End Function
		
	
	'------------------------------------------------------------
	'-- Return the default access level for the BUILDIN groups --
	'-- This level is different for each website model
	'------------------------------------------------------------
	Private Function GetDefautLevel(m_sGroup)
		Dim website_model : website_model = GetAttribute(g_oWebSiteXML.documentElement, "model", CONST_WEBSITE_MODEL_PRIVATE)
		
		'-- administrator case is hard coded for safety
		If m_sGroup = "administrator" Then
			GetDefautLevel = CONST_ACCESS_LEVEL_ADMINISTRATOR
			Exit Function
		End If
		
		'-- check if the modules are loaded
		LoadModulesInMemory false
		
		'-- read in the config file
		Dim oNodeList
		Set oNodeList = Application(g_sServerName & "_appSettings").DocumentElement.SelectNodes("/configuration/permissions/GetDefautLevel/model[@id='"&website_model&"']/group[@id='" & m_sGroup & "']")
		If oNodeList.Length=1 Then
			dim tmp : tmp = getAttribute(oNodeList.item(0), "level", "CONST_ACCESS_LEVEL_VIEWER")
			GetDefautLevel = cint(eval(tmp))
		Else
			'todo : set this depending of the website model			
			GetDefautLevel = CONST_ACCESS_LEVEL_VIEWER
		End If	
		
	End Function
	
'	Private Function readDefaultObjectPermission(p_sModuleName, m_sObject, m_sGroup)
'		Dim website_model : website_model = GetAttribute(g_oWebSiteXML.documentElement, "model", CONST_WEBSITE_MODEL_PRIVATE)
'		
'		'-- administrator case is hard coded for safety
'		if m_sGroup = "administrator" then
'			readDefaultObjectPermission = CONST_ACCESS_LEVEL_ADMINISTRATOR
'			Exit Function
'		end if
'		
'		'-- check if the modules are loaded
'		LoadModulesInMemory false
'		
'		'-- "System" module by default
'		IFF Len(p_sModuleName)=0, "system", p_sModuleName
'		
'		Dim oNodeList
'		Set oNodeList = Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[@name='"&p_sModuleName&"']/permissions/object[@id='" & m_sObject  & "']/model[@id='"&website_model&"']/group[@id='" & m_sGroup & "']")
'		If oNodeList.Length=1 Then
'			dim tmp : tmp = getAttribute(oNodeList.item(0), "level", "CONST_ACCESS_LEVEL_VIEWER")
'			readDefaultObjectPermission = cint(eval(tmp))
'		Else
'			'todo : set this depending of the website model			
'			readDefaultObjectPermission = CONST_ACCESS_LEVEL_VIEWER
'		End If	
'	
'	End Function
	
	
		
	'------------------------------------------------------
	'-- Execute the process with permission verification --
	'-- TODO: incorporate the webmaster processes
	'------------------------------------------------------
	Private Sub ExecuteProcess()
		'-- pointer to the modules.xml
		Dim oModulesXML : set oModulesXML = Application(APPVAR_DOM_MODULES)		
		
		if lenb(g_sProcess)=0 then
			EXIT SUB
		end if
		
		'-- admin can execute it all
		If g_oUser.Group = "administrator" Then			
			Call Execute(g_sProcess)
		
		'TODO: add the check of unauthorized process (modules, skins, administrators, webmasters, etc...)
		ElseIf g_oUser.Group = "webmaster" Then
			Call Execute(g_sProcess)
		
		ElseIf g_oUser.Group = "anonymous" and oModulesXML.SelectNodes("/modules/module/permissions/anonymous/action[@id='"&g_sProcess&"']").length>0 Then
			Call Execute(g_sProcess)
		
		ElseIf oModulesXML.SelectNodes("/modules/module/permissions/authenticated/action[@id='"&g_sProcess&"']").length>0 Then
			Call Execute(g_sProcess)
				
		'-- check by object/level explicit permission	
		ElseIf oModulesXML.SelectNodes("/modules/module/permissions/level[@id <= "& g_oUser.PagePermissionLevel &"]/action[@id='"&g_sProcess&"']").length>0 Then
			Call Execute (g_sProcess)							
		
		'-- access denied
		Else			
			FatalError String("system", "permissions", "accessdenied"), String("system", "permissions", "accessdeniedmsg") & " [" & g_sWebform & "]"
			LogIt "fx4.lib.asp", "ExecuteProcess", WARNING, "Attempt to execute unauthorized process", g_oUser.Login & " try to execute :" & g_sProcess 
		End If
		
		
		'-- Close the popup
		If InStr(g_sScriptName, "popup.asp")>0 Then
			Call ReloadMainFrameAndClose()
			Response.End
		End If
		
		'-- Redirect after process
		If lenB(getParam("afterprocesswebform"))>0 Then
			Call AfterProcessWebform(getParam("afterprocesswebform"))
			Response.End
		End If
			
	End Sub
	
	
	'-----------------------
	'-- Show the web form --
	'-----------------------
	'TODO: incorporate the webmaster webforms
	Private function DisplayWebForm
		'-- pointer to the modules.xml
		Dim oModulesXML : set oModulesXML = Application(APPVAR_DOM_MODULES)
					
				
		if g_oUser.Group = "administrator" Then			
			Call Execute(g_sWebform)
				
		Elseif g_oUser.Group = "webmaster" Then			
			Call Execute(g_sWebform)
		
		'-- Webforms authorized for everyone, even anonymous users
		ElseIf oModulesXML.SelectNodes("/modules/module/permissions/anonymous/webform[@id='"&g_sWebform&"']").length>0 Then
			Call Execute(g_sWebform)
		
		ElseIf oModulesXML.SelectNodes("/modules/module/permissions/authenticated/webform[@id='"&g_sWebform&"']").length>0 Then
			Call Execute(g_sWebform)
		
		ElseIf oModulesXML.SelectNodes("/modules/module/permissions/level[@id <= "& g_oUser.PagePermissionLevel &"]/webform[@id='"&g_sWebform&"']").length>0 Then
			Call Execute (g_sWebform)							
		
		Else			
			FatalError String("system", "permissions", "accessdenied"), String("system", "permissions", "accessdeniedmsg") & " [" & g_sWebform & "]"
		End If
		
	End Function
	
	Sub webform_default
		Response.Write "<table width=650 height=400 style='border: 1px outset; background-color: F0F0F0'><tr><td align=center><b>Fullxml4</b><br> Developped by John Roland<br><a href=mailto:john.roland@fullxml.com>john.roland@fullxml.com</a><br><a href=http://www.fullxml.com>http://www.fullxml.com</a></td></tr></table>"
	End Sub
	
	'-- this function is called after a process to avoid a double execution of the process
	Sub AfterProcessWebform(p_Webform)
		response.Redirect g_sScriptName & "?webform=" & unescape(p_Webform)
	End Sub
	
	'------------------------------------------
	'-- Load the xml doms into memory		 --
	'-- And initialize some global variables --
	'------------------------------------------
	' todo How can we cache the doms
	Sub LoadDOMInMemory(bForce)
		
				
		'---------------------------
		'-- Load the website data --
		Set g_oWebSiteXML = CreateFreeDomDocument
		If not g_oWebSiteXML.Load (website_xml) then
			LogIt "loader.dom.asp", "LoadDOMInMemory", FATAL, "Fail to load website data: " & website_xml, g_oWebSiteXML.ParseError.Reason
			FatalError "Under Construction", "Our website is currently under construction, please check back later."
		End if
		
		
		'----------------------------
		'-- load the language list --
		Set g_oLanguagesXML = CreateFreeDomDocument
		If NOT g_oLanguagesXML.Load(languages_path) then
			LogIt "loader.dom.asp", "LoadDOMInMemory", FATAL, "Can't load languages.xml", g_oLanguagesXML.parseerror.reason
			FatalError "Fatal error", "Our website is currently under construction, please check back later."
		Else
			'-- set the culture, lcid and encoding global variables
			Dim oLanguageList			
			Set oLanguageList = g_oLanguagesXML.SelectNodes("languages/language[@id='" & getAttribute(g_oWebSiteXML.documentElement, "culture", "en") & "']")
								
			If oLanguageList.Length=1 then				
				g_sCulture = getAttribute(oLanguageList.item(0), "file", "en")
				g_iLCID = cint(getAttribute(oLanguageList.item(0), "lcid", "1033"))
				g_sEncoding = getAttribute(oLanguageList.item(0), "encoding", "utf-8")		
			Else
				g_sCulture = "en-us"
				g_iLCID = 1033
				g_sEncoding = "utf-8"
			End If
		End If
		
		
		'-- IF NOT in ADMIN -> Default webpage selection
		if instr(1, g_sScriptName, "admin.asp")=0 then
			if len(g_sPageID)=0 then
				Dim oWebpageList
				Set oWebpageList = g_oWebSiteXML.SelectNodes("website/menus/menu/page")
				If oWebpageList.length>0 then
					g_sPageID = oWebpageList.item(0).Attributes.GetNamedItem("id").Value
				Else
					LogIt "loader.dom.asp", "LoadDOMInMemory", FATAL, "No page for the website ", website_xml
					FatalError "Under Construction", "Our website is currently under construction, please check back later."
				End if
			End if
			
			If Len(g_sMenuID)=0 Then
				Dim oMenuList
				set oMenuList = g_oWebSiteXML.SelectNodes("website/menus/menu[page/@id='" & g_sPageID & "']")
				if oMenuList.Length=0 then
					set oMenuList = g_oWebSiteXML.SelectNodes("website/menus/menu")					
				End If
				g_sMenuID	= oMenuList.item(0).attributes.GetNamedItem("id").Value
			End If
		End If		
		
		
		'-- Load the webpage data --
		if len(g_sPageID)>0 then
			
			Set g_oWebPageXML = CreateFreeDomDocument
			If not g_oWebPageXML.Load (webpage_xml) then
				LogIt "loader.dom.asp", "LoadDOMInMemory", FATAL, "Fail to load webpage data: " & webpage_xml, g_oWebPageXML.ParseError.Reason
				FatalError "This page does not exist", "Please press the back button of your browser."
			End if
			
			'-- Read values from webpage
			'g_sAutomenu	= g_oWebPageXML.documentElement.attributes.GetNamedItem("automenu").Value
			g_sTemplate = g_oWebPageXML.documentElement.attributes.GetNamedItem("template").Value
			g_sTheme	= g_oWebPageXML.documentElement.attributes.GetNamedItem("theme").Value
			
			'-- Read values from website 
			g_sSkin		= g_oWebSiteXML.documentElement.attributes.GetNamedItem("skin").Value
		End If
		
		
		'-- load the version info
		if instr(1, g_sScriptName, "admin.asp")>0 then
			Set g_oLocalVersion = CreateDomDocument
			If NOT g_oLocalVersion.load(version_path) Then
				LogIt "admin.updates.asp", "CheckUpdates", ERROR, g_oLocalVersion.ParseError.Reason, g_oLocalVersion.url
				Response.Write String("system", "fullxmlupdates", "noversion")
			End If	
		end if
	End Sub
	
	
	'-----------------------------
	'-- TODO :: save with FSO ! --
	'-----------------------------
	Public Function save_webpage
		on error resume next
		g_oWebPageXML.save webpage_xml
		if err<>0 then
			err.clear
			Logit "global.asp", "save_webpage", ERROR, err.number, err.description
		end if
		on error goto 0
	End Function
	
	
	'---------------------------------------------------------------------------
	' Get a string 
	' Input:	
	'			ModuleName : name of the module
	'			SectionID : id of the section
	'			Identifier of the string (1 -> x)
	' Output:	
	'			The string or identifier if is not found
	'---------------------------------------------------------------------------
	Function String(p_sModuleName, p_sSectionID, p_sStringID)
		Dim oStrings
				
		'-- check if the modules are loaded
		LoadModulesInMemory false
		
		'-- "System" module by default
		IFF Len(p_sModuleName)=0, "system", p_sModuleName		
		
		'-- try to get the string in the requested language (g_oUser.Culture)
		Set oStrings = Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[@name='"&p_sModuleName&"']/culture[@code='" & trim(g_sCulture) & "']/section[@id='" & p_sSectionID  & "']/string[@id='" & p_sStringID & "']")
		
		'-- otherwise, get it in english
		If (oStrings.Length=0) Then
			Set oStrings = Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[@name='"&p_sModuleName&"']/culture[@code='en']/section[@id='" & p_sSectionID  & "']/string[@id='" & p_sStringID & "']")
		end if
		
		'-- return the value
		If (oStrings.Length>0) Then
			String = oStrings.Item(0).text
		Else
			String = "{[" + p_sModuleName + "], [" + p_sSectionID + "], [" & p_sStringID & "]}"
		End If		
	End Function
	
	'---------------------------------------------------------------------------
	' Get if a string exists
	' Input:	
	'			ModuleName : name of the module
	'			SectionID : id of the section
	'			Identifier of the string (1 -> x)
	' Output:	
	'			true when the string exists, false otherwise
	'---------------------------------------------------------------------------
	Function StringExists(p_sModuleName, p_sSectionID, p_sStringID)
		Dim oStrings
				
		'-- check if the modules are loaded
		LoadModulesInMemory false
		
		'-- "System" module by default
		IFF Len(p_sModuleName)=0, "system", p_sModuleName		
		
		'-- try to get the string in the requested language (g_oUser.Culture)
		Set oStrings = Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[@name='"&p_sModuleName&"']/culture[@code='" & trim(g_sCulture) & "']/section[@id='" & p_sSectionID  & "']/string[@id='" & p_sStringID & "']")
		
		'-- otherwise, get it in english
		If (oStrings.Length=0) Then
			Set oStrings = Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[@name='"&p_sModuleName&"']/culture[@code='en']/section[@id='" & p_sSectionID  & "']/string[@id='" & p_sStringID & "']")
		end if
		
		'-- return the value
		If (oStrings.Length>0) Then
			StringExists = true
		Else
			StringExists = false
		End If		
	End Function
	
	'------------------------------------------------
	'-- Display a listbox of all available culture --
	'------------------------------------------------
	Function HtmlComponent_Culture(sName, sValue)
		Dim oXML, oLanguage
		Set oXML = CreateDomDocument		
		if not oXML.Load(languages_path) then
			LogIt "cultures.asp", "HtmlComponent_Culture", ERROR, oXML.parseError.reason, oXML.url
			exit function
		end if
		
		Dim t 
		t = t & "<select name='" & sName & "'><option></option>"
				
		for each oLanguage in oXML.SelectNodes("languages/language")
			t = t & "<option value='" & getAttribute(oLanguage, "id", "") & "'" & iff(getAttribute(oLanguage, "id", "en-us") = sValue, " selected", "") & ">" & getAttribute(oLanguage, "name", "") & "</option>"
		next
		t = t & "</select>"
		
		HtmlComponent_Culture = t
	End Function
	
	
	'----------------------------------------------------
	'-- Display the console menu                       --
	'-- Which depend of the user group and permissions --
	'----------------------------------------------------
	Sub ConsoleMenu()
		
		Dim oTemplate
		Dim ModuleName, ToolName
		Dim oToolList, index
				
		'-- be sure to have modules loaded if not
		LoadModulesInMemory(false)
		Dim oModulesXML : set oModulesXML = Application(APPVAR_DOM_MODULES)
		
		
		'-- create the template object
		Set oTemplate = new AspTemplate
		oTemplate.TemplateDir = g_sServerMapPath & ADMIN_FOLDER & "templates/"
		oTemplate.Template = "menu.html"
		oTemplate.ClearBlock "MenuBlock"
		
		
		
		'-- Display the main menu block
		If g_oUser.Group = "administrator" or g_oUser.Group = "webmaster" then		
				
			'-- Fill the title
			oTemplate.Slot( "title") = String("system", "interface", "websitemanager")
			oTemplate.Slot( "subtitle") = String("system", "website", "website")
			oTemplate.Slot( "subtitleaction") = "webform_update_website"
						
			'-- Fill each link
			oTemplate.ClearBlock "PageBlock"
			
			'-- Menus
			oTemplate.Slot("image") = "engine/admin/media/T.png"
			oTemplate.Slot("menuimage") = "engine/admin/media/website.png"
			oTemplate.Slot("menuaction") = "webform_tree_website"
			oTemplate.Slot("menulabel") = String("system", "menus", "menus")
			oTemplate.RepeatBlock "PageBlock"
			
			'-- Special pages
			oTemplate.Slot( "image") = "engine/admin/media/T.png"
			oTemplate.Slot( "menuimage") = "engine/admin/media/specialpages.png"
			oTemplate.Slot( "menuaction") = "webform_list_special_pages"
			oTemplate.Slot( "menulabel") =String("system", "webpages", "specialpages")
			oTemplate.RepeatBlock "PageBlock"
			
			
			'-- tools of the modules that are not independant
			set oToolList = oModulesXML.SelectNodes("/modules/module[@independant='false' and @enabled='true']/tools/tool")
			For index=0 to oToolList.Length-1
				ModuleName = getAttribute(oToolList(index).parentNode.parentNode, "name","")
				ToolName = getAttribute(oToolList(index), "name", "")
								
				if index=oToolList.length-1 Then
					oTemplate.Slot("image") = "engine/admin/media/L.png"
				else
					oTemplate.Slot("image") = "engine/admin/media/T.png"
				end if
				
				oTemplate.Slot("menuimage") = appSettings("MODULES_FOLDER") & "/" & ModuleName & "/media/tool_" & ToolName & ".png"
				oTemplate.Slot("menuaction") = "webform_" & ModuleName & "_" & ToolName
				oTemplate.Slot("menulabel") = String(ModuleName, "tools", ToolName)
				oTemplate.RepeatBlock "PageBlock"
			Next
			
			'-- validate the last block			
			oTemplate.RepeatBlock "MenuBlock"
					
			
			'-- Create a menu box for each 'independant' module 
			Dim oModulesList, oModule
						
			Set oModulesList = oModulesXML.SelectNodes("/modules/module[@independant='true' and @enabled='true' and tools/tool]")			
			For Each oModule in oModulesList
				ModuleName = getAttribute(oModule, "name","")
												
				'-- Fill the title
				oTemplate.Slot("title") = String("system", "interface", "module") & " " & ModuleName
				oTemplate.Slot("subtitle") = ModuleName
				oTemplate.Slot("subtitleaction") = "webform_edit_module_" & ModuleName
			
			
				'-- Loop on each tool of the module				
				Set oToolList = oModule.SelectNodes("tools/tool")
				
				oTemplate.ClearBlock "PageBlock"
								
				'-- add a link for each tool (in the module menu box)
				For index=0 to oToolList.Length-1
					
					ToolName = getAttribute(oToolList(index), "name", "")
								
					If index=oToolList.length-1 Then
						oTemplate.Slot( "image") =  "engine/admin/media/L.png"
					Else
						oTemplate.Slot( "image") = "engine/admin/media/T.png"
					End If
					
					oTemplate.Slot( "menuimage") = appSettings("MODULES_FOLDER") & "/" & ModuleName & "/media/" & ToolName & ".png"
					oTemplate.Slot( "menuaction") = "webform_" & ModuleName & "_" & ToolName
					oTemplate.Slot( "menulabel") = String(ModuleName, "tools", ToolName)
					oTemplate.RepeatBlock "PageBlock"
				Next
				
				
				'-- validate the last block			
				oTemplate.RepeatBlock "MenuBlock"	
			Next
			
		End If
		
		
		'-- ADMINISTRATOR panel
		if g_oUser.Group = "administrator" then
			
			'-- Fill the title
			oTemplate.Slot("title") = String("system", "interface", "controlpanel")
			oTemplate.Slot("subtitle") = String("system", "interface", "settings")
			oTemplate.Slot("subtitleaction") = "webform_update_website"
			
			'-- Fill each link
			oTemplate.ClearBlock "PageBlock"
			
			'-- 1st link :: Groups
				oTemplate.Slot("image") =  "engine/admin/media/T.png"
				oTemplate.Slot("menuimage") = "engine/admin/media/groups.png"
				oTemplate.Slot("menuaction") = "webform_list_groups"
				oTemplate.Slot("menulabel") = String("system", "groups", "title")
				oTemplate.RepeatBlock "PageBlock"
						
			'-- 2st link :: Users
				oTemplate.Slot("image") =  "engine/admin/media/T.png"
				oTemplate.Slot("menuimage") = "engine/admin/media/users.png"
				oTemplate.Slot("menuaction") = "webform_list_users"
				oTemplate.Slot("menulabel") = String("system", "users", "title")
				oTemplate.RepeatBlock "PageBlock"
				
			'-- 3nd link :: Modules
				oTemplate.Slot("image") = "engine/admin/media/T.png"
				oTemplate.Slot("menuimage") = "engine/admin/media/module.png"
				oTemplate.Slot("menuaction") = "webform_list_modules"
				oTemplate.Slot("menulabel") = String("system", "modules", "modules")
				oTemplate.RepeatBlock "PageBlock"
				
			'-- 4nd link :: Skins
				oTemplate.Slot("image") = "engine/admin/media/T.png"
				oTemplate.Slot("menuimage") = "engine/admin/media/skin.png"
				oTemplate.Slot("menuaction") = "webform_list_skins"
				oTemplate.Slot("menulabel") = String("system", "skins", "skins")
				oTemplate.RepeatBlock "PageBlock"
				
			'-- 4nd link :: Updates
				oTemplate.Slot("image") = "engine/admin/media/L.png"
				oTemplate.Slot("menuimage") = "engine/admin/media/update.png"
				oTemplate.Slot("menuaction") = "webform_fullxmlupdate"
				oTemplate.Slot("menulabel") = String("system", "fullxmlupdates",  "fullxmlupdate")
				oTemplate.RepeatBlock "PageBlock"
							
			'-- validate the last block			
			oTemplate.RepeatBlock "MenuBlock"
		End If
		
		
		'----------------
		'-- STATS menu --
		'----------------
		If g_oUser.Group = "administrator" or g_oUser.Group = "webmaster" then		
					
			'-- Fill the title
			oTemplate.Slot("title") = String("system", "webtraffic", "webtraffic")
			oTemplate.Slot("subtitle") = String("system", "interface", "settings")
			oTemplate.Slot("subtitleaction") = "webform_webtraffic_settings"
			
			
			'-- Fill each link
			oTemplate.ClearBlock "PageBlock"
			
			'-- Summary
			oTemplate.Slot("image") = "engine/admin/media/T.png"
			oTemplate.Slot("menuimage") = "engine/admin/media/statsummary.png"
			oTemplate.Slot("menuaction") = "webform_webtraffic_summary"
			oTemplate.Slot("menulabel") = String("system", "webtraffic", "summary")
			oTemplate.RepeatBlock "PageBlock"
			
			'-- Pages
			oTemplate.Slot("image") =  "engine/admin/media/T.png"
			oTemplate.Slot("menuimage") = "engine/admin/media/page.png"
			oTemplate.Slot("menuaction") = "webform_webtraffic_pages" & iff(len(request.QueryString("date"))>0, "&date="&request.QueryString("date"), "")
			oTemplate.Slot("menulabel") = String("system", "webtraffic", "pages")
			oTemplate.RepeatBlock "PageBlock"
			
			'-- Users
			oTemplate.Slot("image") = "engine/admin/media/T.png"
			oTemplate.Slot("menuimage") = "engine/admin/media/users.png"
			oTemplate.Slot("menuaction") = "webform_webtraffic_users"  & iff(len(request.QueryString("date"))>0, "&date="&request.QueryString("date"), "")
			oTemplate.Slot("menulabel") = String("system", "webtraffic", "users")
			oTemplate.RepeatBlock "PageBlock"
			
			'-- Languages
			oTemplate.Slot("image") = "engine/admin/media/T.png"
			oTemplate.Slot("menuimage") = "engine/admin/media/languages.png"
			oTemplate.Slot("menuaction") = "webform_webtraffic_languages" & iff(len(request.QueryString("date"))>0, "&date="&request.QueryString("date"), "")
			oTemplate.Slot("menulabel") = String("system", "webtraffic", "languages")
			oTemplate.RepeatBlock "PageBlock"
			
			'-- Browsers
			oTemplate.Slot("image") = "engine/admin/media/L.png"
			oTemplate.Slot("menuimage") = "engine/admin/media/browsers.png"
			oTemplate.Slot("menuaction") = "webform_webtraffic_browsers" & iff(len(request.QueryString("date"))>0, "&date="&request.QueryString("date"), "")
			oTemplate.Slot("menulabel") = String("system", "webtraffic", "browsers")
			oTemplate.RepeatBlock "PageBlock"
			
			'-- validate the last block			
			oTemplate.RepeatBlock "MenuBlock"
		End If
		
		
		oTemplate.Generate
		
		Set oTemplate = Nothing
	End Sub
	
	
	
	
	
	'---------------
	'-- Log visit --
	'---------------
	Sub LogVisit(p_exec)
		on error resume next
				
		Dim obrowser
		Set obrowser = New CBrowser
		
		'-- log the error into a file
		Dim oLog
		Set oLog = New LogFile
			oLog.FieldSeparator = GetSeparator
			oLog.Columns = array("date", "pageid", "username", "ipaddress", "script_name", "querystring", "browser", "language", "exec")
			oLog.TemplateFileName = DATA_FOLDER & STATS_FOLDER & "%y-%m-%d.csv"
			Call oLog.Log(array(g_sPageID, g_oUser.Login, chr(34) & Request.ServerVariables("REMOTE_ADDR") & chr(34), g_sScriptName, Request.QueryString, obrowser.browser & " "& obrowser.version, trim(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") ), p_exec))
		Set oLog = nothing
		
		'-- clear and remove the error trapping
		if err<>0 then
			LogIt "Utilities.asp", "LogVisit", ERROR, err.number, err.Description
			err.Clear
		end if
		
		Set obrowser = Nothing
		
		on error goto 0		
	End Sub
	
	
	Sub Debug(sText)
		if g_oUser.Group = "administrator" and g_oUser.ScreenName="debuger" then
			response.write "<span class=debug>" & sText & "</span><br>"
		end if
	End Sub
%>