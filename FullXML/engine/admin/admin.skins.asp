<%	'=====================================================================================================
	' THIS FILE CONTAINS ALL THE SKINS RELATED FUNCTIONS
	'=====================================================================================================
	
	
	'<summary>
	'	<name>Skins</name>
	'	<description>Display the skins list from the xml file</description>
	'	<inputs></inputs>
	'	<return></return>
	'	<remark></remark>
	'</summary>
	Sub webform_list_skins
		Call XmlDatagrid("skins", skins_xml, "/skins/skin", Array(String("system", "common", "name")), Array("id"), "", "id", "id", false)
		Response.Write "<br><center><input type=button onclick=""document.location='" & g_sScriptName & "?afterprocesswebform=webform_list_skins&action=do_refresh_skins" & "';"" value='" & String("system", "skins", "autodetect") & "'></center>"
	End Sub
	
		
	'--	Create the file and redirect
	Function do_refresh_skins
		
		'-- create the skins.xml definition file
		CreateSkinsXMLFile
				
		LoadSkinsInMemory true
				
	End Function
	
	
	'-- Build the skin xml file from the File System
	Sub CreateSkinsXMLFile
		Dim oSkinFolder, oFolder
		Dim oRoot, oSkin
		Dim oXML, oNodeList, oNode, att, pi
		
		'-- Objects
		Set oSkinFolder = g_oFSO.GetFolder(SKINS_FOLDER)
		Set oXML = CreateDomDocument
				
		
		'-- Add the document element
		set oRoot = oXML.createElement("skins")
		
		'-- Loop on each skin folder
		For each oFolder in oSkinFolder.SubFolders
			
			if LCase(oFolder.name)<>"cvs" then
				Response.Write oFolder.name & "<br>"
				
				Set oSkin = oXML.CreateElement("skin")
				Set att = oXML.createAttribute("id")		: att.value = oFolder.Name : oSkin.Attributes.SetNamedItem(att)
				
				'-- get the templates
				Dim oTemplateFile, oTemplateNode
				For each oTemplateFile in g_oFSO.GetFolder(SKINS_FOLDER & "\" & oFolder.Name & "\templates").Files
					set oTemplateNode = oXML.CreateElement("template")
					Set att = oXML.createAttribute("id")	: att.value = oTemplateFile.Name : oTemplateNode.Attributes.SetNamedItem(att)
					Set att = oXML.createAttribute("name")	: att.value = RemoveExtension(oTemplateFile.Name) : oTemplateNode.Attributes.SetNamedItem(att)
					oSkin.AppendChild(oTemplateNode)
				Next
				
				'-- get the themes
				Dim oThemeFolder, oThemeNode
				For each oThemeFolder in g_oFSO.GetFolder(SKINS_FOLDER & "\" & oFolder.Name & "\themes").SubFolders
					
					if LCase(oThemeFolder.name)<>"cvs" then
					
						set oThemeNode = oXML.CreateElement("theme")
						Set att = oXML.createAttribute("id")	: att.value = oThemeFolder.Name : oThemeNode.Attributes.SetNamedItem(att)
						
						'-- loop on each css of a theme
						Dim oCssFile, oCssNode
						For Each oCssFile in g_oFSO.GetFolder(oThemeFolder.path).Files
							set oCssNode = oXML.CreateElement("css")
							Set att = oXML.createAttribute("id")	: att.value = oCssFile.Name : oCssNode.Attributes.SetNamedItem(att)
							Set att = oXML.createAttribute("name")	: att.value = RemoveExtension(oCssFile.Name) : oCssNode.Attributes.SetNamedItem(att)
							oThemeNode.AppendChild(oCssNode)
						Next
						
						oSkin.AppendChild(oThemeNode)
					
					End If
				Next
				
				'-- get the boxes
				Dim oBoxFile, oBoxNode
				For each oBoxFile in g_oFSO.GetFolder(SKINS_FOLDER & "\" & oFolder.Name & "\boxes").Files
					set oBoxNode = oXML.CreateElement("box")
					Set att = oXML.createAttribute("id")	: att.value = oBoxFile.Name : oBoxNode.Attributes.SetNamedItem(att)
					Set att = oXML.createAttribute("name")	: att.value = RemoveExtension(oBoxFile.Name) : oBoxNode.Attributes.SetNamedItem(att)
					oSkin.AppendChild(oBoxNode)
				Next
							
				oRoot.appendChild(oSkin)
			
			End If
		Next
		
		oxml.appendChild(oRoot)
		oXML.Save skins_xml
		Set oXML = Nothing
		
	End Sub
	
	
'	'---------------------------------
'	'-- Display the insert/edit form
'	Sub webform_update_skin
'		
'		SkinProcess
'		
'		Dim skinID : skinID = Request.QueryString("id")
'		Dim process : process = "do_insert_skin"
'		Dim name, author, version, url, description
'		Dim oXML, oNodeList, oNode
'				
'		Set oXML = CreateDomDocument
'		if not oXML.Load (skins_xml) then
'			LogIt "admin.skins.asp", "EditSkin", ERROR, oXML.parseerror.reason, skins_xml
'		end if
'		
'		'-- If an id is passed, then we are editing the data, so load the old value
'		if len(skinID)>0 Then
'			process = "do_update_skin"
'			
'			Set oNodeList = oXML.SelectNodes("/skins/skin[@id='" & skinID & "']")	
'			
'			if oNodeList.length>0 then				
'				name = GetAttribute(oNodeList(0), "name", "")
'				author = GetAttribute(oNodeList(0), "author", "")
'				version = GetAttribute(oNodeList(0), "version", "")
'				url = GetAttribute(oNodeList(0), "url", "")
'				description = GetAttribute(oNodeList(0), "description", "")				
'			end if		
'		End If
'		
'				
'		With Response
'			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
'			.Write "<form action=" & g_sURL & " method=post id=frmEdit name=frmEdit>"
'			.Write "<input type=hidden name=process value='" & process & "'>"
'			.Write "<input type=hidden name=skinID value='" & skinID & "'>"
'			.Write "<caption>" & String("system", "skins", "skin") & "</caption>"
'			
'			.Write "<tr class=datagrid_editrow><th>" & String("system", "skins", "name") & "</th><td><input type=text class=large name=name value='" & name & "'></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & String("system", "skins", "author") & "</th><td><input type=text class=large name=author value=""" & author & """></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & String("system", "skins", "version") & "</th><td><input type=text class=small name=version value=""" & version & """></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & String("system", "skins", "url") & "</th><td><input type=text class=large name=url value=""" & url & """></td></tr>"
'			.Write "<tr class=datagrid_editrow><th>" & String("system", "skins", "description") & "</th><td><textarea name=description class=small>" & description & "</textarea></td></tr>"
'			
'			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "cancel") & """ onclick=""document.location='" & g_sScriptName & "?webform_list_skins';""></td></tr>"
'			if len(skinID) then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_skin';document.forms[0].submit();}""></td></tr>"
'			.Write "</form>"
'			.Write "</table>"			
'		End With
'		
'		Set oXML = Nothing
'		
'	End Sub
	
	
'	'-- process action
'	Sub do_skin_detection
'		
'		Select case Request.Form("process")
'			case "do_update_skin"
'				Call UpdateNode (skins_xml, "/skins/skin[@id='" & getParam("skinID") & "']", Array("name", "author", "version", "url", "description"), Array(getParam("name"), getParam("author"), getParam("version"), getParam("url"), getParam("description")))
'				Response.Redirect g_sScriptName & "?webform=webform_list_skins"
'			
'			case "do_insert_skin"
'				Call InsertNode (skins_xml, "/skins" , "skin", Array("name", "author", "version", "url", "description"), Array(getParam("name"), getParam("author"), getParam("version"), getParam("url"), getParam("description")), true, "")
'				Response.Redirect g_sScriptName & "?webform=webform_list_skins"
'			
'			case "do_delete_skin"
'				Call DeleteNode (skins_xml, "/skins/skin[@id='" & getParam("skinID") & "']")
'				Response.Redirect g_sScriptName & "?webform=webform_list_skins"
'		'
'		End Select
'	End Sub
	
	
	'-- Load each available modules ----------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Sub LoadSkinsInMemory(bForceUpdate)
		
		'Set Application(APPVAR_DOM_SKINS) = Nothing
		'Application(APPVAR_DOM_SKINS) = Empty
				
		If bForceUpdate OR (not isObject(Application(APPVAR_DOM_SKINS))) or Request.QueryString("force")="yes" Then
			Dim oXML
					
			'-- 
			LogIt "admin.skins.asp", "LoadSkinsInMemory", INFO, "Load skins", ""
			
			'-- clean old datas
			Set Application(APPVAR_DOM_SKINS) = Nothing
			Application(APPVAR_DOM_SKINS) = Empty
		
			Set oXML = CreateFreeDomDocument
			If NOT oXML.load(skins_xml) Then
				LogIt "cultures.asp", "LoadSkinsInMemory", ERROR, oXML.ParseError & " : " & oXML.ParseError.Reason, modules_xml
				
				if oXML.ParseError=-2146697210 then
					CreateSkinsXMLFile
					
					If NOT oXML.load(skins_xml) Then
						LogIt "admin.skins.asp", "LoadSkinsInMemory", FATAL, "Cant create the modules file.", modules_xml
						FatalError "Fatal Error", "A fatal error occured. [ref fx004 : Can't Create the modules file.]"
					End If
				end if
			
			End If	
			
			Set Application(APPVAR_DOM_SKINS) = oXML
		
		End If		
	End Sub

%>