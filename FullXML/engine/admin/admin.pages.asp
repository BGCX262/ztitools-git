<%	
	'+-----------------------------------------------------------------
	'| DESCRIPTION: 
	'| List, add and edit the webpages
	'|
	'+------------------------------------------------------------------
	



	'------------------------
	'-- Special Pages list --
	'------------------------
	Public Sub webform_list_special_pages
		Call XmlDatagrid(String("system", "webpages", "specialpages"), g_oWebSiteXML, "/website/specialpages/page", Array(String("system", "webpages", "name"), String("system", "webpages", "template"), String("system", "webpages", "theme")), Array("name", "template", "theme"), "webform_update_page", "id", "pID", true)
	End Sub
	
	
	'------------------------------------------------
	'-- Display the pages list, filtered by menuID --
	'------------------------------------------------
	Sub webform_list_pages_bymenu
		
		'-- get the list of datas 
		Dim aNam, aVal, oPage 
		Set aNam = new Collection
		Set aVal = new Collection
			
		For each oPage in g_oWebSiteXML.SelectNodes("/website/menus/menu[@id='" & g_sMenuID & "']/page" & iff(len(request("pID"))>0, "[@id='"&g_sPageID&"']/page", ""))
			aNam.Add(oPage.attributes.GetNamedItem("name").value)
			aVal.Add(oPage.attributes.GetNamedItem("id").value)
		Next
		
			
		'-- now that we have the datas, we create the form	
		Dim o
		Set o = new CtrlListEditor
		o.Width = 570
		
		If len(request("pID"))>0 Then
			o.Title = String("system", "webpages", "subpages")
		Else
			o.Title = String("system", "webpages", "pages")
		End If
		o.Labels = aNam.ToArray()
		o.Values = aVal.ToArray()
		
		'-- buttons label
		o.AddButtonLabel = String("system", "common", "add")
		o.EditButtonLabel = String("system", "common", "edit")
		o.DeleteButtonLabel = String("system", "common", "delete")
		o.MoveUpButtonLabel = String("system", "common", "moveup")
		o.MoveDownButtonLabel = String("system", "common", "movedown")
		o.ConfirmDeleteWarning = String("system", "common", "confirmdelete")
		
		'asp functions that are called on buttons
		o.EditUrl = g_sScriptName & "?webform=webform_update_page&mID=" & g_sMenuID & "&pID="
		o.AddUrl = g_sScriptName & "?webform=webform_insert_page&mID=" & g_sMenuID & "&parentpID=" & g_sPageID & "&beforepageID=" 
		o.DeleteUrl = g_sScriptName & "?afterprocesswebform=" & server.URLEncode("webform_update_menu&mID=" & g_sMenuID) & "&mID=" & g_sMenuID & "&process=do_delete_page&pID="
		o.MoveDownUrl = g_sScriptName & "?afterprocesswebform=" & server.URLEncode("webform_update_menu&mID=" & g_sMenuID) & "&mID=" & g_sMenuID & "&process=do_movedown_page&pID="
		o.MoveUPUrl = g_sScriptName & "?afterprocesswebform=" & server.URLEncode("webform_update_menu&mID=" & g_sMenuID) & "&mID=" & g_sMenuID & "&process=do_moveup_page&pID="
		
		o.Display	
		
		Set o = Nothing
	End Sub
	
	sub webform_insert_page
		Call private_webform_edit_page()
	end sub
	
	sub webform_update_page
		Call private_webform_edit_page()
	end sub
	
		
	'--------------------------------------------------
	'-- Display the insert/edit form for the webpage --
	'--------------------------------------------------
	Private Sub private_webform_edit_page
		Dim usenewpagetemplate : usenewpagetemplate = USE_NEWPAGE_TEMPLATE
		Dim process : process = "do_insert_page"
		Dim oNodeList, oNode, skin	
		Dim name, template, theme, metatitle, metadescription, metakeywords, published
		
		'-- default value for insertion
		skin = GetAttribute(g_oWebSiteXML.DocumentElement, "skin", "")
		published = appSettings("DEFAULT_PUBLICATION_STATE")
		
		'--TODO: Get that from the newpage template
		
		
		'-- in the case of an update, load old value
		if len(g_sPageID)>0 Then
			process = "do_update_page"
						
			'-- get the datas to fill the form
			Set oNodeList = g_oWebPageXML.SelectNodes("/page")
			
			If oNodeList.length>0 then				
				name = GetAttribute(oNodeList(0), "name", "")
				template = GetAttribute(oNodeList(0), "template", "")
				theme = GetAttribute(oNodeList(0), "theme", "")
				published = GetAttribute(oNodeList(0), "published", appSettings("DEFAULT_PUBLICATION_STATE"))
				
				'-- metas
				Dim oMETANode
				Set oMETANode = g_oWebPageXML.SelectNodes("/page/metas")
				if oMETANode.Length>0 then
					metatitle = GetChild(oMETANode.item(0), "title", "")
					metadescription = GetChild(oMETANode.item(0), "description", "")
					metakeywords = GetChild(oMETANode.item(0), "keywords", "")					
				end if
			End if
		
		'-- read info into the newpage template
		ElseIF USE_NEWPAGE_TEMPLATE Then
				Dim oNewPageXml	
				Set oNewPageXml = CreateDomDocument		
				If Not oNewPageXml.Load(DATA_FOLDER & PAGES_FOLDER & "newpage" & XMLFILE_EXTENSION) Then
					LogIt "admin.pages.asp", "Do_Insert_Page", ERROR, oXML.parseerror.reason, DATA_FOLDER & PAGES_FOLDER & "newpage" & XMLFILE_EXTENSION
					Response.Write "Can't load newpage.xml"
					Exit sub 
				End if
				
				name = getAttribute(oNewPageXml.documentElement, "name", "")
				template = getAttribute(oNewPageXml.documentElement, "template", "")
				theme = getAttribute(oNewPageXml.documentElement, "theme", "")
				
				Dim oMETAs : Set oMETAs = oNewPageXml.SelectNodes("/page/metas")
				if oMETAs.Length>0 then
					metatitle = GetChild(oMETAs.item(0), "title", "")
					metadescription = GetChild(oMETAs.item(0), "description", "")
					metakeywords = GetChild(oMETAs.item(0), "keywords", "")					
				end if
				
				Set oNewPageXml = Nothing
		Else
			template = "simple.html"
			theme = "classic"
		End If
		
		'-- the  naviguation for: settings / contents / permissions
		If Len(g_sPageID)>0 Then
			Response.Write "<div class=tabs><table class=tabs cellspacing=1 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_page&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page", " disabled","") & ">" & String("system", "webpages", "pagesettings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_update_page_permissions&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page_permissions", " disabled","") & ">" & String("system", "webpages", "pagepermissions") & "</a></td><td bgcolor=#FFEAFF><a href=admin.asp?webform=webform_update_page_contents&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page_contents", " disabled","") & ">" & String("system", "webpages", "pagecontents") & "</a></td></tr></table></div>"
		End If
		
		'- Print the form
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			
			if len(g_sMenuID)>0 then
				.Write "<input type=hidden name=afterprocesswebform value='webform_update_menu&mID=" & g_sMenuID& "'>"
			Else
				.Write "<input type=hidden name=afterprocesswebform value='webform_list_special_pages'>"
			End If
			
			.Write "<caption>" & String("system", "webpages", "webpage") & " - " & String("system", "webpages", "pagesettings") & "</caption>"
			
			
			If Len(g_sPageID)=0 Then
				.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "pageid") & "</th><td><input type=text class=medium name=pageID></td></tr>"
			End If
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "name") & "</th><td><input type=text class=large name=name value=""" & name & """></td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "template") & "</th><td>" & XMLListBox("template", "id", "name", skins_xml , "skins/skin[@id='" & skin & "']/template", template, null, null) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "theme") & "</th><td>" & XMLListBox("theme", "id", "id", skins_xml , "skins/skin[@id='" & skin & "']/theme", theme, null, null) & "</td></tr>"
								
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "metatitle") & "</th><td><input type=text class=large name=metatitle value=""" & metatitle & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "metadescription") & "</th><td><input type=text class=large name=metadescription value=""" & metadescription & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "metakeywords") & "</th><td><input type=text class=large name=metakeywords value=""" & metakeywords & """></td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "published") & "</th><td>" & HtmlComponent_PublicationState("frmEdit", "published", published) & "</td></tr>"
			
			
			If USE_NEWPAGE_TEMPLATE Then			
				.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "webpages", "usenewpagetemplate") & "</th><td>" & HtmlComponent_Bool("frmEdit", "usenewpagetemplate", usenewpagetemplate) & "</td></tr>"
			End If
						
			'-- ok / cancel buttons	
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.go(-1);""></td></tr>"
			
			'-- delete button
			if len(g_sPageID) then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value=""" & String("system", "common", "browse") & """ onclick=""document.location='default.asp?pID=" & g_sPageID & "';"">&nbsp;<input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_page';document.forms[0].submit();}""></td></tr>"
		
			.Write "</form>"
			.Write "</table><br/>"
			
			'-- display the link of this page
			if len(g_sPageID)>0 then 
				.Write String("system", "webpages", "link")& ": <input style='border: 1px #808080 solid; width: 500px;' disabled type=text value='default.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "' size=100><br><br>"
				
				'-- display subpages list
				Call webform_list_pages_bymenu()
			End If
		End With		
	End Sub
	
	
	'------------------------------------
	'-- Show the page permissions form --
	'------------------------------------
	Sub webform_update_page_permissions
				
		'-- the  naviguation for: settings / contents / permissions
		Response.Write "<div class=tabs><table class=tabs cellspacing=1 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_page&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page", " disabled","") & ">" & String("system", "webpages", "pagesettings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_update_page_permissions&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page_permissions", " disabled","") & ">" & String("system", "webpages", "pagepermissions") & "</a></td><td bgcolor=#FFEAFF><a href=admin.asp?webform=webform_update_page_contents&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page_contents", " disabled","") & ">" & String("system", "webpages", "pagecontents") & "</a></td></tr></table></div>"
		Call webform_edit_ObjectPermissions_groups("/website/menus/menu[@id='" & g_sMenuID  & "']//page[@id='" & g_sPageID & "']")
		
	End sub
	
	
	'----------------------------------------------------
	'-- Display the list of content that are in a page --
	'----------------------------------------------------
	Sub webform_update_page_contents
		
		'-- the  naviguation for: settings / contents / permissions
		Response.Write "<div class=tabs><table class=tabs cellspacing=1 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_page&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page", " disabled","") & ">" & String("system", "webpages", "pagesettings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_update_page_permissions&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page_permissions", " disabled","") & ">" & String("system", "webpages", "pagepermissions") & "</a></td><td bgcolor=#FFEAFF><a href=admin.asp?webform=webform_update_page_contents&pID=" & g_sPageID & "&mID=" & g_sMenuID & iff(g_sWebform="webform_update_page_contents", " disabled","") & ">" & String("system", "webpages", "pagecontents") & "</a></td></tr></table></div>"		
		Dim skin, template, oNodeList
		
		skin = GetAttribute(g_oWebSiteXML.DocumentElement, "skin", "")
		Set oNodeList = g_oWebPageXML.SelectNodes("/page")
			
		If oNodeList.length>0 then				
			template = GetAttribute(oNodeList(0), "template", "")
		End If
				
				
		'-- template file exists ?
		Dim templatePath : templatePath = SKINS_FOLDER & skin & "\templates\" & template
		if not g_oFSO.FileExists(templatePath) then
			Response.Write "The template file does not exist."
			exit sub
		end if
		
		'-- load content of the file
		Dim oTemplateFile, skeleton			
		Set oTemplateFile = g_oFso.OpenTextFile(templatePath)
		skeleton = oTemplateFile.readall
		oTemplateFile.Close
		
		
		
		'-- Loop on each placeholder founded in the skeleton			
		Dim g_oRegExp
		Set g_oRegExp = New RegExp
		g_oRegExp.IgnoreCase = True
		g_oRegExp.Global = True
		' not necessary :: g_oRegExp.Multiline = True
	
		Dim objMatches, match, submatch, placeholder	
		Dim oXMLPlaceHolder : Set oXMLPlaceHolder = CreateDomDocument
		g_oRegExp.Pattern = "(<placeholder[^\/>*].*/>)"
		Set objMatches = g_oRegExp.Execute(skeleton)
		Response.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
		Response.Write "<caption>" & String("system", "webpages", "webpage") & " - " & String("system", "webpages", "pagecontents") & "</caption>"
		
		Response.Write "<tr class=datagrid_column><th>&nbsp;</th><th>" & String("system", "placeholders", "placeholder") & "</th><th>" & String("system", "contents", "contents") & "</th><th>" & String("system", "placeholders", "paging") & "</th><th>" & String("system", "placeholders", "pagesize") & "</th></tr>"
		For each Match in objMatches
			For each submatch in Match.SubMatches		
				if oXMLPlaceHolder.LoadXML(submatch) Then
					placeholder = getAttribute(oXMLPlaceHolder.documentElement, "id", "")
					Dim nodeList, paging, pagesize
					paging = false
					pagesize = 10
					Set nodeList = g_oWebPageXML.documentElement.selectNodes("//placeholder[@id='" & placeholder & "']")
					if nodeList.Length>0 then
						paging = getAttribute(nodeList.item(0), "paging", false)
						pagesize = getAttribute(nodeList.item(0), "pagesize", 10)
					end if
					Response.Write "<tr class=datagrid_row><th>&nbsp;</th><td><a href=""admin.asp?webform=webform_update_placeholder&mID="& g_sMenuID & "&pID=" & g_sPageID & "&placeholder=" & placeholder & """>" & placeholder & "</a></td><td>" & oNodeList(0).SelectNodes("/page/placeholders/placeholder[@id='" & placeholder & "']/content").Length & "</td><td>" & paging  & "</td><td>" & pagesize  & "</td></tr>"
				end if					
			Next
		Next
		Response.Write "</table>"
	End Sub
	
	
	'+-----------------------------+
	'| Create a new webpage
	'+-----------------------------+
	Sub Do_Insert_Page()
		Dim parentpageID : parentpageID = getParam("parentpID")
		
		Dim arrayAttributes, arrayValues
		Dim beforeXpath : beforeXpath = Request.QueryString("beforepageID")
		
		'-- get the correct pageID
		g_sPageID = iff(len(getParam("pageID"))>0, getParam("pageID"), GetGuid())
		
		'-- 1. Update the website.xml, so we verify that pageID is unique
		if XPathChecker(website_xml, "/website//page[@id='" & g_sPageID & "']")=0 then
							
			arrayAttributes = array("id", "name", "template", "theme", "published")
			arrayValues = array(g_sPageID, getParam("name"), getParam("template"), getParam("theme"), getParam("published"))
			
			'-- insert the new webpage in the website.xml
			If len(g_sMenuID)=0 Then
				'-- special page
				Call InsertNode (website_xml, "/website/specialpages" , "page", arrayAttributes, arrayValues, false, "")
			
			ElseIf len(beforeXpath)>0 Then 
				Call InsertNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID &"']" & iff(len(parentpageID)>0, "//page[@id='"&parentpageID&"']" ,""), "page", arrayAttributes, arrayValues, false, "//page[@id='" & beforeXpath & "']")
			Else
				Call InsertNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID &"']" & iff(len(parentpageID)>0, "//page[@id='"&parentpageID&"']" ,""), "page", arrayAttributes, arrayValues, false, "")
			End If
			
		Else
			Response.Write "<script>alert('" & String("system", "webpages", "pageiderror") & "'); history.back();</script>"
			Exit sub
		End If
		
				
		'-- create the webpage file
		Dim oXML, oPlaceholdersTemplate
		Set oXML = CreateDomDocument
		
		
		'-- Get the placeholder templates
		If cbool(getParam("usenewpagetemplate")) then			
			
			If Not oXML.Load(DATA_FOLDER & PAGES_FOLDER & "newpage" & XMLFILE_EXTENSION) Then
				LogIt "admin.pages.asp", "Do_Insert_Page", ERROR, "Can't load the newpage template. Creating blank page", oXML.ParseError.Reason
				Response.Write "Can't load newpage.xml. Using blank template"				
			Else
				Set oPlaceholdersTemplate = oXML.SelectSingleNode("page/placeholders").CloneNode(true)
			End If
					
		End If
			
		
		'-- create the dom and add the processing instruction
		Dim oNodeList, oNode, att
		
		'-- add the document element
		Call oXML.LoadXML("<page/>")
		Set oNode = oXML.selectSingleNode("page")
				
		'-- add each webpage attributes		
		Dim index
		For index=LBound(arrayAttributes) to UBound(arrayAttributes)
			set att = oXML.CreateAttribute(arrayAttributes(index))
			att.value = CStr(arrayValues(index))
			oNode.Attributes.SetNamedItem(att)
		Next
			
			
		'-- add the metas
		Dim oMetas, oMeta
		Set oMetas = oXML.createElement("metas")
		Set oMeta = oXML.createElement("title") : oMeta.text = getParam("metatitle") : oMetas.appendChild(oMeta.CloneNode(true))
		Set oMeta = oXML.createElement("description") : oMeta.text = getParam("metadescription") : oMetas.appendChild(oMeta.CloneNode(true))
		Set oMeta = oXML.createElement("keywords") : oMeta.text = getParam("metakeywords") : oMetas.appendChild(oMeta.CloneNode(true))
		oNode.appendCHild(oMetas)
		
		'-- Add a "placeholders" child
		Dim oPlaceholders
		Set oPlaceholders = oXML.CreateElement("placeholders")
			
		'-- Append the sample placeholders
		If cbool(getParam("usenewpagetemplate")) Then
			Dim item
			For each item in oPlaceholdersTemplate.SelectNodes("placeholder")
				Call oPlaceholders.appendChild(item.cloneNode(true))
			Next			
		End If
				
		'-- append it to the page node
		oNode.appendChild(oPlaceHolders)
				
		'-- save
		oXML.Save webpage_xml			
		Set oXML = Nothing
				
	End Sub
	
	
	'+-----------------------------+
	'| Do updates on the page file
	'+-----------------------------+
	Sub Do_Update_Page
		Dim oXML, oNodeList, oNode, att
		Dim arrayAttributes : arrayAttributes = array("name", "template", "theme", "published")
		Dim arrayValues : arrayValues = array(getParam("name"), getParam("template"), getParam("theme"), getParam("published"))
		
		'-- update the page node in the webtree
		Call UpdateNode (website_xml, "//page[@id='" & g_sPageID & "']", arrayAttributes, arrayValues)
				
		'-- Update each attributes of the website
		Dim index
		For index=LBound(arrayAttributes) to UBound(arrayAttributes)
			on error resume next
			g_oWebPageXML.DocumentElement.Attributes.GetNamedItem(arrayAttributes(index)).Value = CStr(getParam(arrayAttributes(index)))
			if err<>0 then
				err.Clear
				on error goto 0
				Set att = g_oWebPageXML.CreateAttribute(arrayAttributes(index))
				att.value = CStr(getParam(arrayAttributes(index)))
				g_oWebPageXML.DocumentElement.Attributes.SetNamedItem(att)
			end if
			on error goto 0
		Next
		
		
		'-- update METAS
		g_oWebPageXML.SelectSingleNode("/page/metas/title").text =  getParam("metatitle")
		g_oWebPageXML.SelectSingleNode("/page/metas/description").text =  getParam("metadescription")
		g_oWebPageXML.SelectSingleNode("/page/metas/keywords").text =  getParam("metakeywords")
		
		'-- save the webpage
		save_webpage
		
	End Sub
	
	
	'-- Deletes of a page
	Sub Do_Delete_Page()
		
		Call DeleteNode (website_xml, "/website//page[@id='" & g_sPageID & "']")
		Call DeleteFile (webpage_xml)	
		
		
	End Sub
	
	
	'-- Move Up a page
	Sub Do_MoveUp_Page
		Call MoveUpNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID & "']/page[@id='" & g_sPageID & "']")
	End Sub
	
	
	'-- Move Up a page
	Sub Do_MoveDown_Page
		Call MoveDownNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID & "']/page[@id='" & g_sPageID & "']")
	End Sub
	
	
	'-- Loop on each placeholder founded in the skeleton			
	'-- End give back a selectbox
	Function HtmlComponent_PlaceholderSelect(sName, sValue)
		Dim t
		
		Dim skin, template, oNodeList
		
		skin = GetAttribute(g_oWebSiteXML.DocumentElement, "skin", "")
		Set oNodeList = g_oWebPageXML.SelectNodes("/page")
			
		If oNodeList.length>0 then				
			template = GetAttribute(oNodeList(0), "template", "")
		End If
				
		If len(sValue)=0 then
		   sValue = "main"
		end if
				
		'-- template file exists ?
		Dim templatePath : templatePath = SKINS_FOLDER & skin & "\templates\" & template
		if not g_oFSO.FileExists(templatePath) then
			Response.Write "The template file does not exist."
			exit function
		end if
		
		'-- load content of the file
		Dim oTemplateFile, skeleton			
		Set oTemplateFile = g_oFso.OpenTextFile(templatePath)
		skeleton = oTemplateFile.readall
		oTemplateFile.Close
				
		
		Dim g_oRegExp
		Set g_oRegExp = New RegExp
		g_oRegExp.IgnoreCase = True
		g_oRegExp.Global = True
		
		Dim objMatches, match, submatch, placeholder
		Dim oXMLPlaceHolder : Set oXMLPlaceHolder = CreateDomDocument
		g_oRegExp.Pattern = "(<placeholder[^\/>*].*/>)"
		Set objMatches = g_oRegExp.Execute(skeleton)
		
		t = t & "<select name='"&sName&"'>"
		
		For each Match in objMatches
			For each submatch in Match.SubMatches		
				if oXMLPlaceHolder.LoadXML(submatch) Then
					placeholder = getAttribute(oXMLPlaceHolder.documentElement, "id", "")
					t = t & "<option value='" & placeholder & "'" & iff(placeholder=sValue, " selected", "")&">"& placeholder &"</option>"
				end if					
			Next
		Next
		t = t & "</select>"
		HtmlComponent_PlaceholderSelect = t
	End Function
	
	
	'-- Display a select box for pages -------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Function HtmlComponent_SelectPage(sName, sValue)
		Dim oPage
			
		HtmlComponent_SelectPage = HtmlComponent_SelectPage & "<select name='" & sName & "'>"
		
		For each oPage in g_oWebSiteXML.SelectNodes("website/menus/menu/page")
			HtmlComponent_SelectPage = HtmlComponent_SelectPage & "<option value='" & oPage.Attributes.GetNamedItem("id").Value & "'"
			if sValue=oPage.Attributes.GetNamedItem("id").Value then
				HtmlComponent_SelectPage = HtmlComponent_SelectPage & " selected"
			end if
			HtmlComponent_SelectPage = HtmlComponent_SelectPage & ">"
			HtmlComponent_SelectPage = HtmlComponent_SelectPage & oPage.ParentNode.Attributes.GetNamedItem("name").Value & " --> " & oPage.Attributes.GetNamedItem("name").Value
			HtmlComponent_SelectPage = HtmlComponent_SelectPage & "</option>"
		Next
		
		HtmlComponent_SelectPage = HtmlComponent_SelectPage & "</select>"
		
	End Function
%>