<%
'	DIM TOOLBAR_FILE : TOOLBAR_FILE = appSettings("TOOLBAR_FILE")
'	
'	'-- Give the path to the current toolbar data ----
'	Function toolbar_xml
'		toolbar_xml = DATA_FOLDER & TOOLBAR_FILE & XMLFILE_EXTENSION
'	End Function

	
	'-- Display the toolbar ------------------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Public Function ToolBar_Render()
		Dim oTemplate
		Set oTemplate = new AspTemplate
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\toolbar\"
		oTemplate.Template = "toolbar.html"
		
		'-- Load the toolbar.xml
		Dim oXML, oItem
		Set oXML = CreateDomDocument
		If not oXML.Load (toolbar_xml) then
			LogIt "contenttype_toolbar.asp", "ToolBar_Render", ERROR, oXML.ParseError.Reason, toolbar_xml
		End if
		
		
		oTemplate.ClearBlock "toolbar_item_block"
		
		'-- Loop on each item
		For Each oItem in oXML.SelectNodes("/toolbar/item")
			Dim linktype
			linktype = GetAttribute(oItem, "linktype", "internal")
			if linktype = "internal" Then
				Dim pageNode, pageNodeList
				Dim pageID
				pageID = GetAttribute(oItem, "pageID","")
				Set pageNodeList = g_oWebSiteXML.documentelement.SelectNodes("/website/menus/menu/page[@id='" & GetAttribute(oItem, "pageID","") & "']")
				
				'-- we check that the page is still exists
				if pageNodeList.length>0 then
					Set pageNode = pageNodeList.item(0)
					oTemplate.Slot("toolbar_item_link") = "default.asp?sID=" & g_iWebSiteID & "&mID=" & GetAttribute(pageNode.ParentNode, "id", "") & "&pID=" & GetAttribute(oItem, "pageID", "")
					oTemplate.Slot("toolbar_item_name") = GetAttribute(pageNode, "name", "")
					oTemplate.Slot("toolbar_item_target") = GetAttribute(oItem, "target", "_self")
					oTemplate.RepeatBlock "toolbar_item_block"
				end if
			
			elseif linktype = "external" Then
				oTemplate.Slot("toolbar_item_link") = GetAttribute(oItem, "link", "")
				oTemplate.Slot("toolbar_item_name") = GetAttribute(oItem, "name", "")
				oTemplate.Slot("toolbar_item_target") = GetAttribute(oItem, "target", "_self")					
				oTemplate.RepeatBlock "toolbar_item_block"
			
			elseif linktype = "login" Then
				oTemplate.Slot("toolbar_item_link") = "javascript: var auth=window.open('user.asp?action=DisplayAuthenticationForm','login','width=400, height=450')"
				oTemplate.Slot("toolbar_item_name") = String("system", "common", "opensession")
				oTemplate.Slot("toolbar_item_target") = GetAttribute(oItem, "target", "_self")					
				oTemplate.RepeatBlock "toolbar_item_block"
			end if
		Next
				
		ToolBar_Render = oTemplate.GetOutput
		
		Set oXML = Nothing
		Set oTemplate = Nothing
	End Function
%>