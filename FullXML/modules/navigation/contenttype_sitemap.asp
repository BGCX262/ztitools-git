<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_navigation_SiteMap(oContent)
			
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_navigation_SiteMap(oNode)
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_navigation_SiteMap(oContent)
		Dim oTemplate, oMenu, oPage
		Set oTemplate = new AspTemplate
				
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\navigation\"
		oTemplate.Template = "sitemap.html"
		
		oTemplate.ClearBlock "menu_item_block"
		
		'-- Loop on each menu
		For Each oMenu in g_oWebSiteXML.selectNodes("website/menus/menu")
			oTemplate.Slot("menu_item_name") = getAttribute(oMenu, "name", "")
		
			oTemplate.ClearBlock "page_item_block"		
			
			'-- Loop on each page
			For each oPage in oMenu.SelectNodes("page")
				oTemplate.Slot("page_item_link") = g_sBaseUrl & "?pID="& getAttribute(oPage, "id", "")
				oTemplate.Slot("page_item_name") = getAttribute(oPage, "name", "")
				oTemplate.RepeatBlock "page_item_block"
			Next
			
			oTemplate.RepeatBlock "menu_item_block"
		Next

		Render_navigation_SiteMap = oTemplate.GetOutput
		Set oTemplate = Nothing
	End Function
%>