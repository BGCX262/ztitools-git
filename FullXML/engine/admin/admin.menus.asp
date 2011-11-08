<%
	'=====================================================================================================
	' MENU RELATED FUNCTIONS
	'=====================================================================================================
	
	
	'-- menus list
	Sub webform_list_menus
		'-- get the list of datas
		Dim aNam, aVal, oMenu 
		Set aNam = new Collection
		Set aVal = new Collection
		For each oMenu in g_oWebSiteXML.SelectNodes("/website/menus/menu")
			aNam.Add(oMenu.attributes.GetNamedItem("name").value)
			aVal.Add(oMenu.attributes.GetNamedItem("id").value)
		Next
				
		Dim o
		Set o = new CtrlListEditor
		
		o.Title = String("system", "menus", "menus")
		o.Labels = aNam.ToArray() 
		o.Values = aVal.ToArray()
		
		'-- buttons label
		o.AddButtonLabel = String("system", "common", "add")
		o.EditButtonLabel = String("system", "common", "edit")
		o.DeleteButtonLabel = String("system", "common", "delete")
		o.MoveUpButtonLabel = String("system", "common", "moveup")
		o.MoveDownButtonLabel = String("system", "common", "movedown")
		o.ConfirmDeleteWarning = String("system", "common", "confirmdelete")
		
		'URL that are called on buttons
		o.EditUrl = g_sScriptName & "?webform=webform_update_menu&mID="
		o.AddUrl = g_sScriptName & "?webform=webform_insert_menu&beforemenuID="
		o.DeleteUrl = g_sScriptName & "?AfterProcessWebform=webform_list_menus&process=do_delete_menu&mID="
		o.MoveUPUrl = g_sScriptName & "?AfterProcessWebform=webform_list_menus&process=do_moveup_menu&mID="
		o.MoveDownUrl = g_sScriptName & "?AfterProcessWebform=webform_list_menus&process=do_movedown_menu&mID="
				
		o.Display
		
		Set o = Nothing
	End Sub
	
	
	Sub webform_insert_menu
		Call private_webform_menu()
	End Sub
	
	
	Sub webform_update_menu
		Call private_webform_menu()
	End Sub
	
	'+-------------------------------------------+
	'| Display the insert/edit form for the menu
	'+-------------------------------------------+
	Sub private_webform_menu
		Dim process : process = "do_insert_menu"
		Dim name, published
				
		'-- If an id is passed, then we are editing the data, so load the old value
		if len(g_sMenuID)>0 Then
			process = "do_update_menu"			
			
			Dim oNodeList, oNode
					
			Set oNodeList = g_oWebSiteXML.SelectNodes("//menu[@id='" & g_sMenuID & "']")			
			If oNodeList.length>0 then				
				name		= GetAttribute(oNodeList(0), "name", "")
				published	= GetAttribute(oNodeList(0), "published", appSettings("DEFAULT_PUBLICATION_STATE"))
			End if
		else
			published = appSettings("DEFAULT_PUBLICATION_STATE")
		End If		
		
		'-- Website settings, horizontal navigation
		If Len(g_sMenuID)>0 Then
			Response.Write "<div class=tabs><table class=tabs cellspacing=0 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_menu&mID=" & g_sMenuID & iff(g_sWebform="webform_update_menu", " disabled","") & ">" & String("system", "menus", "settings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_update_menu_permissions&mID="& g_sMenuID & iff(g_sWebform="webform_update_menu_permissions", " disabled","") & ">" & String("system", "menus", "permissions") & "</a></td></tr></table></div>"
		End If
		
		'-- print the form			
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=afterprocesswebform value='webform_list_menus'>"			
			.Write "<caption>" & String("system", "menus", "menu") & "</caption>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "menus", "menuid") & "</th><td><input type=text class=medium name=menuID" & iff(len(g_sMenuID)>0, " disabled", "")  & " value='" & g_sMenuID & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "name") & "</th><td><input type=text class=large name=name value=""" & name & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "published") & "</th><td>" & HtmlComponent_PublicationState("frmEdit", "published", published) & "</td></tr>"
			
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.go(-1);""></td></tr>"
			
			'-- delete button
			if len(g_sMenuID) then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_menu';document.forms[0].submit();}""></td></tr>"
		
			.Write "</form>"
			.Write "</table><br>"			
		End With
		
		
		'-- Display the list of pages for this menu
		If len(g_sMenuID) then
			Call webform_list_pages_bymenu
		End if
	End Sub
	
	
	Sub webform_update_menu_permissions 
		'-- Website settings, horizontal navigation
		Response.Write "<div class=tabs><table class=tabs cellspacing=0 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_menu&mID=" & g_sMenuID & iff(g_sWebform="webform_update_menu", " disabled","") & ">" & String("system", "menus", "settings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_update_menu_permissions&mID="& g_sMenuID & iff(g_sWebform="webform_update_menu_permissions", " disabled","") & ">" & String("system", "menus", "permissions") & "</a></td></tr></table></div>"
		Call EditObjectPermissions("/website/menus/menu[@id='" & g_sMenuID  & "']")
	End Sub
	
	
	'-- Insert a Menu
	Sub Do_Update_Menu
		Call UpdateNode (website_xml, "//menu[@id='" & g_sMenuID & "']", Array("name", "published"), Array(getParam("name"), getParam("published")))
	End Sub
	
	
	'-- update a menu
	Sub Do_Insert_Menu
		Dim menuID : menuID = getParam("menuID")
		Dim beforeXpath : if len(Request.QueryString("beforemenuID"))>0 then  : beforeXpath = "//menu[@id='" & Request.QueryString("beforemenuID") & "']" : end if
				
		'-- if the user has not indicated an ID
		If len(menuID)=0 Then
			Call InsertNode (website_xml, "/website/menus" , "menu", Array("name", "published"), Array(getParam("name"), getParam("published")), true, beforeXpath)
		Else
			'-- else we check that it's free
			if XPathChecker(website_xml, "/website/menus/menu[@id='" & menuID & "']")=0 then
				Call InsertNode (website_xml, "/website/menus" , "menu", Array("id", "name", "published"), Array(menuID, getParam("name"), getParam("published")), false, beforeXpath)
			Else
				response.write "<script>alert('" & String("system", "menus", "menuiderror") & "'); history.back();</script>"
			End If
		End If
				
	End Sub
	
	
	'-- Delete a menu
	Sub Do_Delete_Menu
		Call DeleteNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID & "']")
	End Sub
	
	
	'-- Move up a menu
	Sub Do_MoveUp_Menu
		Call MoveUpNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID & "']")
	End Sub
	
	
	'-- Move down a menu
	Sub Do_MoveDown_Menu
		Call MoveDownNode (website_xml, "/website/menus/menu[@id='" & g_sMenuID & "']")
	End Sub
%>