<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_navigation_AllMenu(oContent)
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_navigation_AllMenu(oNode)
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_navigation_AllMenu(oContent)
		Dim t
		Dim menu
		Dim oMenu, oMenuItem
		
		Set t = new ASPTemplate
		t.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\navigation\"
		t.Template = "menu.html"
		
		t.ClearBlock "MenuBlock"
		
		For each oMenu in g_oWebSiteXML.selectNodes("/website/menus/menu[@published='2']")
			
			Dim menuID : menuID = getAttribute(oMenu, "id", "")
			
			if readUserPermission("//menu[@id='"& menuID &"']", g_oUser.Login, g_oUser.Group)>=CONST_ACCESS_LEVEL_VIEWER then
						
				t.Slot("menu_name") = oMenu.Attributes.GetNamedItem("name").Value
				t.Slot("menu_id") = menuID
				
				'-- MenuItems
				t.ClearBlock "PageBlock"
				For each oMenuItem in oMenu.selectNodes("page[@published='2']")
					Dim pageID : pageID = getAttribute(oMenuItem, "id", "")
					If readUserPermission("//page[@id='" & pageID & "']", g_oUser.Login, g_oUser.Group)>= CONST_ACCESS_LEVEL_VIEWER Then
						t.Slot("page_link") = "default.asp?mID=" & menuID & "&pID=" & pageID
						t.Slot("page_name") = oMenuItem.Attributes.GetNamedItem("name").Value
						t.RepeatBlock "PageBlock"
					End If				
					
				Next
				
				'-- In the case of editor, add a ADDPAGE link
				if readUserPermission("//menu[@id='"& menuID &"']", g_oUser.Login, g_oUser.Group)>=CONST_ACCESS_LEVEL_ADMINISTRATOR then
				'if g_oUser.isGrantedFor(CONST_ACCESS_LEVEL_ADMINISTRATOR, oMenu.Attributes.GetNamedItem("id").Value) Then
					t.Slot("page_link") = "admin.asp?webform=webform_insert_page&mID=" & menuID
					t.Slot("page_name") = String("system", "webpages", "addpage")
					t.RepeatBlock "PageBlock"
				End If
				
				t.RepeatBlock "MenuBlock"
			End If			
		Next
		
		Render_navigation_AllMenu = t.GetOutput
		set t = Nothing			
	End Function
%>