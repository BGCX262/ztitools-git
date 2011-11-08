<%
	'----------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules
	'----------------------------------------------------------------
	Sub InsertUpdate_Navigation_SingleMenu(oContent)
		Call InsertUpdateExtraContent(oContent, array("menuID"))	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_navigation_SingleMenu(oNode)
		Dim menuID
								
		'-- We try to get the value of 'html', in the case of an update
		If not isempty(oNode) Then 
			menuID = GetChild(oNode, "menuID", "")
		End If
		
		'-- The form element
		Response.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_singlemenu", "menu") & "</th><td>" & XMLListBox("menuID", "id", "name", g_oWebSiteXML, "//menu", menuID, null, null)  & "</td></tr>"
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_navigation_SingleMenu(oContent)
		Dim menuID : menuID = GetChild(oContent, "menuID", "")
				
		Dim t
		Dim menu
		Dim oMenu, oMenuItem
		
		Set t = new ASPTemplate
		t.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\navigation\"
		t.Template = "menu.html"
		
		t.ClearBlock "MenuBlock"
		
		If readUserPermission("//menu[@id='"& menuID &"']", g_oUser.Login, g_oUser.Group)>=CONST_ACCESS_LEVEL_VIEWER then
					
			For each oMenu in g_oWebSiteXML.selectNodes("/website/menus/menu[@id='" & menuID  & "']")
								
				t.Slot("menu_name") = oMenu.Attributes.GetNamedItem("name").Value
			
				'-- MenuItems
				t.ClearBlock "PageBlock"
				For each oMenuItem in oMenu.selectNodes("page[@published='2']")
					
					If readUserPermission("//page[@id='"&getAttribute(oMenuItem, "id", "")&"']", , g_oUser.Login, g_oUser.Group)>=CONST_ACCESS_LEVEL_VIEWER Then
						t.Slot("page_link") = "default.asp?mID=" & menuID & "&pID=" & oMenuItem.Attributes.GetNamedItem("id").Value
						t.Slot("page_name") = oMenuItem.Attributes.GetNamedItem("name").Value
						t.RepeatBlock "PageBlock"
					End If				
					
				Next
				
				'-- In the case of editor, add a ADDPAGE link
				if readUserPermission("//menu[@id='"& menuID &"']", , g_oUser.Login, g_oUser.Group)>=CONST_ACCESS_LEVEL_ADMINISTRATOR then
					t.Slot("page_link") = "admin.asp?webform=webform_insert_page&mID=" & menuID
					t.Slot("page_name") = String("system", "webpages", "addpage")
					t.RepeatBlock "PageBlock"
				End If
				
				t.RepeatBlock "MenuBlock"
			Next		
		End If	
		
		Render_navigation_SingleMenu = t.GetOutput
		set t = Nothing			
	End Function
%>