<%
	'---------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules
	'---------------------------------------------------------------
	Sub InsertUpdate_navigation_HorizontalMenu(oContent)
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_navigation_HorizontalMenu(oNode)
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_navigation_HorizontalMenu(oContent)
		Dim t
		
		t = t & "<script type=""text/javascript"" src=""modules/system/lib/javascript.js""></script>"
		t = t & "<script type=""text/javascript"" src=""modules/system/lib/browserdetect.js""></script>"
		
		t = t & "<ul id='hmenu'>"
		
		Dim oMenu
		For each oMenu in g_oWebSiteXML.selectNodes("/website/menus/menu[@published='2']")
			Dim menuID : menuID = getAttribute(oMenu, "id", "")
			
			If readUserPermission("//menu[@id='"& menuID &"']", g_oUser.Login, g_oUser.Group)>=CONST_ACCESS_LEVEL_VIEWER Then
				t = t & "<li><a href='#"&menuID&"' title="""">"& oMenu.Attributes.GetNamedItem("name").Value  &"<span>,</span></a>"
				t = t & "<ul>" & GetPageMenu(menuID, oMenu) & "</ul>"
				t = t & "</li>"
			End If			
		Next
		
		t = t & "</ul>"	
		t = t & "<script type=""text/javascript"">initmenu();</script>"
		
		Render_navigation_HorizontalMenu = t
	End Function
	
	
	'-- recursive function, used to display pages and subpages
	Function GetPageMenu(byref p_menuID, byref p_oNodes)
		Dim tt
		Dim oMenuItem
		For each oMenuItem in p_oNodes.selectNodes("page[@published='2']")
			
			Dim pageID : pageID = getAttribute(oMenuItem, "id", "")
						
			If readUserPermission("//page[@id='" & pageID & "']", g_oUser.Login, g_oUser.Group)>= CONST_ACCESS_LEVEL_VIEWER Then
				
				tt = tt & "<li><a href=default.asp?mID=" & p_menuID & "&pID=" & pageID &">"& oMenuItem.Attributes.GetNamedItem("name").Value &"<span>,</span></a>"
				
				If oMenuItem.selectNodes("page[@published='2']").Length>0 then
					tt = tt & "<ul>" & GetPageMenu(p_menuID, oMenuItem) & "</ul>"
				End If
				
				tt = tt & "</li>"
			
			End If
			
		Next
		
		GetPageMenu = tt
	End Function
%>