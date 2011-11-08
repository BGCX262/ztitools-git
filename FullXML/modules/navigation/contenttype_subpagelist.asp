<%
	'----------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules --
	'----------------------------------------------------------------
	Sub InsertUpdate_Navigation_SubPageList(oContent)
	End Sub
	
	
	'------------------------------------------------------------
	'-- This function write the part of the INSERT/UPDATE FORM --
	'------------------------------------------------------------
	Sub Edit_Navigation_SubPageList(oNode)
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_Navigation_SubPageList(oContent)
		Dim oSubPages, oSubPage

		Set oSubPages = g_oWebSiteXML.selectNodes("/website/menus/menu[@id='" & g_sMenuID  & "']//page[@id='"&g_sPageID&"']/page[@published='2']")

		If oSubPages.Length=0 then
            Set oSubPages = g_oWebSiteXML.selectNodes("/website/menus/menu[@id='" & g_sMenuID  & "']//page[@id='"&g_sPageID&"']/parent::*/page[@published='2']")
		end if

		if oSubPages.Length > 0 then
            Dim t

			Set t = new ASPTemplate
			t.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\navigation\"
			t.Template = "subpagelist.html"

			t.Slot("parentpage_name") = oSubPages.item(0).parentNode.attributes.getnameditem("name").value
			t.Slot("parentpage_link") = "default.asp?mID=" & g_sMenuID & "&pID=" & oSubPages.item(0).parentNode.Attributes.GetNamedItem("id").Value
			
			t.ClearBlock "SubPageBlock"
			For each oSubPage in oSubPages

				If readUserPermission("//page[@id='"&getAttribute(oSubPage, "id", "")&"']", g_oUser.Login, g_oUser.Group )>=CONST_ACCESS_LEVEL_VIEWER Then
					t.Slot("page_link") = "default.asp?mID=" & g_sMenuID & "&pID=" & oSubPage.Attributes.GetNamedItem("id").Value
					t.Slot("page_name") = oSubPage.Attributes.GetNamedItem("name").Value
					t.RepeatBlock "SubPageBlock"
				End If

			Next
			Render_Navigation_SubPageList = t.GetOutput
			set t = Nothing
		Else
		    Render_Navigation_SubPageList = ""
		End If
				
	End Function
%>