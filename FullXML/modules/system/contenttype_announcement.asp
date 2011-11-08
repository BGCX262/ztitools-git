<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this module
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Announcement(oContent)
		
		Call InsertUpdateExtraContent(oContent, array("subtitle", "headlines", "pic", "link"))
				
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Announcement(oNode)
		Dim subtitle, headlines, pic, link
								
		'-- We try to get the value of 'html', in the case of an update
		If not isempty(oNode) Then 
			subtitle = GetChild(oNode, "subtitle", "")
			headlines = GetChild(oNode, "headlines", "")
			pic = GetChild(oNode, "pic", "")
			link = GetChild(oNode, "link", "")
		end if
		
		'-- The form element
		With Response
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_announcement", "subtitle") & "</th><td><input type=text name=subtitle class=large value=""" & subtitle & """></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_announcement", "image") & "</th><td>" & HtmlComponent_SelectImage("pic", pic) & "</td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_announcement", "headlines") & "</th><td><textarea name=headlines style='height: 60px'>" & headlines & "</textarea></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_announcement", "link") & "</th><td><input type=text name=link class=large value=""" & link & """></td></tr>"
			'.Write "<tr class=datagrid_editrow valign=top><td colspan=2>"
			'Call HtmlComponent_TextEditor("headlines", body, "100%", 200)
			'.Write "</td></tr>"
		End With
	End Sub
	
	
	':::::::::::::::::::::::::::::::::::::::::::
	':: Rendering function for the headline
	':::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_Announcement(oNode)
		Dim  announcement_title, announcement_subtitle, announcement_text, announcement_img, announcement_link
		announcement_title = GetAttribute(oNode, "name", "")
		announcement_subtitle = GetChild(oNode, "subtitle", "")
		announcement_text = GetChild(oNode, "headlines", "")
		announcement_img = GetChild(oNode, "pic", "")
		announcement_link = GetChild(oNode, "link", "")
				
		Dim oTemplate
		Set oTemplate = new AspTemplate
				
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\system\"
		oTemplate.Template = "announcement.html"
		
		oTemplate.Slot("announcement_title") = announcement_title
		oTemplate.Slot("announcement_subtitle") = announcement_subtitle
		oTemplate.Slot("announcement_link") = announcement_link
		oTemplate.Slot("announcement_text") = announcement_text
		oTemplate.Slot("announcement_img") = announcement_img
		oTemplate.Slot("announcement_more") = String("system", "contenttype_announcement", "more")
		
		
		Render_System_Announcement = oTemplate.GetOutput
		Set oTemplate = Nothing
		
	End Function	
%>