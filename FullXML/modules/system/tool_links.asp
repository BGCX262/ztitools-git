<%
	'----------------
	'-- Links list --
	'----------------
	Sub System_Links
		Call XmlDatagrid("links", links_xml, "/links/link", Array(String("system", "tool_links", "label"), String("system", "tool_links", "url"), String("system", "tool_links", "count")), Array("label", "url", "count"), "system_editlink", "id", "id", true)
	End Sub
	
	
	Sub System_Links_Selector
		Dim oXml
		Set oXML = CreateDomDocument
		if Not oXML.Load(links_xml) then
			LogIt "tool_links.asp", "System_Links_Selector", ERROR, links_xml, oXML.parseError.Reason
			Exit sub
		end If
		
		Reponse.write "<table class=datagrid>"
		
		For Each Link in oXML.SelectNodes("links/link")
			Response.Write "<tr><th></th><td>"& getAttribute(oXML, "id", "") & "</td></tr"
		Next
		
		response.Write "</table>"
		
		set oXML = Nothing
	End Sub
	
	'----------------------------------
	'-- Form to insert/update a link --
	'----------------------------------
	Sub System_EditLink
		Dim process : process = "do_insert_link"
		Dim linkID	: linkID = getParam("id")
		Dim label, url, count : count = 0
		
		'-- If we edit a existing link 
		if len(linkID)>0 Then
			process = "do_update_link"
						
			Dim oXML, oNodeList
			Set oXML = CreateDomDocument
			If NOT oXML.Load (links_xml) then
				LogIt "tool_links.asp", "System_EditLink", ERROR, oXML.ParseError.reason, links_xml
				Exit Sub
			End if
			
			Set oNodeList = oXML.SelectNodes("links/link[@id='" & linkID & "']")
			if oNodeList.length>0 then
				label = GetAttribute(oNodeList(0), "label", "")
				url = GetAttribute(oNodeList(0), "url", "")
				count = GetAttribute(oNodeList(0), "count", "0")
			End If
		End If
		
		
		'-- Display the form
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<caption>" & String("system", "tool_links", "link") & "</caption>"
			.Write "<form action=" & g_sURL & " method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=linkID value='" & linkID & "'>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_links", "label") & "</th><td><input type=text class=large name=label value=""" & label & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_links", "url") & "</th><td><input type=text class=large name=url value=""" & url & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "tool_links", "count") & "</th><td><input type=text class=small name=count value=""" & count & """></td></tr>"
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.go(-1);""></td></tr>"
			if len(linkID)>0 then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.frmEdit.process.value = 'do_delete_link';document.frmEdit.submit();}""></td></tr>"
			.Write "</table>"
			
			'-- display the link to this ressource
			if len(linkID)>0 then .Write String("system", "tool_links", "link")& ": <input style='border: 1px #808080 solid; width: 600px;' disabled type=text value='modules/system/lib/redirect.asp?action=System_LinkRedirect&linkID=" & linkID & "' size=100><br><br>"
		
		End With
	End Sub
	
	
	'-----------------------
	'-- Insert a new Link --
	'-----------------------
	Sub Do_Insert_Link		
		Dim linkID
		linkID = InsertNode(links_xml, "/links" , "link", Array("label", "url", "count"), Array(getParam("label"), getParam("url"), getParam("count")), true, "")
			
		Response.Redirect g_sScriptName & "?action=System_EditLink&id=" & linkID
	End Sub
	
	
	'-----------------
	'-- Update Link --
	'-----------------
	Sub Do_Update_Link		
		Call UpdateNode(links_xml, "/links/link[@id='" & getParam("linkID") & "']", Array("label", "url", "count"), Array(getParam("label"), getParam("url"), getParam("count")))
		Response.Redirect g_sScriptName & "?action=system_links"
	End Sub
	
	
	'-------------------
	'-- Delete a user --
	'-------------------
	Sub Do_Delete_Link
		Call DeleteNode (links_xml, "/links/link[@id='" & getParam("linkID") & "']")
		Response.Redirect g_sScriptName & "?action=system_links"			
	End Sub
	
%>