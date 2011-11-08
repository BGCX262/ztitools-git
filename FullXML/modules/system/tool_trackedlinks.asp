<%
	'------------------------
	'-- Tracked Links list --
	'------------------------
	Sub System_TrackedLinks
		Call XmlDatagrid(String("system", "tool_trackedlinks", "trackedlinks"), redirects_xml, "/redirects/redirect", Array(String("system", "tool_trackedlinks", "url"), String("system", "tool_trackedlinks", "count")), Array("value", "count"), "", "", "", false)
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
%>