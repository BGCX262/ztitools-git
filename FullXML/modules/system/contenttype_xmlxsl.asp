<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_XmlXsl(oContent)	
		Call InsertUpdateExtraContent(oContent, array("xml", "xsl"))	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_XmlXsl(oNode)
		Dim xml, xsl
		', cachetimeout
								
		'-- We try to get the value of 'html', in the case of an update
		If not isempty(oNode) Then 
			xml = GetChild(oNode, "xml", "")
			xsl = GetChild(oNode, "xsl", "")
			'cachetimeout = GetChild(oNode, "cachetimeout", "")
		End If
		
		'-- The form element
		Response.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_xmlxsl", "xml") & "</th><td><input type=text name=xml class=large value='" & xml & "'></td></tr>"
		Response.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_xmlxsl", "xsl") & "</th><td><input type=text name=xsl class=large value='" & xsl & "'></td></tr>"
		'Response.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_xmlxsl", "cachetimeout") & "</th><td><input type=text name=cachetimeout class=large value='" & cachetimeout & "'></td></tr>"
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_XmlXsl(oContent)
		Dim xml, xsl	', cachename, cachetimeout
		'cachename = getAttribute(oContent, "id", "")
		xml = GetChild(oContent, "xml", "")
		xsl = GetChild(oContent, "xsl", "")
		'cachetimeout = GetChild(oContent, "cachetimeout", 3600)
		'Render_System_XmlXsl = XmlXsl(xml, xsl, cachename, cachetimeout)
		Render_System_XmlXsl = Transform(xml, xsl)
	End Function
%>