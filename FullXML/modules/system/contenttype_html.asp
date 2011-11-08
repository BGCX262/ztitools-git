<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Html(oContent)
	
		'-- append the 'html' child within a cdata section
		'Call SetChildNodeValue(oContent, "cdata", "html", getparam("html"), true)
		Call InsertUpdateExtraContent( oContent, array("html") )
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Html(oNode)
		Dim html
								
		'-- We try to get the value of 'html', in the case of an update
		If not isempty(oNode) Then html = GetChild(oNode, "html", "")
		
		'-- The form element
		Response.Write "<tr class=datagrid_editrow valign=top><td colspan=2>"
	
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.Value = html
		oFCKeditor.CreateFCKeditor "html", "100%", 300
		Set oFCKeditor = Nothing

		Response.Write "</td></tr>"
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_Html(oContent)
		Render_System_Html = GetChild(oContent, "html", "")
	End Function
%>