<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Text(oContent)
		
		'-- append the 'text' child within a cdata section
		'Call SetChildNodeValue(oContent, "cdata", "text", getparam("text"), true)
		Call InsertUpdateExtraContent( oContent, array("text") )
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Text(oNode)
		Dim text
								
		'-- We try to get the value of 'text', in the case of an update
		If not isempty(oNode) Then 
			text = GetChild(oNode, "text", "")
		End If
		
		'-- The form element
		Response.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_text", "text") & "</th><td><textarea name=text class=medium>" & text & "</textarea></td></tr>"
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_Text(oContent)
		Render_System_Text = GetChild(oContent, "text", "")
	End Function
%>