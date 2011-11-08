<%
	'----------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules --
	'----------------------------------------------------------------
	Sub InsertUpdate_Website_Name(oContent)
	
	End Sub
	
	
	'----------------------------------------------------------------
	'-- This function write the part of the INSERT/UPDATE FORM
	'----------------------------------------------------------------
	Sub Edit_Website_Name(oNode)
	End Sub
	
	
	'----------------------------------------------------------------
	'-- Module rendering function
	'----------------------------------------------------------------
	Function Render_Website_Name(oContent)
		Render_Website_Name = "<span class=website_name>" & getAttribute(g_oWebSiteXML.documentElement, "name", "")  & "</span><br>"
	End Function
%>