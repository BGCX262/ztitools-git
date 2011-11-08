<%
	'----------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules --
	'----------------------------------------------------------------
	Sub InsertUpdate_Website_Copyright(oContent)
	
	End Sub
	
	
	'----------------------------------------------------------------
	'-- This function write the part of the INSERT/UPDATE FORM
	'----------------------------------------------------------------
	Sub Edit_Website_Copyright(oNode)
	End Sub
	
	
	'----------------------------------------------------------------
	'-- Module rendering function
	'----------------------------------------------------------------
	Function Render_Website_Copyright(oContent)
		Render_Website_Copyright = "<span class=website_copyright>" & getAttribute(g_oWebSiteXML.documentElement, "copyright", "")  & "</span>"
	End Function
%>