<%
	'----------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules --
	'----------------------------------------------------------------
	Sub InsertUpdate_Website_Slogan(oContent)
	
	End Sub
	
	
	'----------------------------------------------------------------
	'-- This function write the part of the INSERT/UPDATE FORM
	'----------------------------------------------------------------
	Sub Edit_Website_Slogan(oNode)
	End Sub
	
	
	'----------------------------------------------------------------
	'-- Module rendering function
	'----------------------------------------------------------------
	Function Render_Website_Slogan(oContent)
		Render_Website_Slogan = "<span class=website_slogan>" & getAttribute(g_oWebSiteXML.documentElement, "slogan", "")  & "</span><br>"
	End Function
%>