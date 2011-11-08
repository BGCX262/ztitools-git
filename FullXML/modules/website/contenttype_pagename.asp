<%
	'----------------------------------------------------------------
	'-- Execute the insert/update of specific data of this modules --
	'----------------------------------------------------------------
	Sub InsertUpdate_Website_PageName(oContent)
	
	End Sub
	
	
	'----------------------------------------------------------------
	'-- This function write the part of the INSERT/UPDATE FORM
	'----------------------------------------------------------------
	Sub Edit_Website_PageName(oNode)
	End Sub
	
	
	'----------------------------------------------------------------
	'-- Module rendering function
	'----------------------------------------------------------------
	Function Render_Website_PageName(oContent)
		Render_Website_PageName = "<span class=webpage_name>" & getAttribute(g_oWebPageXML.documentElement, "name", "")  & "</span>"
	End Function
%>