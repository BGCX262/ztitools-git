<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_Search_Box(oContent)
				
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_Search_Box(oNode)
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_Search_Box(oContent)
		Dim oTemplate
		Set oTemplate = new AspTemplate		
		
		'-- define the template source
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\search\"
		oTemplate.Template = "search_box.html"
		
		'--	fill some variables
		oTemplate.Slot("search_button_label") = String("search", "contenttype_box", "search_button_label")
		oTemplate.Slot("search_value") = getparam("s")
		
		'-- get the output
		Render_Search_Box = oTemplate.GetOutput	
		
		'-- release object	
		Set oTemplate = Nothing		
	End Function
%>