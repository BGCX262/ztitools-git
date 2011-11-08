<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Iframe(oContent)
		
		Call InsertUpdateExtraContent( oContent, array("src", "width", "height", "frameborder", "scrolling", "marginwidth", "marginheight", "bordercolor") )
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Iframe(oNode)
		Dim src, height, width, frameborder, scrolling, marginwidth, marginheight, bordercolor
								
		'-- We try to get the value of 'html', in the case of an update
		If not isempty(oNode) Then 
			src = GetChild(oNode, "src", "")
			width = GetChild(oNode, "width", "")
			height = GetChild(oNode, "height", "")
			frameborder = GetChild(oNode, "frameborder", "0")
			scrolling = GetChild(oNode, "scrolling", "no")
			marginwidth = GetChild(oNode, "marginwidth", "0")
			marginheight = GetChild(oNode, "marginheight", "0")
			bordercolor = GetChild(oNode, "bordercolor", "")
		Else
			frameborder = "0"
			scrolling = "no"
			marginwidth = "0"
			marginheight = "0"
		End If
		
		'-- Print the form elements
		With Response
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "url") & "</th><td><input type=text class=large name=src value='" & src & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "width") & "</th><td><input type=text class=small name=width value='" & width & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "height") & "</th><td><input type=text class=small name=height value='" & height & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "frameborder") & "</th><td><input type=text class=small name=frameborder value='" & frameborder & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "scrolling") & "</th><td><input type=text class=small name=scrolling value='" & scrolling & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "marginwidth") & "</th><td><input type=text class=small name=marginwidth value='" & marginwidth & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "marginheight") & "</th><td><input type=text class=small name=marginheight value='" & marginheight & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_iframe", "bordercolor") & "</th><td><input type=text class=small name=bordercolor value='" & bordercolor & "'></td></tr>"
		End With
	
	End Sub
		
	
	'-- Display the IFRAME
	Function Render_System_Iframe(oContent)
		Dim src, height, width, frameborder, scrolling, marginwidth, marginheight, bordercolor
		src = GetChild(oContent, "src", "")
		width = GetChild(oContent, "width", "")
		height = GetChild(oContent, "height", "")
		frameborder = GetChild(oContent, "frameborder", "0")
		scrolling = GetChild(oContent, "scrolling", "no")
		marginwidth = GetChild(oContent, "marginwidth", "0")
		marginheight = GetChild(oContent, "marginheight", "0")
		bordercolor = GetChild(oContent, "bordercolor", "")
					
		Render_System_Iframe = 	"<IFRAME SRC=""" & src & """" &_
								iff(len(width)>0, " width="& width, "") &_
								iff(len(height)>0, " height="& height, "") &_
								iff(len(frameborder)>0, " frameborder="& frameborder, "") &_
								iff(len(scrolling)>0, " scrolling="& scrolling, "") &_
								iff(len(marginwidth)>0, " marginwidth="& marginwidth, "") &_
								iff(len(marginheight)>0, " marginheight="& marginheight, "") &_
								iff(len(bordercolor)>0, " bordercolor="& bordercolor, "") &_
								"></IFRAME>"		
	End Function
%>