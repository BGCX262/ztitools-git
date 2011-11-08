<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Image(oContent)
		
		Call InsertUpdateExtraContent( oContent, array("path", "width", "height", "border", "align", "valign", "hspace", "vspace", "link", "target") )
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Image(oNode)
		Dim path, width, height, border, align, valign, hspace, vspace, link, target
								
		'-- We try to get the value of 'html', in the case of an update
		If not isempty(oNode) Then 
			path = GetChild(oNode, "path", "")
			width = GetChild(oNode, "width", "")
			height = GetChild(oNode, "height", "")
			border = GetChild(oNode, "border", "")
			align = GetChild(oNode, "align", "")
			valign = GetChild(oNode, "valign", "")
			hspace = GetChild(oNode, "hspace", "0")
			vspace = GetChild(oNode, "vspace", "0")
			link = GetChild(oNode, "link", "")
			target = GetChild(oNode, "target", "_self")
		Else
		    border = 0
		    hspace = 0
		    vspace = 0
		End If
		
		'-- Print the form elements
		With Response
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "path") & "</th><td>" & HtmlComponent_SelectImage("path", path) & "</td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "width") & "</th><td><input type=text class=small name=width value='" & width & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "height") & "</th><td><input type=text class=small name=height value='" & height & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "border") & "</th><td><input type=text class=small name=border value='" & border & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "hspace") & "</th><td><input type=text class=small name=hspace value='" & hspace & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "vspace") & "</th><td><input type=text class=small name=vspace value='" & vspace & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "align") & "</th><td>" & HtmlComponent_ImageAlign("align", align) & "</td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "valign") & "</th><td>" & HtmlComponent_ImageVAlign("valign", valign) & "</td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "link") & "</th><td><input type=text class=large name=link value='" & link & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("system", "contenttype_image", "target") & "</th><td>" & HtmlComponent_LinkTarget("target", target) & "</td></tr>"
		End With
	
	End Sub
		
	
	Function Render_System_Image(oContent)
	
		Dim path, height, width, border, align, hspace, vspace, alt, link, target
		path 	= GetChild(oContent, "path", "")
		height 	= GetChild(oContent, "height", "")
		width 	= GetChild(oContent, "width", "")
		border 	= cstr(GetChild(oContent, "border", ""))
		align 	= GetChild(oContent, "align", "")
		hspace 	= cstr(GetChild(oContent, "hspace", ""))
		vspace 	= cstr(GetChild(oContent, "vspace", ""))
		alt 	= trim(GetChild(oContent, "alt", ""))
		link 	= trim(GetChild(oContent, "link", ""))
		target 	= GetChild(oContent, "target", "")

		Render_System_Image = 	iff(align="center", "<center>", "") &_
								iff(len(link)>0, "<a href=""" & link & """" & iff(len(target)>0, " target=" & target, "") & ">", "") &_
								"<img src=""" & path & """" &_
								iff(len(align)>0 and align<>"center", " align=" & align, "") &_
								iff(len(border)>0, " border=" & border, "") &_
								iff(len(width)>0, " width=" & width, "") &_
								iff(len(height)>0, " height=" & height, "") &_
								iff(len(hspace)>0, " hspace=" & hspace, "") &_
								iff(len(vspace)>0, " vspace=" & vspace, "") &_
								iff(len(alt)>0, " alt=""" & alt & """", "") &_
								">" &_
								iff(len(link)>0, "</a>", "") &_
								iff(align="center", "</center>", "")


		
	End Function
%>