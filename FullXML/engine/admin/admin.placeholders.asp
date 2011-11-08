<%
		
	'-- Display the form to edit a placeholder ------------------------------------------------------
	'------------------------------------------------------------------------------------------------
	Sub webform_update_placeholder
		Dim process
		Dim placeholder : placeholder = Request.QueryString("placeholder")
		Dim paging, pagesize
		
	
		Dim oPlaceholderList
		Set oPlaceholderList = g_oWebPageXML.SelectNodes("/page/placeholders/placeholder[@id='" & placeholder & "']")
		If oPlaceholderList.Length=1 then
			process = "do_update_placeholder"
			pagesize = GetAttribute(oPlaceholderList(0), "pagesize", "10")
			paging = GetAttribute(oPlaceholderList(0), "paging", "false")
		Else
			process = "do_insert_placeholder"
			pagesize = 10
			paging = false
		End If
		
		
		'- print the form	
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=afterprocesswebform value='" & Server.URLEncode("webform_update_page&mID=" & g_sMenuID & "&pID=" & g_sPageID) & "'>"
			.Write "<caption>" & String("system", "placeholders", "placeholder") & " - " &placeholder  &  "</caption>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "placeholders", "paging") & "</th><td>" & HtmlComponent_Bool("frmEdit", "paging", paging) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "placeholders", "pagesize") & "</th><td>" & HtmlComponent_Number("frmEdit", "pagesize", pagesize, 1, 20) & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "cancel") & """ onclick=""document.location='" & g_sScriptName & "?webform=webform_update_page&mID=" & g_sMenuID & "&pID=" & g_sPageID & "';""></td></tr>"
						
			.Write "</table></form><br>"
		end with
		
		'-- contents list	
		Call webform_list_contents
	End sub
	
	
	'-- Insert a new placeholder
	Sub Do_Insert_Placeholder
		Dim placeholder : placeholder = Request.QueryString("placeholder")
		Call InsertNode (webpage_xml, "/page/placeholders" , "placeholder", Array("id", "paging", "pagesize"), Array(getParam("placeholder"), getParam("paging"), getParam("pagesize")), false, "")
	End Sub		
	
	
	'-- Update a placeholder
	Sub Do_Update_Placeholder
		Dim placeholder : placeholder = Request.QueryString("placeholder")
		Call UpdateNode (webpage_xml, "/page/placeholders/placeholder[@id='" & placeholder & "']", Array("paging", "pagesize"), Array(getParam("paging"), getParam("pagesize")))
	End Sub

	
	'-----------------------------------------------------------------
	'-- Insert the missing placeholder, used when inserting content --
	'-----------------------------------------------------------------
	Sub InsertPlaceHolder(placeholderid)				
		if g_oWebPageXML.selectNodes("/page/placeholders/placeholder[@id='" & placeholderid & "']").length=0 then
		
			Dim oPlaceHolder, att					
			Set oPlaceHolder = g_oWebPageXML.CreateElement("placeholder")
			Set att = g_oWebPageXML.createAttribute("id")		: att.value = placeholderid : oPlaceHolder.Attributes.SetNamedItem(att)
			Set att = g_oWebPageXML.createAttribute("pagesize")	: att.value = "10" : oPlaceHolder.Attributes.SetNamedItem(att)
			Set att = g_oWebPageXML.createAttribute("paging")	: att.value = "false" : oPlaceHolder.Attributes.SetNamedItem(att)
			
			g_oWebPageXML.SelectSingleNode("/page/placeholders").appendChild(oPlaceHolder)
			save_webpage			
		End if		
	End Sub
%>