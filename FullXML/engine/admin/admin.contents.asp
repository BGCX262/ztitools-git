<%
	'-- Display the list of content ----------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Sub webform_list_contents
		Dim placeholder : placeholder = Request.QueryString("placeholder")
		if len(placeholder)>0 then
			Call XmlDatagrid(String("system", "contents", "contents"), g_oWebPageXML, "/page/placeholders/placeholder[@id='" & placeholder & "']/content", Array(String("system", "common", "name"), String("system", "modules", "contenttype"), String("system", "common", "startdatetime"), String("system", "common", "enddatetime"), String("system", "common", "published"), String("system", "common", "writer")), Array("name", "contenttype", "startdatetime", "enddatetime", "published", "writer"), "webform_update_content&mID=" & g_sMenuID & "&pID=" & g_sPageID & "&placeholder=" & placeholder, "id", "id", true)
		else
			Call XmlDatagrid(String("system", "contents", "contents"), g_oWebPageXML, "/page/placeholders/placeholder/content", Array(string("system", "common", "name"), string("system", "contents", "contenttype"), string("system", "contents", "startdatetime"), string("system", "contents", "enddatetime"), string("system", "contents", "published"), string("system", "contents", "writer")), Array("name", "contenttype", "startdatetime", "enddatetime", "published", "writer"), "webform_update_content&mID=" & g_sMenuID & "&pID=" & g_sPageID & "&placeholder=" & placeholder , "id", "id", true)
		end if
	End Sub

	
	'----------------------------------------------
	'-- Display the list of current user content --
	'----------------------------------------------
	Sub webform_list_submited_contents
		Call XmlDatagrid(String("system", "contents", "mycontents"), g_oWebPageXML, "/page/placeholders/placeholder/content[@writer='" & g_oUser.Login & "']", Array(string("system", "common", "name"), String("system", "common","type"), String("system", "common", "startdatetime"), String("system", "common", "enddatetime"), String("system", "common", "published"), String("system", "common", "writer")), Array("name", "contenttype", "startdatetime", "enddatetime", "published", "writer"), "webform_update_submited_content&mID=" & g_sMenuID & "&pID=" & g_sPageID, "id", "id", false)
	End Sub
	
	'----------------------------------------------
	'-- Display the list of current user content --
	'----------------------------------------------
	Sub webform_list_authored_contents
		Call XmlDatagrid(String("system", "contents", "mycontents"), g_oWebPageXML, "/page/placeholders/placeholder/content[@writer='" & g_oUser.Login & "']", Array(string("system", "common", "name"), String("system", "common","type"), String("system", "common", "startdatetime"), String("system", "common", "enddatetime"), String("system", "common", "published"), String("system", "common", "writer")), Array("name", "contenttype", "startdatetime", "enddatetime", "published", "writer"), "webform_update_content&mID=" & g_sMenuID & "&pID=" & g_sPageID, "id", "id", false)
	End Sub
	
	'----------------------------------------------
	'-- Display the list of current user content --
	'----------------------------------------------
	Sub webform_list_pending_contents
		Call XmlDatagrid(String("system", "contents", "pendingcontents"), g_oWebPageXML, "/page/placeholders/placeholder/content[@published='false']", Array(string("system", "common", "name"), String("system", "common","type"), String("system", "common", "startdatetime"), String("system", "common", "enddatetime"), String("system", "common", "published"), String("system", "common", "writer")), Array("name", "contenttype", "startdatetime", "enddatetime", "published", "writer"), "webform_update_content&mID=" & g_sMenuID & "&pID=" & g_sPageID, "id", "id", false)
	End Sub
	
	
	Sub webform_insert_submited_content()
		call private_edit_content("do_insert_submited_content", "")
	End Sub
	
	Sub webform_update_submited_content()
		call private_edit_content("do_update_submited_content", Request.QueryString("id"))
	End Sub
	
	Sub webform_insert_content()
		call private_edit_content("do_insert_content", "")
	End Sub
	
	Sub webform_update_content()
		call private_edit_content("do_update_content", Request.QueryString("id"))
	End Sub
	
	
	
	'----------------------------------------------------
	'-- Display the form for Content Edition/Insertion --
	'----------------------------------------------------
	Sub private_edit_content(process, contentID)		
						
		'-- execute process if form is posted
		if g_sprocess="do_delete_content" then exit sub
		
		' Content metadatas
		Dim name
		Dim contenttype		: contenttype = iff(len(Request.QueryString("contenttype"))>0, Request.QueryString("contenttype"), "system_text")
		Dim startdatetime	: startdatetime = YYYYMMDDHHNN(now())
		Dim enddatetime		: enddatetime = ""
		Dim published		: published = DEFAULT_PUBLICATION_STATE
		Dim writer			: 
		Dim placeholder		: placeholder = iff(len(Request.QueryString("placeholder"))>0, Request.QueryString("placeholder"), "main")
		Dim box				: box = "normal.html"
		Dim cachetimeout	: cachetimeout = 60
		
		'-- Get the skin from website settings		
		Dim skin : skin = GetAttribute(g_oWebSiteXML.DocumentElement, "skin", "")
		Dim arrCT
		Dim oNodeList
		Dim oNode : oNode = empty
			
		
		'-- If an id is passed, then we are editing the data, so load the old value
		if len(contentID)>0 Then			
			'process = "do_update_content"			
			Set oNodeList = g_oWebPageXML.SelectNodes("/page/placeholders/placeholder/content[@id='" & contentID & "']")			
			If oNodeList.length>0 then				
				set oNode = oNodeList(0)
				name = GetAttribute(oNode, "name", "")
				contenttype = GetAttribute(oNode, "contenttype", "")
				arrCT = split(contenttype, "_")

				startdatetime = GetAttribute(oNode, "startdatetime", "")
				enddatetime = GetAttribute(oNode, "enddatetime", "")
				published = cbool(GetAttribute(oNode, "published", DEFAULT_PUBLICATION_STATE))
				writer = GetAttribute(oNode, "writer", "")
				box = GetAttribute(oNode, "box", "normal.html")
				cachetimeout = GetAttribute(oNode, "cachetimeout", "60")
				placeholder = oNode.parentNode.attributes.getnameditem("id").value	
				
				'-- UPDATE Security check --
				If (g_oUser.isGranted(CONST_ACCESS_LEVEL_CONTRIBUTOR) and g_oUser.Login=writer and published=false) or (g_oUser.isGranted(CONST_ACCESS_LEVEL_AUTHOR) and g_oUser.ScreenName=writer) OR g_oUser.isGranted(CONST_ACCESS_LEVEL_MODERATOR) Then 
					'response.write "Update authorized"
				Else
					response.write "Update denied"
					Exit sub
				End If				
							
			Else
				response.Write "can't find content"
			End if			
		Else
						
			arrCT = split(contenttype, "_")

			if len(Request.QueryString("name"))>0 then
		        name = Request.QueryString("name")
		    else
		        name = String(arrCT(0), "contenttype_"&arrCT(1), "defaultname")
		    end if
			    						
			'-- INSERT Security check --
			If NOT g_oUser.isGranted(CONST_ACCESS_LEVEL_CONTRIBUTOR) Then 
				Response.Write "SECURITY :: Insert denied"
				Exit sub
			End If
			
		End If		
		

		'-- get the specific content type settings
		Dim caching, boxing
		caching 	= cbool(getAttribute(Application(APPVAR_DOM_MODULES).DocumentElement.SelectSingleNode("/modules/module[@name='"&arrCT(0)&"']/contenttypes/contenttype[@name='"&arrCT(1)&"']"), "caching", "false"))
		boxing 		= cbool(getAttribute(Application(APPVAR_DOM_MODULES).DocumentElement.SelectSingleNode("/modules/module[@name='"&arrCT(0)&"']/contenttypes/contenttype[@name='"&arrCT(1)&"']"), "boxing", "false"))

		
		'-- display the form
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=contentID value='" & contentID & "'>"
			.Write "<caption>" & String("system", "contents", "content") & "</caption>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "name") & "</th><td><input type=text class=large name=name value=""" & name & """></td></tr>"
			
			'-- Content Type	
			If len(contentID)>0 then
				.Write "<tr class=datagrid_editrow><th>" & String("system", "modules", "contenttype") & "</th><td><input type=hidden name=contenttype value='" & contenttype & "'>" & contenttype & "</td></tr>"
			Else
				.Write "<tr class=datagrid_editrow><th>" & String("system", "modules", "contenttype") & "</th><td>" & ContentTypeListBox(contenttype) & "</td></tr>"
			End If

			
			'-- Some fields are not available for contributor
			if  g_oUser.PagePermissionLevel >= CONST_ACCESS_LEVEL_AUTHOR Then
				
				' boxes list only if boxing is allowed for this content type
				if boxing then
					.Write "<tr class=datagrid_editrow><th>" & String("system", "contents", "box") & "</th><td>" & XMLListBox("box", "id", "name", skins_xml , "skins/skin[@id='" & skin & "']/box", box, array(""), array(String("system", "contents", "nobox"))) & "</td></tr>"
				else
					.Write "<tr class=datagrid_editrow><th disabled>" & String("system", "contents", "box") & "</th><td>"& String("system", "contents", "unavailable") & "</td></tr>"
				end if

				' Cachetimeout choice only if boxing is allowed for this content type
				if caching then
					.Write "<tr class=datagrid_editrow><th>" & String("system", "contents", "cachetimeout") & "</th><td><input type=text class=small name=cachetimeout value=""" & cachetimeout & """></td></tr>"
				else
					.Write "<tr class=datagrid_editrow><th disabled>" & String("system", "contents", "cachetimeout") & "</th><td>"& String("system", "contents", "unavailable") & "</td></tr>"
				end if
				
				'placeholder
				if len(placeholder)>0 AND len(contentID)>0 then
					.Write "<tr class=datagrid_editrow><th>" & String("system", "placeholders", "placeholder") & "</th><td><input type=hidden name=placeholder value='" & placeholder & "'>" & placeholder & "</td></tr>"
				else
					.Write "<tr class=datagrid_editrow><th>" & String("system", "placeholders", "placeholder") & "</th><td>"&HtmlComponent_PlaceholderSelect("placeholder", placeholder)&"</td></tr>" 
				end if
			
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "published") & "</th><td>" & HtmlComponent_Bool("frmEdit", "published", published) & "</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "startdatetime") & "</th><td>" & HtmlComponent_DateTime("startdatetime", startdatetime) & "</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "enddatetime") & "</th><td>" & HtmlComponent_DateTime("enddatetime", enddatetime) & "</td></tr>"
			End If		
			
		
			''' Now display the specific part of the form, depending of the contenttype
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
			Execute ("Edit_" & contenttype & "(oNode)")
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>"
			.Write "	<input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;"
			if instr(1, lcase(g_sScriptName), "popup.asp")>0 then
				.Write "	<input type=button value=""" & String("system", "common", "close") & """ onclick='self.close();'>"
			else
				.Write "	<input type=button value=""" & String("system", "common", "cancel") & """ onclick=""document.location='" & g_sScriptName & "?webform=webform_update_page&mID=" & g_sMenuID & "&pID=" & g_sPageID & "';"">"
			end if
			.Write "</td></tr>"
			
			if len(contentID)>0 then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_content';document.forms[0].submit();}""></td></tr>"
		
			.Write "</form>"
			.Write "</table>"			
		End With
		
		Set oNode = Nothing		
	End sub
	
	
	'---------------------------------------------------------
	'-- Display a confirmation form before deleting content --
	'---------------------------------------------------------
	Sub webform_delete_content
		Dim contentID	: contentID = Request.QueryString("id")
		
		with Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid style='width: 400px'>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<caption>Confirmation</caption>"
			.Write "<input type=hidden name=process value='do_delete_content'>"
			.Write "<input type=hidden name=contentID value='" & contentID & "'>"
			
			.Write "<tr height=100><td colspan=2 align=center>" & String("system", "contents", "confirmdeletecontent") & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>"
			.Write "<input type=submit value='" & String("system", "common", "yes") & "'>"
			.Write "&nbsp;"
			.Write "<input type=button value=""" & String("system", "common", "no") & """ onclick=""self.close();"">"
			.Write "</td></tr>"
		
			.Write "</form>"
			.Write "</table>"
		end with
	End Sub


	'---------------------------------------------------------
	'-- Display a confirmation form before move up content --
	'---------------------------------------------------------
	sub webform_moveup_content
		Dim contentID	: contentID = Request.QueryString("id")
		
		with Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid style='width: 400px'>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<caption>Confirmation</caption>"
			.Write "<input type=hidden name=process value='do_moveup_content'>"
			.Write "<input type=hidden name=contentID value='" & contentID & "'>"
			
			.Write "<tr height=100><td colspan=2 align=center>" & String("system", "contents", "confirmmoveupcontent") & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>"
			.Write "<input type=submit value='" & String("system", "common", "yes") & "'>"
			.Write "&nbsp;"
			.Write "<input type=button value=""" & String("system", "common", "no") & """ onclick=""self.close();"">"
			.Write "</td></tr>"
		
			.Write "</form>"
			.Write "</table>"
		end with
	End Sub


	'---------------------------------------------------------
	'-- Display a confirmation form before move up content --
	'---------------------------------------------------------
	Sub webform_movedown_content
		Dim contentID	: contentID = Request.QueryString("id")
		
		with Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid style='width: 400px'>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<caption>Confirmation</caption>"
			.Write "<input type=hidden name=process value='do_movedown_content'>"
			.Write "<input type=hidden name=contentID value='" & contentID & "'>"
			
			.Write "<tr height=100><td colspan=2 align=center>" & String("system", "contents", "confirmmovedowncontent") & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>"
			.Write "<input type=submit value='" & String("system", "common", "yes") & "'>"
			.Write "&nbsp;"
			.Write "<input type=button value=""" & String("system", "common", "no") & """ onclick=""self.close();"">"
			.Write "</td></tr>"
		
			.Write "</form>"
			.Write "</table>"
		end with
	End Sub


	'---------------------------------------------------------
	'-- Display a confirmation form before move up content --
	'---------------------------------------------------------
	Sub webform_changebox_content
		Dim contentID	: contentID = Request.QueryString("id")
		
		with Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid style='width: 400px'>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<caption>Confirmation</caption>"
			.Write "<input type=hidden name=process value='Do_changebox_content'>"
			.Write "<input type=hidden name=contentID value='" & contentID & "'>"
			
			.Write "<tr height=100><td colspan=2 align=center>" & String("system", "contents", "confirmchangebox") & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>"
			.Write "<input type=submit value='" & String("system", "common", "yes") & "'>"
			.Write "&nbsp;"
			.Write "<input type=button value=""" & String("system", "common", "no") & """ onclick=""self.close();"">"
			.Write "</td></tr>"
		
			.Write "</form>"
			.Write "</table>"
		end with
	End Sub

	
	'-- this function is used ton insert when you only have contributor permissions
	Sub do_submit_content
		Call private_insert_content(true)
	End Sub
	
	Sub do_insert_content
		Call private_insert_content(false)
	End Sub
	
	'-- this function is used ton insert when you only have contributor permissions
	Sub do_update_submited_content
		Call private_update_content(true)
	End Sub
	
	Sub do_update_content
		Call private_update_content(false)
	End Sub
	
	'------------------------------------
	'-- Insert content  private method --
	'------------------------------------
	Sub private_insert_content(ForceUnpublished)
		Dim contentID, oContentNode, PublishedValue
		
		'-- special case for CONTRIBUTOR :: they submit unpublished content
		if ForceUnpublished then
			PublishedValue = "false"
		else
			PublishedValue = getParam("published")
		end if 
		
		'-- insert metadatas as parameters
		InsertPlaceHolder getParam("placeholder")
		
		'-- insert the content node into the placeholder, and put it in the first position!
		contentID = InsertNode (g_oWebPageXML, "/page/placeholders/placeholder[@id='" & getParam("placeholder") & "']" , "content", Array("name", "contenttype", "startdatetime", "enddatetime", "published", "writer", "box", "cachetimeout"), Array(getParam("name"), getParam("contenttype"), getParam("startdatetime"), getParam("enddatetime"), PublishedValue, g_oUser.Login, getParam("box"), getParam("cachetimeout")), true, "/page/placeholders/placeholder[@id='" & getParam("placeholder") & "']/content[1]")
				 
		'-- get the content node
		set oContentNode = g_oWebPageXML.SelectSingleNode("/page/placeholders/placeholder/content[@id='" & contentID & "']")
			
		'-- insert specific properties of the content
		execute("call InsertUpdate_" & GetParam("contenttype") & " (oContentNode)")
				
		'-- Save Webpage
		call save_webpage()
		
	'	'-- redirect
	'	if instr(g_sScriptName, "popup.asp")=0 then
	'		Response.Redirect g_sScriptName & "?webform=webform_update_placeholder&pID=" & g_sPageID & "&placeholder=" & getParam("placeholder")
	'	end if	
	End Sub
	
	
	'----------------------
	'-- Update a content --
	'----------------------
	Sub private_update_content(ForceUnpublished)
		Dim PublishedValue
		Dim oContentNode
		Dim contentID	: contentID = Request("contentID")
		
		'-- special case for CONTRIBUTOR :: they submit unpublished content
		if ForceUnpublished then
			PublishedValue = "false"
		else
			PublishedValue = getParam("published")
		end if 
		
		'-- Update the meta datas
		Call UpdateNode (g_oWebPageXML, "/page/placeholders/placeholder/content[@id='" & getParam("contentID") & "']" , Array("name", "startdatetime", "enddatetime", "published", "writer", "box", "cachetimeout"), Array(getParam("name"), getParam("startdatetime"), getParam("enddatetime"), PublishedValue, g_oUser.Login, getParam("box"), getParam("cachetimeout")))
		
		'-- get the content node
		set oContentNode = g_oWebPageXML.SelectSingleNode("/page/placeholders/placeholder/content[@id='" & contentID & "']")
				
		'-- update specific properties of the content
		execute("call InsertUpdate_" & GetParam("contenttype") & " (oContentNode)")
		
		
		'-- @todo : update cache content
		Application("content_" & contentID) = ""
		
		'-- Save Webpage
		call save_webpage()
		
		'-- redirect
		if instr(g_sScriptName, "popup.asp")=0 then
			Response.Redirect g_sScriptName & "?webform=webform_update_placeholder&pID=" & g_sPageID & "&placeholder=" & getParam("placeholder")
		end if
	End Sub
	
	
	'----------------------
	'-- delete a content --
	'----------------------
	Sub do_delete_content
		Dim contentID	: contentID = Request("contentID")
		Call DeleteNode (webpage_xml, "/page/placeholders/placeholder/content[@id='" & contentID & "']" )
		
		'-- redirect
		if instr(g_sScriptName, "popup.asp")=0 then
			Response.Redirect g_sScriptName & "?webform=webform_update_placeholder&pID=" & g_sPageID & "&placeholder=" & getParam("placeholder")
		end if
	End Sub
	
	
	'----------------------
	'-- Content MoveDown --
	'----------------------
	Sub do_movedown_content
		Dim contentID	: contentID = Request("contentID")
		Call MoveDownNode(webpage_xml, "/page/placeholders/placeholder/content[@id='" & contentID & "']" )
		
		'-- redirect
		if instr(g_sScriptName, "popup.asp")=0 then
			Response.Redirect g_sScriptName & "?webform=webform_update_placeholder&pID=" & g_sPageID & "&placeholder=" & getParam("placeholder")
		end if
	End Sub
			
	
	'----------------------
	'-- Content MoveUp
	'----------------------
	Sub do_moveup_content
		Dim contentID	: contentID = Request("contentID")
		
		Call MoveUpNode(webpage_xml, "/page/placeholders/placeholder/content[@id='" & contentID & "']" )
		
		if instr(g_sScriptName, "popup.asp")=0 then
			Response.Redirect g_sScriptName & "?webform=webform_update_placeholder&pID=" & g_sPageID & "&placeholder=" & getParam("placeholder")
		end if	
	End Sub		
	

	'---------------------------------
	'-- Change the box of a content --
	'---------------------------------
	Sub Do_changebox_content
		Dim contentID	: contentID = Request("contentID")
		Call UpdateNode (webpage_xml, "/page/placeholders/placeholder/content[@id='" & contentID & "']" , Array("box"), Array(getParam("box")))
		
		'-- clear cache 
		Application("content_" & contentID) = ""
	End sub

	Sub do_refresh_content
		Dim contentID	: contentID = Request("id")
		Application("content_" & contentID) = ""
	End sub
	
	
	'--------------------------------------------------------------------
	'-- Display a SELECT of available ContentTypes -------------------------------------------------------------------------
	'-- Inputs:
	'			- sSelectedValue: the pre selected value
	'-----------------------------------------------------------------------------------------------------------------------
	Function ContentTypeListBox(sSelectedValue)
		Dim oXML : set oXML = CreateDomDocument
		Dim oNodeList, oNode
		if IsNull(sSelectedValue) or LenB(sSelectedValue) = 0 then sSelectedValue = ""
		Dim arrSel : arrSel = split(sSelectedValue, ",")
		Dim j
				
		'-- Load the xml file
		If NOT oXML.load(modules_xml) Then
			LogIt "admin.contents.asp", "ContentTypeListBox" , ERROR, oXML.ParseError.reason, modules_xml
			Exit Function
		End If
		
		Set oNodeList = oXML.SelectNodes("/modules/module/contenttypes/contenttype")
		
		ContentTypeListBox = "<select name='contenttype' id='contenttype' onchange=""document.location='" & g_sScriptName & "?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&placeholder=" & request("placeholder") & "&name=" & "' + frmEdit.name.value + '&webform=" & g_sWebform &  "&contenttype=' + " & "this.options[this.selectedIndex].value;"">"
		
		
		'-- create the options
		For each oNode in oNodeList
			Dim mod_name : mod_name = oNode.parentnode.parentnode.Attributes.GetNamedItem("name").value
			Dim ctt_name : ctt_name = oNode.Attributes.GetNamedItem("name").value
			Dim opt_val	 : opt_val = mod_name & "_" & ctt_name
			
			ContentTypeListBox = ContentTypeListBox & "<option value='" & opt_val & "'" 
			'on test s'il fait parti des selections
			For j=lBound(arrSel) to uBound(arrSel) : If cstr( opt_val)=cstr(trim(arrSel(j))) Then : ContentTypeListBox = ContentTypeListBox & " selected" : End If : Next
			ContentTypeListBox = ContentTypeListBox & ">&nbsp;" & oNode.Attributes.GetNamedItem("name").value & "</option>"
		Next
		ContentTypeListBox = ContentTypeListBox & "</select>"
		
		set oNodeList = nothing
		Set oXML = Nothing		
	End Function
	
	
	'------------------------------------------------------------------------
	'-- Append some posted parameter as CDATA child of the oNode
	'-- the array parameter contains the list of request.form element name --
	'-- children are added with the same name
	'------------------------------------------------------------------------
	Sub InsertUpdateExtraContent(oNode, arrayNode)
		Dim index
		For index=LBound(arrayNode) to UBound(arrayNode)
			Call SetChildNodeValue(oNode, "cdata", arrayNode(index), getparam(arrayNode(index)), true)
		Next
	End Sub
	
%>