<%
	DIM TOOLBAR_FILE : TOOLBAR_FILE = appSettings("TOOLBAR_FILE")
	
	'-- Give the path to the current toolbar data ----
	Function toolbar_xml
		toolbar_xml = DATA_FOLDER & TOOLBAR_FILE & XMLFILE_EXTENSION
	End Function
	
		
	'-- Display the toolbar editor 
	'-----------------------------------------------------------------------------------------------------------------------
	Sub Toolbar_ToolbarEditor
		Dim oXML
		Set oXML = CreateDomDocument
		if not oXML.Load (toolbar_xml) then
			LogIt "tool_toolbar.asp", "Toolbar_ToolbarEditor", ERROR, oXML.ParseError.reason, toolbar_xml
		end if
		
		'-- load the xml of the website cause we need to read some settings
		Dim oXW
		Set oXW = CreateDomDocument
		If NOT oXW.Load (website_xml) then
			Logit "tool_toolbar.asp", "Toolbar_ToolbarEditor", ERROR, oXW.ParseError.reason, website_xml
		end if
		
		'-- get the list of datas
		Dim aNam, aVal, oItem 
		Set aNam = new Collection
		Set aVal = new Collection
			
		For each oItem in oXML.SelectNodes("/toolbar/item")
			
			'-- Iten: login
			if oItem.attributes.GetNamedItem("linktype").value="login" then
				aNam.Add(getString("opensession"))
							
			'-- Item: Fullxml page
			ElseIf oItem.attributes.GetNamedItem("linktype").value="internal" then
				Dim pageID, oPageList
				pageID = getAttribute(oItem, "pageID", "")
				Set oPageList = oXW.documentelement.SelectNodes("/website/menus/menu/page[@id='" & pageID & "']") 
				if oPageList.Length=1 then
					aNam.Add(getAttribute(oPageList(0), "name", ""))
				else
					aNam.Add(gString("toolbar", "deletedpage"))
				end if				
				
				
			'-- Item : free link
			else			
				aNam.Add(oItem.attributes.GetNamedItem("name").value)				
			end if
		
		'-- the id is not depending of the link type 
		aVal.Add(oItem.attributes.GetNamedItem("id").value)
		
		Next
		
		
		
		Set oXML = Nothing
			
			
		'-- now that we have the datas, we create the form	
		Dim o
		Set o = new CtrlListEditor
		
		o.Title = getString("toolbar")
		o.Labels = aNam.ToArray()
		o.Values = aVal.ToArray()
		
		'-- buttons label
		o.AddButtonLabel = getString("add")
		o.EditButtonLabel = getString("edit")
		o.DeleteButtonLabel = getString("delete")
		o.MoveUpButtonLabel = getString("moveup")
		o.MoveDownButtonLabel = getString("movedown")
		o.ConfirmDeleteWarning = getString("confirmdelete")
		
		'asp functions that are called on buttons
		o.EditFunction = "EDIT_TOOLBARITEM"
		o.AddFunction = "ADD_TOOLBARITEM"
		o.DeleteFunction = "DELETE_TOOLBARITEM"
		o.MoveDownFunction = "MOVEDOWN_TOOLBARITEM"
		o.MoveUPFunction = "MOVEUP_TOOLBARITEM"
		
		o.Display	
		
		Set o = Nothing
	End Sub
	
	
	'-- function for the CtrlListEditor ------------------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Sub ADD_TOOLBARITEM(p_ID) : Response.Redirect g_sScriptName & "?action=toolbar_edititem&sID=" & g_iWebSiteID & "&beforemenuID=" & p_ID : end sub
	Sub EDIT_TOOLBARITEM(p_ID) : Response.Redirect g_sScriptName & "?action=toolbar_edititem&sID=" & g_iWebSiteID  & "&toolbaritemID=" & p_id : end sub
	Sub DELETE_TOOLBARITEM(p_ID) : Response.Redirect g_sScriptName & "?action=toolbar_edititem&sID=" & g_iWebSiteID  & "&toolbaritemID=" & p_id & "&process=do_delete_toolbaritem" : end sub
	Sub MOVEUP_TOOLBARITEM(p_ID) : Response.Redirect g_sScriptName & "?action=toolbar_edititem&sID=" & g_iWebSiteID  & "&toolbaritemID=" & p_id & "&process=do_moveup_toolbaritem" : end sub
	Sub MOVEDOWN_TOOLBARITEM(p_ID) : Response.Redirect g_sScriptName & "?action=toolbar_edititem&sID=" & g_iWebSiteID  & "&toolbaritemID=" & p_id & "&process=do_movedown_toolbaritem" : end sub
	
	
	
	
	'-- Display the form to insert a toolbar item --------------------------------------------------------------------------
	'-----------------------------------------------------------------------------------------------------------------------
	Sub Toolbar_EditItem
		Dim toolbaritemID : toolbarItemID = Request.QueryString("toolbaritemID")
		Dim process : process = "do_insert_toolbaritem"
		Dim name, link, target, pageID, linktype
		Dim oXML, oNodeList, oNode
		
		
		Set oXML = CreateDomDocument
		if not oXML.Load (toolbar_xml) then
			LogIt "tool_toolbar.asp", "Toolbar_EditItem", ERROR, oXML.ParseError.reason, toolbar_xml
		end if
		
		'-- If an id is passed, then we are editing the data, so load the old value
		if len(toolbaritemID)>0 Then
			
			process = "do_update_toolbaritem"
			
			Set oNodeList = oXML.SelectNodes("/toolbar/item[@id='" & toolbaritemID & "']")	
			
			If oNodeList.length>0 then				
				name = GetAttribute(oNodeList(0), "name", "")
				link = GetAttribute(oNodeList(0), "link", "")
				target = GetAttribute(oNodeList(0), "target", "")
				linktype = GetAttribute(oNodeList(0), "linktype", "internal")
				pageID = GetAttribute(oNodeList(0), "pageID", "")
			End if
		Else
			target = "_self"
			linktype = "internal"
		End If
		Set oXML = Nothing
		
		'-- display the form				
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=toolbaritemID value='" & toolbaritemID & "'>"
			.Write "<caption>" & getString("toolbaritem") & "</caption>"
			
			.Write "<tr class=datagrid_editrow><td colspan=2><input type=radio name=linktype value=external" & IFF(linktype="external", " checked", "") &"><b>" & gString("toolbar", "externallink") & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & gString("toolbar", "itemname") & "</th><td><input type=text class=large name=name value=""" & name & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & gString("toolbar", "link") & "</th><td>" & HtmlComponent_URL("link", link) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & gString("toolbar", "target") & "</th><td>" & HtmlComponent_Select("target", target, array("samewindow", "newwindow"), array("_self", "_blank")) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><td colspan=2></td></tr>"
			
			.Write "<tr class=datagrid_editrow><td colspan=2><input type=radio name=linktype value=internal" & IFF(linktype="internal", " checked", "") &"><b>" & gString("toolbar", "internallink") & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & gString("toolbar", "fxpage") & "</th><td>" & HtmlComponent_SelectPage("pageID", pageID) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><td colspan=2></td></tr>"
			
			.Write "<tr class=datagrid_editrow><td colspan=2><input type=radio name=linktype value=login" & IFF(linktype="login", " checked", "") &"><b>" & gString("toolbar", "login") & "</td></tr>"
			.Write "<tr class=datagrid_editrow><td colspan=2></td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & getString("ok") & "'>&nbsp;<input type=button value=""" & getString("cancel") & """ onclick=""document.location='" & g_sScriptName & "?action=toolbar_toolbareditor&sID=" & g_iWebSiteID & "';""></td></tr>"
			if len(toolbaritemID) then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & getString("delete") & "' onclick=""if (confirm('" & getString("confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_toolbaritem';document.forms[0].submit();}""></td></tr>"	
		End With		
	End Sub 
	
	
	'-- Insert a toolbar item
	Sub Do_Insert_toolbaritem
		Dim beforeXpath
		If len(Request.QueryString("beforemenuID"))>0 then
			beforeXpath = "/toolbar/item[@id='" & Request.QueryString("beforemenuID") & "']"
		End if
		
		Dim arrayAttributes	: arrayAttributes = Array("name", "link", "target", "linktype", "pageID")
		Dim arrayValues	: arrayValues = Array(getParam("name"), getParam("link"), getParam("target"), getParam("linktype"), getParam("pageID"))
		
		
		Call InsertNode (toolbar_xml, "/toolbar" , "item", arrayAttributes, arrayAttributes, true, beforeXpath)
		Response.Redirect g_sScriptName & "?action=Toolbar_ToolbarEditor&sID=" & g_iWebsiteID
	End Sub
	
	
	Sub Do_update_toolbaritem
		Dim toolbaritemID	: toolbaritemID = Request.QueryString("toolbaritemID")
		
		Dim arrayAttributes	: arrayAttributes = Array("name", "link", "target", "linktype", "pageID")
		Dim arrayValues		: arrayValues = Array(getParam("name"), getParam("link"), getParam("target"), getParam("linktype"), getParam("pageID"))
		
		Call UpdateNode (toolbar_xml, "/toolbar/item[@id='" & toolbaritemID & "']", arrayAttributes, arrayValues)
		
		Response.Redirect g_sScriptName & "?action=Toolbar_ToolbarEditor&sID=" & g_iWebsiteID
	End Sub
	
	
	Sub do_delete_toolbaritem
		Dim toolbaritemID	: toolbaritemID = Request.QueryString("toolbaritemID")
		Call DeleteNode (toolbar_xml, "/toolbar/item[@id='" & toolbaritemID & "']")
		Response.Redirect g_sScriptName & "?action=Toolbar_ToolbarEditor&sID=" & g_iWebsiteID
	End Sub
	
	
	Sub do_moveup_toolbaritem
		Dim toolbaritemID	: toolbaritemID = Request.QueryString("toolbaritemID")
		Call MoveUpNode (toolbar_xml, "/toolbar/item[@id='" & toolbaritemID & "']")
		Response.Redirect g_sScriptName & "?action=Toolbar_ToolbarEditor&websiteID=" & g_iWebsiteID
	End Sub		
	
	
	Sub do_movedown_toolbaritem
		Dim toolbaritemID	: toolbaritemID = Request.QueryString("toolbaritemID")
		Call MoveDownNode (toolbar_xml, "/toolbar/item[@id='" & toolbaritemID & "']")
		Response.Redirect g_sScriptName & "?action=Toolbar_ToolbarEditor&sID=" & g_iWebsiteID
	End Sub
%>