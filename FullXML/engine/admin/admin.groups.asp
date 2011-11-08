<%
	'---------------
	'-- Groups list --
	'---------------
	Sub webform_list_groups
		Call XmlDatagrid(String("system", "groups", "title"), groups_xml, "/groups/group", Array(String("system", "common", "name"), String("system", "common", "description"), String("system", "common", "enable")), Array("id", "description", "enable"), "webform_edit_group", "id", "id", true)
	End Sub
	
	
	'----------------------------
	'-- Group modification form --
	'----------------------------
	Sub webform_edit_group
		Dim groupID : groupID = getParam("id")
		Dim process : process = "do_insert_group"
		Dim enable, description, buildin, canupload, uploadlimit
				
		'-- If an id is passed, then we are editing the data, so load the old value
		if len(groupID)>0 Then
			process = "do_update_group"			
						
			Dim oXML, oNodeList, oNode
			Set oXML = CreateDomDocument
			if Not oXML.Load(groups_xml) then
				LogIt "admin.groups.asp", "EditGroup", ERROR, oXML.ParserError.Reason, groups_xml
				exit sub
			end if
							
			Set oNodeList = oXML.DocumentElement.SelectNodes("/groups/group[@id='" & groupID & "']")			
			If oNodeList.length>0 then				
				enable		= GetAttribute(oNodeList(0), "enable", appSettings("DEFAULT_PUBLICATION_STATE"))
				description	= GetAttribute(oNodeList(0), "description", "")
				buildin		= cbool(GetAttribute(oNodeList(0), "buildin", "false"))
				canupload	= cbool(GetAttribute(oNodeList(0), "canupload", "false"))
				uploadlimit	= clng(GetAttribute(oNodeList(0), "uploadlimit", "1024"))
			End if
		else
			enable = appSettings("DEFAULT_PUBLICATION_STATE")
			buildin = false
			canupload = false
			uploadlimit = 1024
		End If
		
		
		'--------------------
		'-- Print the form --
		'--------------------
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=afterprocesswebform value='webform_list_groups'>"
			.Write "<caption>" & String("system", "groups", "title") & "</caption>"
			
			If BuildIn then
			    .Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "groupid") & "</th><td>" & groupID  &"</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "enable") & "</th><td>" & cstr(enable) & "</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "description") & "</th><td>" & description & "</td></tr>"
			'	.Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "canupload") & "</th><td>" & HtmlComponent_Bool("frmEdit", "canupload", canupload) & "</td></tr>"
			'	.Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "uploadlimit") & "</th><td><input type=text class=small name=uploadlimit value=""" & uploadlimit & """>Kb</td></tr>"
				
			Else				
				.Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "groupid") & "</th><td><input type=text class=large name=id value=""" & groupID & """" & iff(len(groupID)>0, " disabled", "")&"></td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "enable") & "</th><td>" & HtmlComponent_Bool("frmEdit", "enable", enable) & "</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "description") & "</th><td><textarea class=medium name=description>" & description & "</textarea></td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "canupload") & "</th><td>" & HtmlComponent_Bool("frmEdit", "canupload", canupload) & "</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "uploadlimit") & "</th><td><input type=text class=small name=uploadlimit value=""" & uploadlimit & """>Kb</td></tr>"
				
				
				.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.back(-1);""></td></tr>"
				
				'-- delete button --
				If len(groupID)>0 Then
					.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_group';document.forms[0].submit();}""></td></tr>"
				End If
			End If

		
			.Write "</form>"
			.Write "</table>"
				
		End With
		
		If len(groupID)>0 then
			response.write "<br>"
			Call XmlDatagrid(String("system", "users", "title"), users_xml, "/users/user[@group='"&groupID&"']", Array(String("system", "users", "screenname"), String("system", "groups", "group"), String("system", "users", "email")), Array("screenname", "group", "email"), "", "email", "id", false)
		End If
	End Sub
		
	
	'-------------------
	'-- update a Group
	'-------------------
	Sub Do_Insert_Group
		Dim groupID : groupID = getParam("id")
		Call InsertNode (groups_xml, "/groups" , "group", Array("id", "description", "enable", "canupload", "uploadlimit"), Array(groupID, getParam("description"), getParam("enable"), getParam("canupload"), getParam("uploadlimit")), false, "")
	End Sub
	
	
	'-------------------
	'-- Insert a Group --
	'-------------------
	Sub Do_Update_Group
		Dim groupID : groupID = getParam("id")
		Call UpdateNode (groups_xml, "/groups/group[@id='" & groupID & "']", Array("description", "enable", "canupload", "uploadlimit"), Array(getParam("description"), getParam("enable"), getParam("canupload"), getParam("uploadlimit")))
	End Sub
		
	
	'--------------------
	'-- Delete a Group --
	'--------------------
	Sub Do_Delete_Group
		Dim groupID : groupID = getParam("id")
		Call DeleteNode (groups_xml, "/groups/group[@id='" & groupID & "']")
	End Sub
%>