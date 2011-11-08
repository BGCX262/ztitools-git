<%	
	'----------------------------------------------------------------------------------
	' This file contains the functions used by the admin console for the permissions --
	'----------------------------------------------------------------------------------
		
	'---------------------------------------
	'-- Display a grid to edit permission --
	'---------------------------------------
	Sub webform_edit_ObjectPermissions_groups(m_sXPath)
			
		Dim oXML, oGroup, i
		set oXML = CreateDomDocument
		If Not oXML.Load(groups_xml) Then
			LogIt "admin.permissions.asp", "EditObjectPermissions", ERROR, oXML.parseerror.reason, groups_xml
			Exit sub
		End If
		
		With Response
		
			'-- Table caption
			.Write "<table class=datagrid cellspacing=0 cellpadding=0>"
			.Write "<caption>" & String("system", "permissions", "title") & "</caption>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='do_insertupdate_objectpermissions'>"
			.Write "<input type=hidden name=xpath value="""&m_sXPath&""">"
				
			
			'-- Table Header columns
			.Write "<tr class=datagrid_column><th>&nbsp;</th>"
			For i=lbound(g_arrAccessLevel) To ubound(g_arrAccessLevel)
				.Write "<th>" & String("system", "permissions", g_arrAccessLevel(i)) & "</th>"
			Next
			.Write "</tr>"
			
			
			'-- Display each group as a line with checkbox for each level
			For each oGroup In oXML.DocumentElement.SelectNodes("/groups/group")
				Dim tmpGroup : tmpGroup = GetAttribute(oGroup, "id", "")
				Dim tmpLevel : tmpLevel = readGroupPermission(m_sXpath, tmpGroup)
				
				
				
				.Write "<tr class=datagrid_row><th bgcolor=menu>&nbsp;" & tmpGroup & "</th>"
				For i=lbound(g_arrAccessLevel) To ubound(g_arrAccessLevel)
					.Write "<td align=center><input type=radio name='group_" & tmpGroup  & "' value='" & i & "'" & iff(cint(tmpLevel)=i, " checked", "") & iff(tmpGroup="administrator" or tmpGroup="webmaster", " disabled", "") & "></td>"
				Next		
				.Write "</th></tr>"
						
			Next
			
			
			'-- display each inserted user level
			Dim oNodeList, oUser
			Set oNodeList = g_oWebSiteXML.SelectNodes(m_sXPath & "/permission[@user]")
			if oNodeList.length>0 Then
				.Write "<tr class=datagrid_buttonrow><td colspan=7></td></tr>"
			end If
			
			For Each oUser in oNodeList
				Dim tmpUser : tmpUser = getAttribute(oUser, "user", "")
				Dim tmpUserLevel : tmpUserLevel = getAttribute(oUser, "value", CONST_ACCESS_LEVEL_VIEWER)
				
				.Write "<tr class=datagrid_row><th bgcolor=menu>&nbsp;" & tmpUser & "</th>"
				For i=lbound(g_arrAccessLevel) To ubound(g_arrAccessLevel)
					.Write "<td align=center><input type=radio name='user_" & tmpUser  & "' value='" & i & "'" & iff(cint(tmpUserLevel)=i, " checked", "") & "></td>"
				Next		
				.Write "</th></tr>"
			Next
				
		
		
			.Write "<tr class=datagrid_buttonrow><td colspan=" & (3 + UBound(g_arrAccessLevel) - lbound(g_arrAccessLevel)) & ">"
			.Write "<input type=submit value=""" & String("system", "common", "ok") & """> "
			.Write "<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.go(-1);"">"
			.Write "</td></tr>"
			.write "</form>"
			.Write "</table>"
			
			.write String("system", "permissions", "help")
			
			'-- 
			if len(message)>0 then
				.write "<p><span class=error>" & String("system", "permissions", message) & "</span>"
			end if
			
			
			'-- Add a user permission to this grid
			.Write "<form method=post action="&g_sUrl&">"
			.Write "<input type=hidden name=process value='do_insertupdate_userpermission'>"
			.Write "<input type=hidden name=xpath value="""&m_sXPath&""">"
			
			.Write "<b>" & String("system", "permissions", "adduserpermission") & "</b><br>"
			.Write string("system", "users", "email") &": <input type=text name=user class=medium>"
			.Write "<select name=permission>"
			For i=lbound(g_arrAccessLevel) To ubound(g_arrAccessLevel)
				.Write "<option value="&i&">" & String("system", "permissions", g_arrAccessLevel(i)) & "</option>"
			Next
			.Write "</select>"
			.Write" <input type=submit value="& String("system", "common", "add") &">"
			.Write "</form>"
		
		End With
				
		Set oXML = Nothing
	End Sub
	
	
	
	'------------------------------------------------
	'-- Insert/Update the permissions on an object --
	'-- xpath querystring is used to point the object
	'-- into website.xml
	'------------------------------------------------
	Sub do_insertupdate_objectpermissions
		Dim xpath :xpath = getParam("xpath")
		
		'-- delete all previous permissions
		Call DeleteNode (g_oWebSiteXML, xPath & "/permission")
		
		'-- load the list of groups	
		Dim oXML, oGroup
		set oXML = CreateDomDocument
		If Not oXML.Load(groups_xml) Then
			LogIt "admin.permissions.asp", "do_insertupdate_objectpermissions", ERROR, oXML.parseerror.reason, groups_xml
			Exit sub
		End If
		
		
		'-- Insert the permission level	for each group		
		For each oGroup In oXML.DocumentElement.SelectNodes("/groups/group[@id!='administrator' and @id!='webmaster']")
			If len(getParam("group_" & GetAttribute(oGroup, "id", "")))>0 Then
				Call InsertNode (g_oWebSiteXML, xPath, "permission", array("group", "value"), array(GetAttribute(oGroup, "id", ""), getParam("group_" & GetAttribute(oGroup, "id", ""))), false, "")
			End If
		Next
		
		'-- Insert the permission level	for each user		
		Dim Item
		For each Item In Request.Form
			If Left(Item, 5) = "user_"  then 
				Call InsertNode (g_oWebSiteXML, xPath, "permission", array("user", "value"), array(mid(Item, 6), getParam(Item)), false, "")
			End If
		Next
				
		'-- release object
		set oXML = Nothing
				
		'-- redirect to the form
		Response.Redirect g_sUrl

	End Sub
	
	
	'------------------------------------------------
	'-- Insert/Update the permissions on an object --
	'-- xpath querystring is used to point the object
	'-- into website.xml
	'------------------------------------------------
	Sub do_insertupdate_userpermission
		Dim xpath :xpath = getParam("xpath")
		Dim user :user = getParam("user")
		Dim permission :permission = getParam("permission")
				
		'-- load the list of users and check that the user exists
		Dim oXML
		Set oXML = CreateDomDocument		
		If oXML.load(users_xml) Then
			Dim oNodeList
			Set oNodeList = oXML.SelectNodes("/users/user[translate(@screenname, 'ABCDEFGHIJKLMNOPQRSTUVWXYZéèà', 'abcdefghijklmnopqrstuvwxyzeea')='"&user&"' or @email='"&lcase(user)&"']")
			if oNodeList.length=1 then
				dim email : email = getAttribute(oNodeList.item(0), "email", "")
					
				'-- delete all previous permissions for this user
				Call DeleteNode (g_oWebSiteXML, xPath & "/permission[@user='"&email&"']")
								
				'-- Insert the permission level	for each group		
				Call InsertNode (g_oWebSiteXML, xPath, "permission", array("user", "value"), array(email, permission), false, "")
			else
				message = "usernotfound"
			End If
		Else
			LogIt "admin.permissions.asp", "do_insertupdate_userpermission", ERROR, oXML.ParserError.Reason, users_xml
		End If
				
		'-- redirect to the form
		Response.Redirect g_sUrl & "&message=" & message

	End Sub

%>