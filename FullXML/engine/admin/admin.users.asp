<%
	'-------------------------------
	'-- Display the list of users --
	'-------------------------------
	Sub webform_list_users
		Call XmlDatagrid(String("system", "users", "title"), users_xml, "/users/user", Array(String("system", "users", "screenname"), String("system", "groups", "group"), String("system", "users", "email")), Array("screenname", "group", "email"), "webform_edit_user", "email", "id", true)
	End Sub
	
	
	'-------------------------------------------------
	'-- Display the form to insert or update a user --
	'-------------------------------------------------
	Sub webform_edit_user
		Dim process : process = "do_insert_user"
		Dim email, password
		Dim firstname, lastname, address1, address2, city, zipcode, phone, picture, screenname, signature
		Dim group, culture
		Dim oXML, oNodeList, oNode
		
		email = Request.QueryString("id")
				
		'-- If an id is passed, then we are editing the data, so load the old value
		if len(email)>0 Then
			process = "do_update_user"
			
			Dim user_xml : user_xml = DATA_FOLDER & USERS_FOLDER & email & XMLFILE_EXTENSION
			
			'-- on the user xml
			Set oXML = CreateDomDocument
			if not oXML.Load (user_xml) then
				LogIt "admin.users.asp", "EditUser", ERROR, oXML.ParseError.reason, user_xml
				Exit Sub
			end if
			
			'-- select the node		
			Set oNodeList = oXML.SelectNodes("/user[@email='" & email & "']")	
			
			'-- get each value
			if oNodeList.length>0 then				
				email = GetAttribute(oNodeList(0), "email", "")
				password = GetAttribute(oNodeList(0), "password", "")
				
				firstname = GetAttribute(oNodeList(0), "firstname", "")
				lastname = GetAttribute(oNodeList(0), "lastname", "")
				address1 = GetAttribute(oNodeList(0), "address1", "")
				address2 = GetAttribute(oNodeList(0), "address2", "")
				city = GetAttribute(oNodeList(0), "city", "")
				zipcode = GetAttribute(oNodeList(0), "zipcode", "")
				phone = GetAttribute(oNodeList(0), "phone", "")
				picture = GetAttribute(oNodeList(0), "picture", "")
				screenname = GetAttribute(oNodeList(0), "screenname", GetAttribute(oNodeList(0), "alias", ""))
				signature = GetAttribute(oNodeList(0), "signature", "")
								
				culture = GetAttribute(oNodeList(0), "culture", "")
				group = GetAttribute(oNodeList(0), "group", "member")
			end if
		Else
			culture = "en-gb"
			
			'todo: Put default
			group = getAttribute(g_oWebSiteXML.DocumentElement, "registratedgroup", "member")
		End If
		
				
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action=" & g_sURL & " method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=afterprocesswebform value='webform_list_users'>"
			
			
			.Write "<caption>" & String("system", "users", "user") & "</caption>"
			
			If Len(email)=0 Then
				.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "email") & "</th><td><input type=text class=large name=email value=""" & email & """></td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "screenname") & "</th><td><input type=text class=large name=screenname value=""" & screenname & """></td></tr>"
			
			Else
				.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "email") & "</th><td><input type=hidden name=email value=""" & email & """>" & email & "</td></tr>"
				.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "screenname") & "</th><td><input type=hidden  name=screenname value=""" & screenname & """>" & screenname & "</td></tr>"
			
			End If
			
			if len(email)=0 then .Write "<tr class=datagrid_editrow><th>" & String("system", "users", "password") & "</th><td><input type=password class=large name=password value=""" & password & """></td></tr>"
			
			if g_oUser.Group="administrator" then
				.Write "<tr class=datagrid_editrow><th>" & String("system", "groups", "group") & "</th><td>" & XMLListBox("group", "id", "id", groups_xml, "/groups/group", group, array(), array()) & "</td></tr>"
			end if
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "culture") & "</th><td>" & HtmlComponent_Culture("culture", culture) & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2></td></tr>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "firstname") & "</th><td><input type=text class=large name=firstname value=""" & firstname & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "lastname") & "</th><td><input type=text class=large name=lastname value=""" & lastname & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "address") & "</th><td><input type=text class=large name=address1 value=""" & address1 & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>&nbsp;</th><td><input type=text class=large name=address2 value=""" & address2 & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "city") & "</th><td><input type=text class=large name=city value=""" & city & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "zipcode") & "</th><td><input type=text class=small name=zipcode value=""" & zipcode & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "phone") & "</th><td><input type=text class=medium name=phone value=""" & phone & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "users", "signature") & "</th><td><textarea name=signature class=small>" & signature & "</textarea></td></tr>"
			
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.go(-1);""></td></tr>"
			if len(email)>0 AND XPathChecker(users_xml, "users/user[@email!='" & email & "' and @group='administrator']")>0 then 
				.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.frmEdit.process.value = 'do_delete_user';document.frmEdit.submit();}""></td></tr>"
			end if
			
			.Write "</form>"
			.Write "</table>"			
		End With
		
		Set oXML = Nothing
	End Sub
	
	
	Sub Do_Insert_User
		Call Private_Do_Insert_User(getParam("group"))
	End Sub
	
	
	'-----------------------
	'-- Insert a new User --
	'-----------------------
	Sub Private_Do_Insert_User(p_group)
		Dim email		: email = lcase(getParam("email"))
		Dim screenname	: screenname = lcase(getParam("screenname"))
		Dim user_xml : user_xml = DATA_FOLDER & USERS_FOLDER & email & XMLFILE_EXTENSION
		Dim user_upload_folder : user_upload_folder = upload_path & "\" & screenname
		
		'add verif doublon
		'add verif format email et screenname (regexp)
			
		Call InsertNode (users_xml, "/users" , "user", Array("email", "screenname", "group"), Array(email, screenname, p_group), false, "")
		Call CreateXmlFile(user_xml, "user")
		Call UpdateNode (user_xml, "/user" , Array("email", "screenname", "group", "culture", "firstname", "lastname", "address1", "address2", "city", "zipcode", "phone", "country", "signature", "password"), Array(email, screenname, p_group, getParam("culture"), getParam("firstname"), getParam("lastname"), getParam("address1"), getParam("address2"), getParam("city"), getParam("zipcode"), getParam("phone"), getParam("country"), getParam("signature"), md5(getParam("password"))))
		
		'-- create the media user folder
		on error resume next
			g_oFso.CreateFolder(user_upload_folder)
		
			if err.number<>0 then
				LogIt "admin.users.asp", "Do_Insert_User", ERROR, "Failed to create a user media folder ["&user_upload_folder&"]", err.number & ", " & err.Description
				err.Clear
			end if
		on error goto 0
		
	End Sub
	
	
	Sub Do_Update_User
		Call private_Do_Update_User( lcase(getParam("email")), getParam("group"))
	End Sub
	
	
	'-------------------
	'-- Update a user --
	'-------------------
	Sub private_Do_Update_User(email, group)
		Dim password : password = getParam("password")
		Dim user_xml : user_xml = DATA_FOLDER & USERS_FOLDER & email & XMLFILE_EXTENSION
		
		'-- update password if needed
		If len(password)>0 AND cstr(g_oUser.password) <> Md5(password) then
			Call UpdateNode (g_oUser.XmlDoc, "/user" , Array("password"), Array(md5(password)))
		End If
		
		'-- Call UpdateNode (users_xml, "/users/user[@email='" & email & "']", Array("email", "screenname", "group"), Array(email, getParam("screenname"), group))
		Call UpdateNode (user_xml, "/user" , Array("group", "culture", "firstname", "lastname", "address1", "address2", "city", "zipcode", "phone", "country", "signature"), Array(group, getParam("culture"), getParam("firstname"), getParam("lastname"), getParam("address1"), getParam("address2"), getParam("city"), getParam("zipcode"), getParam("phone"), getParam("country"), getParam("signature")))
		
	End Sub
	
	
	'-------------------
	'-- Delete a user --
	'-------------------
	Sub Do_Delete_User
		Dim email : email = getParam("email")
		Dim user_xml : user_xml = DATA_FOLDER & USERS_FOLDER & email & XMLFILE_EXTENSION
		
		'-- Avoid from deleting the last administrator
		if XPathChecker(users_xml, "users/user[@email!='" & email & "' and @group='administrator']")>0 then		
			Call DeleteNode (users_xml, "/users/user[@email='" & email & "']")
			Call DeleteFile(user_xml)
		End If
		
	End Sub
%>