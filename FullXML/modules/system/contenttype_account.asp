<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Account(oContent)
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Account(oNode)
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_Account(oContent)
		
		If getAttribute(g_oWebSiteXML.DocumentElement, "userstorage", "internal")<>"internal" then
			Render_System_Account = "<span class=largebold>Only available with internal fullxml user database.</span>"
			Exit Function
		End if
		
		Dim t
		Set t = new ASPTemplate
		t.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\system\"
		t.Template = "Account-light.html"
		
		t.Slot("form:action") = g_sUrl
		
		t.Slot("value:message") = Message
			
		If g_oUser.Group="anonymous" Then
			t.Slot("value:process") = "Do_Authentication_Register"
			t.Slot("account:subtitle") = String("system", "contenttype_account", "create")
			t.Slot("value:email") = ""
			t.Slot("value:screenname") = ""
			t.Slot("disable") = ""
		
			t.Slot("label:changepassword") = ""
		Else
			
			t.Slot("value:process") = "Do_Authentication_Update"
			t.Slot("disable") = "disabled"
			
			t.Slot("account:subtitle") = String("system", "contenttype_account", "modify")
			t.Slot("value:email") = g_oUser.Login
			t.Slot("value:screenname") = g_oUser.ScreenName
			
			t.Slot("label:changepassword") = String("system", "contenttype_account", "changepassword")
		End If
			
		t.Slot("label:email") = String("system", "users", "email")
		t.Slot("label:screenname") = String("system", "users", "screenname")
		t.Slot("label:password") = String("system", "users", "password")
		t.Slot("label:confirmpassword") = String("system", "users", "confirmpassword")
		t.Slot("label:uncompleteform") = String("system", "contenttype_account", "uncompleteform")
		t.Slot("label:unmatchingpasswords") = String("system", "contenttype_account", "unmatchingpasswords")
		
		t.Slot("label:submit") = String("system", "contenttype_account", "submit")
		
		Render_System_Account = t.GetOutput
		set t = Nothing			
	End Function
%>