<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_Login(oContent)
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_Login(oNode)
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_Login(oContent)
		  				
  		'-- Display the box
		Dim t
		Set t = new ASPTemplate
		t.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\system\"

		If g_oUser.Group="anonymous" Then
			t.Template = "login.html"
			t.Slot("login:url") = g_sUrl
			
			If Len(Message)>0 then
				t.Slot("value:message") = String("system", "contenttype_login", message) & "<br>"
			Else
				t.Slot("value:message") = ""
			End IF
			
			t.Slot("login:login") = String("system", "contenttype_login", "login")
			t.Slot("login:password") = String("system", "contenttype_login", "password")
			
			If getAttribute(g_oWebSiteXML.DocumentElement, "userstorage", "internal")="internal" then
				t.Slot("login:register") = String("system", "contenttype_login", "register")
				t.Slot("login:forgotpassword") = String("system", "contenttype_login", "forgotpassword")
			Else
				t.Slot("login:register") = ""
				t.Slot("login:forgotpassword") = ""
			End If
				
			t.Slot("login:rememberlogin") = String("system", "contenttype_login", "rememberlogin")
			t.Slot("login:ok") = String("system", "contenttype_login", "submit")
		Else
		    t.Template = "logged.html"
			t.Slot("label:loggedas") = String("system", "contenttype_login", "loggedas")
			
			If getAttribute(g_oWebSiteXML.DocumentElement, "userstorage", "internal")="internal" then
				t.Slot("label:profile") = String("system", "contenttype_login", "profile")
			Else
				t.Slot("label:profile") = ""
			End If
			
			t.Slot("label:logoff") = String("system", "contenttype_login", "logoff")
			t.Slot("value:screenname") = g_oUser.ScreenName
		End If
						
		Render_System_Login = t.GetOutput										
		set t = Nothing			
	End Function
%>