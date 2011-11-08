<%
	Sub webform_authentication_register
		Dim t
		Set t = new ASPTemplate
		t.TemplateDir = g_sServerMappath & ADMIN_FOLDER & "templates\"
		t.Template = "register.html"
		
		t.Slot("form:action") = g_sUrl
		t.Slot("form:ok") = String("system", "common", "ok")
		t.Slot("form:cancel") = String("system", "common", "cancel")
		't.Slot("value:message") = Message
		t.Slot("label:alreadyregister") = string("system", "authentication", "alreadyregister")
		
		t.Slot("label:email") = String("system", "users", "email")
		t.Slot("label:screenname") = String("system", "users", "screenname")
		t.Slot("label:password") = String("system", "users", "password")
		t.Slot("label:confirmpassword") = String("system", "users", "confirmpassword")
		t.Slot("label:firstname") = String("system", "users", "firstname")
		t.Slot("label:lastname") = String("system", "users", "lastname")
		t.Slot("label:address") = String("system", "users", "address")
		t.Slot("label:zipcode") = String("system", "users", "zipcode")
		t.Slot("label:city") = String("system", "users", "city")
		t.Slot("label:phone") = String("system", "users", "phone")
		t.Slot("label:country") = String("system", "users", "country")
		t.Slot("label:signature") = String("system", "users", "signature")
				
		t.Slot("form:process") = "do_authentication_register"
		t.Slot("form:afterprocesswebform") = "webform_authentication_result"
		t.Slot("label:caption") = String("system", "authentication", "create")
		t.Slot("label:message") = String("system", "authentication", "create")
		
		t.Slot("value:email")		= ""
		t.Slot("value:screenname")	= ""
		t.Slot("value:firstname")	= ""
		t.Slot("value:lastname")	= ""
		t.Slot("value:address1")	= ""
		t.Slot("value:address2")	= ""
		t.Slot("value:city")		= ""
		t.Slot("value:zipcode")		= ""
		t.Slot("value:phone")		= ""
		t.Slot("value:country")		= ""
		t.Slot("value:signature")	= ""
				
		t.Generate	
		
		Set t = Nothing
	end sub


	'-- This webform allow the user to update is profile
	Sub webform_authentication_profile
		Dim user_xml : user_xml = DATA_FOLDER & USERS_FOLDER & g_oUser.Login & XMLFILE_EXTENSION
		
		'-- on the user xml
		Dim oXML, oNodeList
		Set oXML = CreateDomDocument
		if not oXML.Load (user_xml) then
			LogIt "authentication.asp", "private_webform_user", ERROR, oXML.ParseError.reason, user_xml
			Exit Sub
		end if
		
		Dim t
		Set t = new ASPTemplate
		t.TemplateDir = g_sServerMappath & ADMIN_FOLDER & "templates\"
		t.Template = "profile.html"
		
		t.Slot("form:action") = g_sUrl
		t.Slot("form:ok") = String("system", "common", "ok")
		t.Slot("form:cancel") = String("system", "common", "cancel")
		
		t.Slot("label:email") = String("system", "users", "email")
		t.Slot("label:screenname") = String("system", "users", "screenname")
		t.Slot("label:firstname") = String("system", "users", "firstname")
		t.Slot("label:lastname") = String("system", "users", "lastname")
		t.Slot("label:address") = String("system", "users", "address")
		t.Slot("label:zipcode") = String("system", "users", "zipcode")
		t.Slot("label:city") = String("system", "users", "city")
		t.Slot("label:phone") = String("system", "users", "phone")
		t.Slot("label:country") = String("system", "users", "country")
		t.Slot("label:signature") = String("system", "users", "signature")
		
		
		t.Slot("label:changepassword") = String("system", "authentication", "changepassword")
		t.Slot("label:password") = String("system", "users", "password")
		t.Slot("label:confirmpassword") = String("system", "users", "confirmpassword")
					
		
		
		'-- select the node		
		Set oNodeList = oXML.SelectNodes("/user[@email='" & g_oUser.Login & "']")	
		
		'-- get each value
		if oNodeList.length=1 then				
			t.Slot("form:process") = "do_authentication_update"
			t.Slot("form:afterprocesswebform") = "webform_authentication_result"
			t.Slot("label:caption") = String("system", "authentication", "modify")
			t.Slot("label:message") = String("system", "authentication", "modify")
	
			t.Slot("value:email") = GetAttribute(oNodeList(0), "email", "")
			t.Slot("value:screenname") = GetAttribute(oNodeList(0), "screenname", GetAttribute(oNodeList(0), "alias", ""))
			t.Slot("value:firstname") = GetAttribute(oNodeList(0), "firstname", "")
			t.Slot("value:lastname") = GetAttribute(oNodeList(0), "lastname", "")
			t.Slot("value:address1") = GetAttribute(oNodeList(0), "address1", "")
			t.Slot("value:address2") = GetAttribute(oNodeList(0), "address2", "")
			t.Slot("value:city") = GetAttribute(oNodeList(0), "city", "")
			t.Slot("value:zipcode") = GetAttribute(oNodeList(0), "zipcode", "")
			t.Slot("value:phone") = GetAttribute(oNodeList(0), "phone", "")
			t.Slot("value:country") = GetAttribute(oNodeList(0), "country", "")
			t.Slot("value:signature") = GetAttribute(oNodeList(0), "signature", "")								
		end if
								
		t.Generate	
		
		Set oXML = Nothing
		Set t = Nothing
	end sub


	Sub webform_authentication_result(message)
		Response.Write "<script>"
		Response.Write "alert(""" & message & """);"
		Response.Write "</script>"
	'	response.End
	End Sub

	'-------------------------------------
	'-- This Sub display the Login form --
	'-------------------------------------
	Sub webform_authentication_login()
		if len(message)=0 then message = "logindetails"
		
		'-- print the form
		with Response
			.Write "<form method=post action=" & g_sURL & " name=frmOpenSession>"
			.Write "<input type=hidden name=process value=do_authentication_login>"
			.Write "<table border='0' cellspacing='5' cellpadding='4' width='100%' align='center' style='background-color:buttonface' id=tblfrm>"
			.Write "	<tr>"
			.Write "	<td><fieldset><legend>" & String("system", "contenttype_login", message) & "</legend><br>"
			.Write "	<table border='0' cellspacing='0' cellpadding='0' width='100%' align='center' style='background-color:buttonface'>"
			.Write "		<tr><td>&nbsp;" & String("system", "contenttype_login", "login") & ":</td><td><input type='text' size='25' name='login:login' value='" & getParam("login:login") & "'></td></tr>"
			.Write "		<tr><td>&nbsp;" & String("system", "contenttype_login", "password") & ":</td><td><input type='password' size='25' name='login:password'></td></tr> "
			.Write "		<tr><td colspan=2><input type=checkbox name=login:RememberCheckbox id=RememberCheckbox value=yes>&nbsp;<label for=RememberCheckbox>" & String("system", "contenttype_login", "rememberlogin") & "</label></td></tr>"
			.Write "	</table>"
			.Write "	</fieldset>"
			.Write "	</td>"
			.Write "	</tr>"
			.Write "	<tr>"
			.Write "		<td>"
			.Write "			<table border='0' cellspacing='0' cellpadding='0' width='95%' align='center'>"
			.Write "				<tr>"
			.Write "					<td width='100%' colspan='4' align='right'>"
			.Write "						<input type='submit' value='" & String("system", "common", "ok") & "' class='buttons'>"
			.Write "						&nbsp;&nbsp;&nbsp;"
			.Write "						<input type='button' value='" & String("system", "common", "cancel") & "' class='buttons' onClick='window.close()'>"
			.Write "					</td>"
			.Write "				</tr>"
			.Write "			</table>"
			.Write "		</td>"
			.Write "	</tr>"
			.Write "</table><sc"&"ript>window.document.forms['frmOpenSession'].elements['login:login'].focus();</scr"&"ipt>"
			.Write "</form>"		
		End with

	End Sub

	
	'-----------------------------------------------
	'-- Check the post of the authentication form --
	'-----------------------------------------------
	Function Do_Authentication_Login()
		'-- get and clean the posted form informations
		Dim uid : uid = replace(Request.Form("login:login"), "'", "")
		Dim pwd : pwd = replace(Request.Form("login:password"), "'", "")
			
		if len(uid)=0 or len(pwd)=0 then exit function
		
		'-- get user storage mode
		Dim userstorage
		userstorage = GetAttribute(g_oWebSiteXML.documentElement, "userstorage", "internal")
		
		
		'-- authenticate
		Dim ret
		SELECT CASE userstorage
			CASE "internal":
				ret = Authentication_Login_INTERNAL(uid, pwd)
				
			CASE "external:nt":
				ret = Authentication_Login_NT(uid, pwd)
			
			CASE "external:db":
				ret = Authentication_Login_DB(uid, pwd)
				
		END SELECT		
			
		
		'-- if user is authenticated and 'remember me' is on, write a cookie for 30 days			
		If ret then
			If getParam("login:RememberCheckbox")="yes" then
				Dim oRc4 
				Set oRc4 = New CRc4
				Response.Cookies("fx4")("u") = cstr(uid)
				Response.Cookies("fx4")("p") = oRc4.EnDeCrypt(cstr(pwd), AppSettings("RC4_KEY"))
				response.cookies("fx4").expires = now + 30
				Set oRc4 = Nothing
			End if
		Else
			Message = "loginfailed"
		End If
		
		Do_Authentication_Login = ret
			
	End Function
	
	
	'----------------------------------------------
	'-- Check the authentication from the cookie --
	'----------------------------------------------
	Function Do_Authentication_Cookie
		Dim uid : uid = replace(Request.Cookies("fx4")("u"), "'", "")
		Dim pwd : pwd = Request.Cookies("fx4")("p")
						
		if len(uid)=0 or len(pwd)=0 then exit function
		
		'-- decrypt password
		Dim oRc4 
		Set oRc4 = New CRc4
		pwd = oRc4.EnDeCrypt(cstr(pwd), AppSettings("RC4_KEY"))
		Set oRc4 = Nothing
					
		'-- get user storage mode
		Dim userstorage
		userstorage = GetAttribute(g_oWebSiteXML.documentElement, "userstorage", "internal")
		
		
		'-- authenticate
		Dim ret
		SELECT CASE userstorage
			CASE "internal":
				ret = Authentication_Login_INTERNAL(uid, pwd)
				
			CASE "external:nt":
				ret = Authentication_Login_NT(uid, pwd)
				
			CASE "external:db":
				ret = Authentication_Login_DB(uid, pwd)
				
		END SELECT
		
		Do_Authentication_Cookie = ret
	End Function
	
	
	
	'---------------------------
	'-- Check authentication  --
	'-- internal user storage --
	'-- todo: add login with screenName
	'---------------------------
	Function Authentication_Login_INTERNAL(p_username, p_password)
		Dim oXML
		Set oXML = CreateDomDocument		
		If oXML.load(users_xml) Then
			Dim oNodeList
			Set oNodeList = oXML.SelectNodes("/users/user[translate(@screenname, 'ABCDEFGHIJKLMNOPQRSTUVWXYZéèà', 'abcdefghijklmnopqrstuvwxyzeea')='"&p_username&"' or @email='"&lcase(p_username)&"']")
			if oNodeList.length=1 then
				dim email : email = getAttribute(oNodeList.item(0), "email", "")
								
				Dim user_xml : user_xml = DATA_FOLDER & USERS_FOLDER & email & XMLFILE_EXTENSION
		
				'-- Load the webmasters XML
				If oXML.load(user_xml) Then
					
					'-- check password
					If cstr(getAttribute(oXML.documentelement, "password", "")) = Md5(p_password) then
										
						'-- User is authenticated > store user infos into session			
						g_oUser.XmlDoc = oXML
						Authentication_Login_INTERNAL = true
										
					Else
						Authentication_Login_INTERNAL = false
					End if
				Else
					Authentication_Login_INTERNAL = false
				End if
			
			else
				Authentication_Login_INTERNAL = false
			end if
		else
			Authentication_Login_INTERNAL = false
		end if	
		
	End Function
	
	
	'---------------------------
	'-- Check authentication  --
	'-- internal user storage --
	'---------------------------
	Function Authentication_Login_NT(p_username, p_password)
		Dim externalntdomain, externalntadmgrp
		externalntdomain = GetAttribute(g_oWebSiteXML.documentElement, "externalntdomain", ".")
		externalntadmgrp = GetAttribute(g_oWebSiteXML.documentElement, "externalntadmgrp", "Administrators")
								
		on error resume next
		Dim oIADS
		Set oIADS = GetObject("WinNT:").OpenDSObject("WinNT://" & externalntdomain,  externalntdomain & "\" & p_username, p_password, 1)
		
		if Err<>0 then
			Authentication_Login_NT = false			
			err.Clear
		Else
			on error goto 0
			
			g_oUser.Login = p_username
			g_oUser.Group = "member"
			
			'-- get user in active directory
			Dim usr
			Set usr = getobject("WinNT://" & externalntdomain & "/" & p_username & ",user") 
			g_oUser.ScreenName = usr.FullName
			Set usr = Nothing
					
			'-- Now we check if the user is a member of the 'administrators' group
			Dim grp
			Set grp = getobject("WinNT://" & externalntdomain & "/" & externalntadmgrp & ",group") 
			If grp.IsMember("WinNT://" & externalntdomain & "/" & p_username) then
				g_oUser.Group = "administrator"
			end if
			Set grp = Nothing
			
			g_oUser.Culture =  "en"
				
			Authentication_Login_NT = true
		End if     
			       
		on error goto 0
		
		'destroy the object
		Set oIADS = Nothing
			
	End Function
	
	
	'---------------------------
	'-- Check authentication  --
	'-- internal user storage --
	'---------------------------
	Function Authentication_Login_DB(p_username, p_password)
		Dim externaldbcnn, externaldbtable, externaldbuserfield, externaldbpasswordfield, externaldbgroupfield, externaldbadmgrp
		externaldbcnn = GetAttribute(g_oWebSiteXML.documentElement, "externaldbcnn", "")
		externaldbtable = GetAttribute(g_oWebSiteXML.documentElement, "externaldbtable", "")
		externaldbuserfield = GetAttribute(g_oWebSiteXML.documentElement, "externaldbuserfield", "")
		externaldbpasswordfield = GetAttribute(g_oWebSiteXML.documentElement, "externaldbpasswordfield", "")
		externaldbgroupfield = GetAttribute(g_oWebSiteXML.documentElement, "externaldbgroupfield", "")
		externaldbadmgrp = GetAttribute(g_oWebSiteXML.documentElement, "externaldbadmgrp", "Administrators")
		
		
		Dim oConn
		Set oConn = server.CreateObject("adodb.connection")		
		oConn.connectionstring = replace(externaldbcnn, "{{servermappath}}", g_sServerMapPath)
		oConn.open 
					
		If oConn.State = 1 Then
			Dim oRst
			Set oRst = Server.CreateObject("adodb.recordset")
			
			'-- get user		
			oRst.Open "SELECT * from " & externaldbtable & " WHERE " & externaldbuserfield & "='" & p_username & "'", oConn, 3, 1
			
			if oRst.EOF OR oRst.BOF then				
				'-- user not found
				Authentication_Login_DB = false
			
			elseif orst.Fields(externaldbpasswordfield).Value <> cstr(p_password) then
				'-- wrong password
				Authentication_Login_DB = false
			else
				g_oUser.Login = p_username
				g_oUser.ScreenName = p_username
			
				'-- check if user is admin
				If orst.Fields(externaldbgroupfield).Value=externaldbadmgrp then
					g_oUser.Group = "administrator"
				else
					g_oUser.Group = "member"
				end if
				
				g_oUser.Culture =  "en"
				
				Authentication_Login_DB = true
			end if		
		Else
			Authentication_Login_DB = false
		End if
	End Function
	

	
	'--------------------
	'-- register profile --
	'--------------------
	Sub Do_Authentication_Register
		
		Dim email				: email = lcase(getParam("email"))
		Dim screenname			: screenname	= getParam("screenname")
		Dim password			: password = getParam("password")
		Dim password2			: password2 = getParam("password2")
		Dim registratedgroup	: registratedgroup = getAttribute(g_oWebSiteXML.DocumentElement, "registratedgroup", "member")	
		
		'-- todo: check fields for authorized value
		
		'-- First user creation is 'administrator'
		registratedgroup = iff(XPathChecker(users_xml, "users/user")=0, "administrator", registratedgroup)
			
		'-- missing fields
		If len(email)=0 OR len(screenname)=0 OR len(password)=0 Then
			Message = String("system", "authentication", "uncompleteform")
		
		'-- email is already used
		ElseIf XPathChecker(users_xml, "users/user[@email='" & email & "']")>0 then
			Message = String("system", "authentication", "emailisnotavailable")
		
		'-- Screen name is already used
		ElseIf XPathChecker(users_xml, "users/user[translate(@screenname, 'ABCDEFGHIJKLMNOPQRSTUVWXYZéèà', 'abcdefghijklmnopqrstuvwxyzeea')='" & screenname & "']")>0 then
			Message = String("system", "authentication", "screennameisnotavailable")
		
		'-- passwords are not matching
		Elseif cstr(password) <> cstr(password2) then
			Message = String("system", "contenttype_account", "unmatchingpassword")
		
		Else
	
			Call Private_Do_Insert_User(registratedgroup)
					
			Message = String("system", "authentication", "sucessfullyregistrated")
			
		End IF
		
		Call webform_authentication_result(message)
	End Sub
	
	
	'--------------------
	'-- Update profile --
	'--------------------
	Sub Do_Authentication_Update
		Dim password			: password = getParam("password")
		Dim password2			: password2 = getParam("password2")			
				
		'-- missing fields
		'if len(screenname)=0 Then
		'	Message = String("system", "contenttype_account", "uncompleteform")
		'
		'-- email is already used
		'ElseIf XPathChecker(users_xml, "users/user[@email='" & email & "']")>0 then
		'	Message = String("system", "contenttype_account", "emailisnotavailable")
		
		'-- Screen name is already used
		'ElseIf screenname<>g_oUser.ScreeNname AND XPathChecker(users_xml, "users/user[@screenname='" & screenname & "']")>0 then
		'	Message = String("system", "contenttype_account", "screennameisnotavailable")
		
		'-- password is bad
		'Else
			
	'		'-- update password if needed
	'		If len(password)>0 AND cstr(g_oUser.password) <> Md5(password) then
	'			Call UpdateNode (g_oUser.XmlDoc, "/user" , Array("password"), Array(md5(password)))
	'		End If
	'		
	'		'-- Update index
	'		Call UpdateNode (users_xml, "/users/user[@email='" & g_oUser.Login & "']" , Array("screenname"), Array(screenname))
	'			
	'		'-- Update user file
	'		Call UpdateNode (g_oUser.XmlDoc, "/user" , Array("screenname", "culture", "firstname", "lastname", "address1", "address2", "city", "zipcode", "phone", "signature"), Array(screenname, getParam("culture"), getParam("firstname"), getParam("lastname"), getParam("address1"), getParam("address2"), getParam("city"), getParam("zipcode"), getParam("phone"), getParam("signature")))
			
		'-- passwords are not matching
		if cstr(password) <> cstr(password2) then
			Message = String("system", "contenttype_account", "unmatchingpassword")
		
		Else
			Call private_Do_Update_User( g_oUser.Login, g_oUser.Group)
			Message = String("system", "contenttype_account", "sucessfullymodified")
		End IF
		
		webform_authentication_result(message)
	End Sub
	
	
	'-------------------------------------------
	'-- Force the destruction of user session --
	'-------------------------------------------
	Sub Do_Authentication_LogOff
		Response.Cookies("fx4")("u") = ""
		Response.Cookies("fx4")("p") = ""					
		response.cookies("fx4").expires = now - 30
		g_oUser.Reset
		
		Response.Redirect Request.ServerVariables("SCRIPT_NAME")
	End Sub
	
%>