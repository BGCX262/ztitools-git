<%
	
	'-------------------------------------------------------------------------------------
	'-- Display the insert/update form for a website
	'-------------------------------------------------------------------------------------
	Sub webform_update_website
		Dim process : process = "do_update_website"
				
		'-- read actual values	
		Dim name			: name = GetAttribute(g_oWebSiteXML.documentElement, "name", "")
		Dim culture			: culture = GetAttribute(g_oWebSiteXML.documentElement, "culture", "")
		Dim skin			: skin = GetAttribute(g_oWebSiteXML.documentElement, "skin", "")
		Dim online			: online = GetAttribute(g_oWebSiteXML.documentElement, "online", "")
		Dim model			: model = GetAttribute(g_oWebSiteXML.documentElement, "model", "")
		Dim slogan			: slogan = GetAttribute(g_oWebSiteXML.documentElement, "slogan", "")
		Dim copyright		: copyright = GetAttribute(g_oWebSiteXML.documentElement, "copyright", "")
		Dim emailcomponent	:	emailcomponent	= GetAttribute(g_oWebSiteXML.documentElement, "emailcomponent", "")
		Dim smtpserver		: smtpserver	= GetAttribute(g_oWebSiteXML.documentElement, "smtpserver", "127.0.0.1")
		Dim userstorage		: userstorage = GetAttribute(g_oWebSiteXML.documentElement, "userstorage", "internal")
		
		'-- internal
		
		dim registratedgroup : registratedgroup	= GetAttribute(g_oWebSiteXML.documentElement, "registratedgroup", "member")
		
		
		'-- external:nt storage options
		Dim externalntdomain : externalntdomain = GetAttribute(g_oWebSiteXML.documentElement, "externalntdomain", "WORKGROUP")
		Dim externalntadmgrp : externalntadmgrp = GetAttribute(g_oWebSiteXML.documentElement, "externalntadmgrp", "Administrators")
				
		'-- external: database storage options
		Dim externaldbcnn			: externaldbcnn = GetAttribute(g_oWebSiteXML.documentElement, "externaldbcnn", "")
		Dim externaldbtable			:externaldbtable = GetAttribute(g_oWebSiteXML.documentElement, "externaldbtable", "")
		Dim externaldbuserfield		:externaldbuserfield = GetAttribute(g_oWebSiteXML.documentElement, "externaldbuserfield", "")
		Dim externaldbpasswordfield :externaldbpasswordfield = GetAttribute(g_oWebSiteXML.documentElement, "externaldbpasswordfield", "")
		Dim externaldbgroupfield	:externaldbgroupfield = GetAttribute(g_oWebSiteXML.documentElement, "externaldbgroupfield", "")
		Dim externaldbadmgrp		: externaldbadmgrp = GetAttribute(g_oWebSiteXML.documentElement, "externaldbadmgrp", "Administrators")
		
		
		'-- Website settings, horizontal navigation
		Response.Write "<div class=tabs><table class=tabs cellspacing=0 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_website" & iff(g_sWebform="webform_update_website", " disabled","") & ">" & String("system", "website", "settings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_edit_website_permissions" & iff(g_sWebform="webform_edit_website_permissions", " disabled","") & ">" & String("system", "website", "permissions") & "</a></td><td bgcolor=#FFEAFF><a href=admin.asp?webform=webform_list_menus" & iff(g_sWebform="webform_list_menus", " disabled","") & ">" & String("system", "website", "menus") & "</a></td></tr></table></div>"
		
		
		'-- Print website form
		With Response
			'-- the main form
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action='" & g_sURL & "' method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<caption>" & String("system", "website", "website") & "</caption>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "name") & "</th><td><input type=text class=large name=name value=""" & name & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "slogan") & "</th><td><input type=text class=large name=slogan value=""" & slogan & """></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "copyright") & "</th><td><input type=text class=large name=copyright value=""" & copyright & """></td></tr>"
						
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
			
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "culture") & "</th><td>" & HtmlComponent_Culture("culture", culture) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "skin") & "</th><td>" & XMLListBox("skin", "id", "id", skins_xml , "skins/skin", skin, null, null) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "common", "online") & "</th><td>" & HtmlComponent_Bool("frmEdit", "online", online) & "</td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
						
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "emailcomponent") & "</th><td>" & HtmlComponent_Select("emailcomponent", emailcomponent, g_arrEmailCom, g_arrEmailCom) & "</td></tr>"
			'.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "smtpserver") & "</th><td><input type=text class=large name=smtpserver value='" & smtpserver & "'></td></tr>"
			
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2>&nbsp;</td></tr>"
				
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "model") & "</th><td>" & HtmlComponent_Select("model", model, array("publicmodel", "signupmodel", "privatemodel"), array("0", "1", "2")) & "</td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "userstorage") & "</th><td>" 
			
			.Write "<select name=userstorage onchange='selOpt()'>"
				.write "<option value='internal'" & iff(userstorage="internal", " selected", "")  & ">" & String("system", "website", "internal") & "</option>"
				.write "<option value='external:nt'" & iff(userstorage="external:nt", " selected", "")  & ">" & String("system", "website", "external:nt") & "</option>"
				.write "<option value='external:db'" & iff(userstorage="external:db", " selected", "")  & ">" & String("system", "website", "external:db") & "</option>"
			.Write "</select>"	
			.Write "</td></tr>"
								
			.Write "</table><br>"
			
			'& HtmlComponent_Select("userstorage", userstorage, array(String("system", "website", "internal"), String("system", "website", "external:nt"), String("system", "website", "external:db"), String("system", "website", "external:ws")), g_arrUsrStorage) & 
			
			
			'-- Form dedicated to internal authentication
			.Write "<div id=op0 style='display: none;'>"
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<caption>" & String("system", "website", "userstorageoption") & " - " & String("system", "website", "internal")  & "</caption>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "internal:registratedgroup") & "</th><td>" & XMLListBox("registratedgroup", "id", "id", groups_xml, "/groups/group", registratedgroup, array(), array()) & "</td></tr>"
			.Write "</table><br>"
			.Write "</div>"
			
			'-- Form dedicated to nt authentication
			.Write "<div id=op1 style='display: none;'>"
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<caption>" & String("system", "website", "userstorageoption") & " - " & String("system", "website", "external:nt")  & "</caption>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:nt:domain") & "</th><td><input type=text class=large name=externalntdomain value='" & externalntdomain & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:nt:admgrp") & "</th><td><input type=text class=medium name=externalntadmgrp value='" & externalntadmgrp & "'></td></tr>"
			.Write "</table><br>"
			.Write "</div>"
			
			'-- Form dedicated to db authentication
			.Write "<div id=op2 style='display: none;'>"
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<caption>" & String("system", "website", "userstorageoption") & " - " & String("system", "website", "external:db")  & "</caption>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:db:cnn") & "</th><td><input type=text class=large name=externaldbcnn value='" & externaldbcnn & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:db:table") & "</th><td><input type=text class=medium name=externaldbtable value='" & externaldbtable & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:db:userfield") & "</th><td><input type=text class=medium name=externaldbuserfield value='" & externaldbuserfield & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:db:passwordfield") & "</th><td><input type=text class=medium name=externaldbpasswordfield value='" & externaldbpasswordfield & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:db:groupfield") & "</th><td><input type=text class=medium name=externaldbgroupfield value='" & externaldbgroupfield & "'></td></tr>"
			.Write "<tr class=datagrid_editrow><th>" & String("system", "website", "external:db:admgrp") & "</th><td><input type=text class=medium name=externaldbadmgrp value='" & externaldbadmgrp & "'></td></tr>"
			
			.Write "</table><br>"
			.Write "</div>" 
			
			'-- javascript to manage the hidden options forms
			.Write "<script language='javascript'>function selOpt() { for(var i=0; i<document.all.userstorage.options.length; i++) document.all['op'+i].style.display = 'none'; document.all['op'+document.all.userstorage.options.selectedIndex].style.display = 'block';} selOpt();</script>"
						
			'-- buttons
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "cancel") & """ onclick=""document.location='" & g_sScriptName & "';""></td></tr>"
			.Write "</table>"
			
			.Write "</form>"
			
						
		End With
		
	End Sub

	Sub webform_edit_website_permissions
		'-- Website settings, horizontal navigation
		Response.Write "<div class=tabs><table class=tabs cellspacing=0 cellpadding=0><tr><td bgcolor=#EAEAFF><a href=admin.asp?webform=webform_update_website" & iff(g_sWebform="webform_update_website", " disabled","") & ">" & String("system", "website", "settings") & "</b></td><td bgcolor=#EAFFEB><a href=admin.asp?webform=webform_edit_website_permissions" & iff(g_sWebform="webform_edit_website_permissions", " disabled","") & ">" & String("system", "website", "permissions") & "</a></td><td bgcolor=#FFEAFF><a href=admin.asp?webform=webform_list_menus" & iff(g_sWebform="webform_list_menus", " disabled","") & ">" & String("system", "website", "menus") & "</a></td></tr></table></div>"
		
		Call EditObjectPermissions("/website")
	End Sub
	
	
	'-------------------------------------------------------------------------------------
	'-- Update a website 
	'-------------------------------------------------------------------------------------
	Sub Do_Update_Website()
		Dim arrayAttributes, index 
		arrayAttributes = array("name", "isdefault", "culture", "skin", "online", "model", "slogan", "copyright", "registratedgroup", "emailcomponent", "smtpserver", "userstorage", "externalntdomain", "externalntadmgrp", "externaldbcnn", "externaldbtable", "externaldbuserfield", "externaldbpasswordfield", "externaldbgroupfield", "externaldbadmgrp")
				
		'-- Update each attributes of the website
		For index=LBound(arrayAttributes) to UBound(arrayAttributes)
			Call SetChildNodeValue(g_oWebSiteXML.DocumentElement, "attribute", arrayAttributes(index), getParam(arrayAttributes(index)), true)
		Next
		
		'-- save
		g_oWebSiteXML.Save website_xml
	End Sub
%>