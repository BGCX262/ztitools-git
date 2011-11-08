<% 
	OPTION EXPLICIT 
	Response.Buffer = true
%>

<!-- #include file="engine/utilities/helpers.asp" -->
<!-- #include file="engine/utilities/md5.asp" -->
<!-- #include file="engine/utilities/crc4.asp" -->
<!-- #include file="engine/utilities/Xml_lib.asp" -->
<!-- #include file="engine/utilities/HtmlComponents.asp" -->
<!-- #include file="engine/utilities/asptemplate.asp" -->
<!-- #include file="engine/utilities/ctransform.asp" -->
<!-- #include file="engine/utilities/CLogFile.asp" -->
<!-- #include file="engine/utilities/cbrowser.asp"-->

<!-- #include file="engine/global.asp" -->
<!-- #include file="engine/fx4.lib.asp"-->
<!-- #include file="engine/Web.Config.asp" -->
<!-- #include file="engine/cuser.asp" -->
<!-- #include file="engine/authentication.asp"-->
<!-- #include file="modules/modules.asp" -->

<!-- #include file="engine/cpage.asp" -->

<!-- #include file="engine/admin/admin.modules.asp" -->
<!-- #include file="engine/admin/admin.skins.asp" -->


<%	
	'-- Execute authentication process
	If g_sProcess="do_authentication_login" then
		if Do_Authentication_Login() then
			Response.Redirect g_sUrl
		end if
	
	ElseIf g_sProcess="do_authentication_register" then
		Call Do_Authentication_Register()
		
	
	ElseIf g_sProcess="do_authentication_update" and g_oUser.Group<>"anonymous" then
		Call do_authentication_update()
	
	ElseIf g_sProcess="do_authentication_logoff" and g_oUser.Group<>"anonymous" then
		Call do_authentication_logoff()
		Response.Redirect g_sUrl
	End If

	'-- Current webpage render
	Dim oPage
	Set oPage = new CPage		
		oPage.Display
	Set oPage = Nothing	
%>
