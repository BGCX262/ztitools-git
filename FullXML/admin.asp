<% 
	OPTION EXPLICIT 
	Response.Buffer = true
	Response.Expires = 60
	Response.Expiresabsolute = Now() - 1
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
	<!-- #include file="engine/utilities/Helpers.asp" -->
	<!-- #include file="engine/utilities/HtmlComponents.asp" -->
	<!-- #include file="engine/utilities/md5.asp" -->
	<!-- #include file="engine/utilities/crc4.asp" -->
	<!-- #include file="engine/utilities/Collections.asp" -->
	<!-- #include file="engine/utilities/Xml_lib.asp" -->
	<!-- #include file="engine/utilities/asptemplate.asp" -->
	<!-- #include file="engine/utilities/ctransform.asp" -->
	<!-- #include file="engine/utilities/CLogFile.asp" -->
	<!-- #include file="engine/utilities/CtrlListEditor.asp" -->
	<!-- #include file="engine/utilities/cbrowser.asp"-->
	
	<!-- #include file="engine/global.asp" -->
	<!-- #include file="engine/fx4.lib.asp" -->
	<!-- #include file="engine/Web.Config.asp" -->
	<!-- #include file="FCKeditor/fckeditor.asp" -->
		
	<!-- #include file="engine/cuser.asp" -->
	<!-- #include file="engine/authentication.asp" -->
	<!-- #include file="modules/modules.asp" -->
	
	<!-- #include file="engine/admin/admin.website.asp" -->
	<!-- #include file="engine/admin/admin.websitetree.asp" -->
	<!-- #include file="engine/admin/admin.menus.asp" -->
	<!-- #include file="engine/admin/admin.pages.asp" -->
	<!-- #include file="engine/admin/admin.placeholders.asp" -->
	<!-- #include file="engine/admin/admin.contents.asp" -->
	
	<!-- #include file="engine/admin/admin.groups.asp" -->
	<!-- #include file="engine/admin/admin.permissions.asp" -->
	<!-- #include file="engine/admin/admin.users.asp" -->
	
	<!-- #include file="engine/admin/admin.modules.asp" -->
	<!-- #include file="engine/admin/admin.skins.asp" -->
	<!-- #include file="engine/admin/admin.stats.asp" -->
	<!-- #include file="engine/admin/admin.updates.asp" -->
	<%
		'-- Execute processes
		Call ExecuteProcess()
	%>
<html>
	<head>
		<title>Fullxml 4 - admin</title>
		<LINK HREF="engine/admin/templates/default/console.css" REL="stylesheet" TYPE="text/css">
		<script type="text/javascript" language="JavaScript" src="engine/admin/sortable.js"></script>
	</head>
	<body>
		<table class=header style="filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr='#0000FF', EndColorStr='#000000')">
			<tr>
				<th width="50%">
					<a href=default.asp><img src='engine/admin/media/home.png' align=absmiddle border=0></a> <%=g_oLocalVersion.SelectSingleNode("application/friendlyname").Text%> [<%=g_oUser.Culture %>]
				</th>
				<td align=right >
					<%	
						If g_oUser.Group<>"anonymous" Then
							Response.Write "<img src='engine/admin/media/user.png' align=absmiddle> "  & " <b>" & g_oUser.ScreenName & "</b> " & String("system", "interface", "authenticated") & " <a href=default.asp?process=Do_Authentication_LogOff><img src='engine/admin/media/close.png' border=0 alt='logoff'></a>"
						End If
					%>
				</td>
			</tr>
		</table>
		<table class="workspace" cellspacing="0" cellpadding="0">
			<tr valign="top">
				<td class="margin">
					<% Call ConsoleMenu() %>
				</td>
				<td class="separator"></td>
				<td class="main">
					<% Call DisplayWebForm() %>
				</td>
			</tr>
		</table>		
	</body>
</html>