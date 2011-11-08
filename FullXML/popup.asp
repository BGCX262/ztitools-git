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
<html>
	<head>
		<title>Fullxml 4 - admin</title>
		<LINK HREF="engine/admin/templates/default/console.css" REL="stylesheet" TYPE="text/css">
	</head>
	<body style="border:0px; margin: 0px">		
		<table cellspacing="0" cellpadding="0" id=tblEdit>
			<tr valign="top">				
				<td>
					<% 
						'-- Execute processes
						Call ExecuteProcess()
						Call DisplayWebform()					
					%>
				</td>
			</tr>
		</table>
		<%Call ResizePopup()%>
	</body>
</html>