<% OPTION EXPLICIT %>


<!-- #include file="../../../engine/utilities/helpers.asp" -->
<!-- #include file="../../../engine/utilities/md5.asp" -->
<!-- #include file="../../../engine/utilities/crc4.asp" -->
<!-- #include file="../../../engine/utilities/Xml_lib.asp" -->
<!-- #include file="../../../engine/utilities/HtmlComponents.asp" -->
<!-- #include file="../../../engine/utilities/asptemplate.asp" -->
<!-- #include file="../../../engine/utilities/ctransform.asp" -->
<!-- #include file="../../../engine/utilities/CLogFile.asp" -->
<!-- #include file="../../../engine/utilities/cbrowser.asp"-->

<!-- #include file="../../../engine/global.asp" -->
<!-- #include file="../../../engine/fx4.lib.asp"-->
<!-- #include file="../../../engine/Web.Config.asp" -->
<!-- #include file="../../../engine/cuser.asp" -->
<!-- #include file="../../../engine/authentication.asp"-->
<!-- #include file="../../../modules/modules.asp" -->

<!-- #include file="../../../engine/cpage.asp" -->

<!-- #include file="../../../engine/admin/admin.modules.asp" -->
<!-- #include file="../../../engine/admin/admin.skins.asp" -->

<%
	'-- Do the redirect and incrementer the counter
	Dim linkID : linkID = getParam("linkID")
	Dim url
	
	if len(linkID)>0 Then
					
		Dim oXML, oNodeList
		Set oXML = CreateDomDocument
		If NOT oXML.Load (links_xml) then
			LogIt "tool_links.asp", "System_EditLink", ERROR, oXML.ParseError.reason, links_xml
			FatalError "Fatal Error", "The ressource you try to access does not exist."
		End if
		
		Set oNodeList = oXML.SelectNodes("links/link[@id='" & linkID & "']")
		if oNodeList.length>0 then
			url = GetAttribute(oNodeList(0), "url", "")
			
			'' TODO :: increment with cookie check
			call SetChildNodeValue(oNodeList(0), "attribute", "count", clng(GetAttribute(oNodeList(0), "count", "0")) + 1, true)
			oXML.save links_xml
			'' --
			
			Set oXML = Nothing
			Response.Redirect url
		End If
		
		set oXML = Nothing			
	End If
%>