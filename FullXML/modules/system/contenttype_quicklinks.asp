<%
	public nbPopup : nbPopup = 0

	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_System_QuickLinks(oContent)
		
		'-- append the 'text' child within a cdata section
		'Call SetChildNodeValue(oContent, "cdata", "text", getparam("text"), true)
		'Call InsertUpdateExtraContent( oContent, array("text") )
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_System_QuickLinks(oNode)
								
		'-- We display a form only when editing
		If not isempty(oNode) Then 
			Response.Write "<tr><td colspan=2>"
			Edit_System_QuickLinks_Links
			Response.Write "</td></tr>"
			'Response.Write "<tr class=datagrid_editrow valign=top><td colspan=2 align=center><a href="&g_sScriptName&"?action=Edit_System_QuickLinks_Links&pID="&g_sPageID&"&cID="& getAttribute(oNode, "id", "") & ">"& String("system","contenttype_quicklinks", "editlinks") & "</a></td></tr>"
		Else
			Response.Write "<tr class=datagrid_buttonrow><td colspan=2>" & "After you've created this module, you can edit ir to insert links." & "</td></tr>"
		End If
	End Sub
	
	
	Sub Edit_System_QuickLinks_Links()
		Dim oNode
		
        Dim contentID : contentID = request("id")
        
		Set oNode = g_oWebPageXML.selectSingleNode("//content[@id='" & contentID & "']")
		
		
		Dim aNam, aVal, oLink
		Set aNam = new Collection
		Set aVal = new Collection
		For each oLink in oNode.SelectNodes("link")
			aNam.Add( getAttribute(oLink, "name", ""))
			aVal.Add( getAttribute(oLink, "id", ""))
		Next
				
		Dim o
		Set o = new CtrlListEditor
		
		o.Title = String("system", "contenttype_quicklinks", "links")
		o.Labels = aNam.ToArray() 
		o.Values = aVal.ToArray()
		o.Height = 200
		o.Width = 570
		
		'-- buttons label
		o.AddButtonLabel = String("system", "common", "add")
		o.EditButtonLabel = String("system", "common", "edit")
		o.DeleteButtonLabel = String("system", "common", "delete")
		o.MoveUpButtonLabel = String("system", "common", "moveup")
		o.MoveDownButtonLabel = String("system", "common", "movedown")
		o.ConfirmDeleteWarning = String("system", "common", "confirmdelete")
		
		'url called on buttons clicks
		o.AddUrl		= g_sScriptName & "?action=Edit_System_QuickLinks_EDIT_LINK&pID=" & g_sPageID & "&cID=" & contentID & "&beforelinkID="
		o.EditUrl		= g_sScriptName & "?action=Edit_System_QuickLinks_EDIT_LINK&pID=" & g_sPageID & "&cID=" & contentID & "&linkID="
		o.DeleteUrl	    = g_sScriptName & "?action=Edit_System_QuickLinks_DELETE_LINK&pID=" & g_sPageID & "&cID=" & contentID & "&linkID="
		o.MoveDownUrl	= g_sScriptName & "?action=Edit_System_QuickLinks_MOVEDOWN_LINK&pID=" & g_sPageID & "&cID=" & contentID & "&linkID="
		o.MoveUPUrl	    = g_sScriptName & "?action=Edit_System_QuickLinks_MOVEUP_LINK&pID=" & g_sPageID & "&cID=" & contentID & "&linkID="
		
		o.Display	
		
		Set o = Nothing
	End Sub
	
	'-- insert / update a link into the quick links module
	Sub Edit_System_QuickLinks_EDIT_LINK()
		Dim linkID : linkID = Request("linkID")
		
		Dim process, name, link, popupparameters, target, t
        
        If len(linkID)>0 then
           process = "do_update_system_quicklinks"
           Dim oNode
           Set oNode = g_oWebPageXML.selectSingleNode("//content[@id='" & g_sContentID & "']/link[@id='" & linkID & "']")
           name = getAttribute(oNode, "name", "")
           link = getAttribute(oNode, "link", "")
           target = getAttribute(oNode, "target", "self")
           popupparameters = getAttribute(oNode, "popupparameters", "")
        Else
            process = "do_insert_system_quicklinks"
            name = ""
            link = ""
            target = "self"
            popupparameters = ""
        End if
        
		
		
		t = t & "<table cellspacing=0 cellpadding=0 class=datagrid>"
		t = t & "<form action='" & g_sScriptName & "?action=Edit_System_QuickLinks_Links&pID="&g_sPageID&"&cID="&g_sContentID&"' method=post id=frmEdit name=frmEdit>"
		t = t & "<input type=hidden name=process value='" & process & "'>"
		t = t & "<input type=hidden name=linkID value='" & linkID & "'>"
		t = t & "<caption>" & String("system", "contenttype_quicklinks", "link") & "</caption>"
		t = t & "<tr class=datagrid_editrow><th>" & String("system", "contenttype_quicklinks", "label") & "</th><td><input type=text class=large name=name value=""" & name & """></td></tr>"
		t = t & "<tr class=datagrid_editrow><th>" & String("system", "contenttype_quicklinks", "link") & "</th><td><input type=text class=large name=link value=""" & link & """></td></tr>"
		t = t & "<tr class=datagrid_editrow><th>" & String("system", "contenttype_quicklinks", "target") & "</th><td>" & HtmlComponent_Select("target", target, array("self", "blank", "popup"), array("_self", "_blank", "popup")) & "</td></tr>"
		t = t & "<tr class=datagrid_editrow><th>" & String("system", "contenttype_quicklinks", "popupparameters") & "</th><td><input type=text class=large name=popupparameters value=""" & popupparameters & """></td></tr>"
			
        '-- ok / cancel buttons	
		t = t & "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "back") & """ onclick=""history.go(-1);""></td></tr>"
		           
        t = t & "</form>"
		t = t & "</table>"
		
		
		response.Write t
	End Sub
	
	Sub Edit_System_QuickLinks_DELETE_LINK
        do_delete_system_quicklinks
    End Sub
    
    Sub Edit_System_QuickLinks_MOVEUP_LINK
        do_moveup_system_quicklinks
    End Sub
    
    Sub Edit_System_QuickLinks_MOVEDOWN_LINK
        do_movedown_system_quicklinks
    End Sub
    
    '--
    Private Sub do_insert_system_quicklinks 
        Dim linkID
        Dim name, link, target, popupparameters
        
        Dim arrayAttributes : arrayAttributes = array("name", "link", "target", "popupparameters")
		Dim arrayValues : arrayValues = array(getParam("name"), getParam("link"), getParam("target"), getParam("popupparameters"))
		
		'-- update the page node in the webtree
		Call InsertNode (g_oWebpageXML, "//content[@id='" & g_sContentID & "']", "link", arrayAttributes, arrayValues, true, "")
    
        '-- redirect to the link list
        response.redirect g_sScriptName & "?action=editcontent&pId=" & g_sPageID & "&id=" & g_sContentID
    
    End Sub
    
    '-- update the current node
    Private Sub do_update_system_quicklinks 
        Dim linkID : linkID = getParam("linkID")
        Dim name, link, popupparameters, target
        
        Dim arrayAttributes : arrayAttributes = array("name", "link", "target", "popupparameters")
		Dim arrayValues : arrayValues = array(getParam("name"), getParam("link"), getParam("target"), getParam("popupparameters"))
		
		'-- update the page node in the webtree
		Call UpdateNode (g_oWebpageXML, "//content[@id='" & g_sContentID & "']/link[@id='" & linkID & "']", arrayAttributes, arrayValues)
        
        '-- redirect to the link list
        response.redirect g_sScriptName & "?action=editcontent&pId=" & g_sPageID & "&id=" & g_sContentID
    End Sub
      
      
    '-- Delete the current node
    Private Sub do_delete_system_quicklinks 
    	Dim linkID : linkID = getParam("linkID")
        Call DeleteNode (g_oWebpageXML, "//content[@id='" & g_sContentID & "']/link[@id='" & linkID & "']")
        
        '-- redirect to the link list
        response.redirect g_sScriptName & "?action=editcontent&pId=" & g_sPageID & "&id=" & g_sContentID
    End Sub
       
    Private Sub do_moveup_system_quicklinks
    	Dim linkID : linkID = getParam("linkID")
        Call MoveUpNode (g_oWebpageXML, "//content[@id='" & g_sContentID & "']/link[@id='" & linkID & "']")
        
        '-- redirect to the link list
        response.redirect g_sScriptName & "?action=editcontent&pId=" & g_sPageID & "&id=" & g_sContentID
    End Sub
    
    
     Private Sub do_movedown_system_quicklinks
    	Dim linkID : linkID = getParam("linkID")
        Call MoveDownNode (g_oWebpageXML, "//content[@id='" & g_sContentID & "']/link[@id='" & linkID & "']")
        
        '-- redirect to the link list
        response.redirect g_sScriptName & "?action=editcontent&pId=" & g_sPageID & "&id=" & g_sContentID
    End Sub
       
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_System_QuickLinks(oContent)
        Dim oLink
        Render_System_QuickLinks = Render_System_QuickLinks & "<ul class=System_QuickLinks>" 
        
        For each oLink in oContent.SelectNodes("link")
            Dim linkID 	  : linkID = getAttribute(oLink, "id", "")
			Dim name 	  : name = getAttribute(oLink, "name", "/")
			Dim link 	  : link = getAttribute(oLink, "link", "/")
			Dim target 	  : target = getAttribute(oLink, "target", "_self")
			Dim popupparameters : popupparameters = getAttribute(oLink, "popupparameters", "")
			
            
			If target <> "popup" Then
			    Render_System_QuickLinks = Render_System_QuickLinks & "<li><a href=popup.asp?process=do_system_redirect&redirect=" & server.urlencode(link) & " class=System_QuickLinks target="& target & ">" & name & "</a></li>"
        	Else
			    nbPopup = nbpopup + 1
            	Render_System_QuickLinks = Render_System_QuickLinks & "<li><a href=""javascript: var pop" & nbPopup & " = window.open('popup.asp?process=do_system_redirect&redirect=" & server.urlencode(link) & "', 'pop" & nbPopup & "', '" & popupparameters & "');"" class=System_QuickLinks>" & name & "</a></li>"
        	End If
			
        Next
    
		Render_System_QuickLinks = Render_System_QuickLinks & "</ul>" 
        
	End Function
	
	
%>