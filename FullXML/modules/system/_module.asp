<%
	Sub webform_system_edit_module
	End sub
	
	Sub webform_system_help_module
	End sub
    
    
    '---------------------------------------------------
    '-- Action function that do a tracked redirection --
    '---------------------------------------------------
    Public Sub do_system_redirect
		Dim redirect : redirect = Request("redirect")
		   
        Dim oXML, oNodeList, oNode
		Set oXML = CreateDomDocument
		if Not oXML.Load(redirects_xml) then
			LogIt "system/_module.asp", "Redirect", ERROR, oXML.ParserError.Reason, redirects_xml
			exit sub
		end if
						
		Set oNodeList = oXML.DocumentElement.SelectNodes("/redirects/redirect[@value='" & redirect & "']")			
		If oNodeList.length = 0 then
        	Call InsertNode(oXML, "/redirects", "redirect", array("value", "count"), array(redirect, 1), false, "")
        
        ElseIf oNodeList.length = 1 then
        	Dim counterUpdate
            counterUpdate = cLng(getAttribute(oNodeList.item(0), "count", 1)) + 1
            Call UpdateNode(oXML,"/redirects/redirect[@value='" & redirect & "']", array("value", "count"), array(redirect, counterUpdate) )
        End if
        
        response.redirect redirect
	End Sub
%>