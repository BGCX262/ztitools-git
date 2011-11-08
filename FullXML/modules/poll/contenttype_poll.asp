<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_Poll_Poll(oContent)
		
		'-- INSERT/UPDATE THE POLLID
		Call SetChildNodeValue(oContent, "node", "pollID", getparam("pollID"), true)
				
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_Poll_Poll(oNode)
		Dim pollID
								
		'-- We try to get the value of 'text', in the case of an update
		If not isempty(oNode) Then 			
			pollID = GetChild(oNode, "pollID", "")
		End If
				
		'-- The form element
		Response.Write "<tr class=datagrid_editrow valign=top><th>" & String("poll", "contenttype_poll", "selectapoll") & "</th><td>" & HtmlComponent_PollsSelect("pollID", "") & "</td></tr>"
	
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_Poll_Poll(p_oContent)		
		Dim l_action	: l_action =  getparam("action")		
		Dim pollID		: pollID = getChild(p_oContent, "pollID", "")		
		
		Dim oXML : set oXML = CreateDomDocument
		if not oXML.Load (DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION) then
			LogIt "contenttype_poll.asp", "Render_Poll_Poll", ERROR, oXML.ParseError.Reason, oXML.url
			Exit Function
		end if
		
		if l_action = "do_pollvote" then
			Render_Poll_Poll = Poll_Results(oXML.documentelement)
		else
			Render_Poll_Poll = Poll_Form(oXML.documentelement)			
		end if
		
		Set oXML = Nothing
		
	End Function
	
%>