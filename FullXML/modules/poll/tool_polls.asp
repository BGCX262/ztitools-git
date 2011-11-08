<%
	
	
	'----------------------------
	'-- Display the polls list --
	'----------------------------
	Sub webform_Poll_Polls
			Dim arrayLabel		: arrayLabel = Array(String("poll", "tool_polls", "pollquestion"))
			Dim arrayAttName	: arrayAttName = Array("question")
			Call XmlDatagrid("polls", polls_xml, "/polls/poll", arrayLabel, arrayAttName, "webform_Poll_EditPoll", "id", "id", true)
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub webform_Poll_EditPoll()
		
		Dim pollID : pollID = getParam("id")
		Dim process : process = "do_insert_poll"
		Dim question, i, ArrayOfChoices
			
				
		'-- We try to get the value of 'text', in the case of an update
		If len(pollID)>0 Then 
			process = "do_update_poll"
			
			Dim oXML
			Set oXML = CreateDomDocument
			if not oXML.Load (DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION) then
				LogIt "tool_polls.asp", "do_Edit_poll", ERROR, oXML.ParseError.Reason, oXML.url
				Exit sub
			end if
			
			'-- get the question
			question = GetAttribute(oXML.DocumentElement, "question", "")
			
			
			'-- get each choice
			dim oList, nb
			Set oList = oXML.DocumentElement.SelectNodes("choice")
			nb = cint(oList.length)
			
			if nb>0 then
				redim ArrayOfChoices(nb)
								
				For i=LBound(ArrayOfChoices) to UBound(ArrayOfChoices)-1
					ArrayOfChoices(i) = GetAttribute(oList.item(i), "value", "")
					'Response.Write i
				Next
			end if
			
		Else
			redim ArrayOfChoices(4)
		End If
		
		'-- Display the form
		With Response
			.Write "<table cellspacing=0 cellpadding=0 class=datagrid>"
			.Write "<form action=" & g_sURL & " method=post id=frmEdit name=frmEdit>"
			.Write "<input type=hidden name=process value='" & process & "'>"
			.Write "<input type=hidden name=pollID value='" & pollID & "'>"
			.Write "<input type=hidden name=referer value='" & Request.QueryString("referer") & "'>"
			.Write "<caption>" & String("poll", "tool_polls", "poll") & "</caption>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("poll", "tool_polls", "pollquestion") & "</th><td><input type=text name=question value='" & question & "' class=medium></td></tr>"
			for i=LBound(ArrayOfChoices) to UBound(ArrayOfChoices)-1
				.Write "<tr class=datagrid_editrow valign=top><th>" & String("poll", "tool_polls", "pollchoice") & " " & (i+1) & "</th><td><input type=text name=choices class=large value='" & ArrayOfChoices(i) & "'></td></tr>"
			next
		
			'-- Empty choice, if user want to add one on edition
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("poll", "tool_polls", "pollchoice") & " " & (UBound(ArrayOfChoices)+1) & "</th><td><input type=text name=choices class=large value=''></td></tr>"
			
			.Write "<tr class=datagrid_buttonrow><td colspan=2><input type=submit value='" & String("system", "common", "ok") & "'>&nbsp;<input type=button value=""" & String("system", "common", "cancel") & """ onclick=""document.location='admin.asp?webform=webform_poll_polls';""></td></tr>"
			if len(pollID)>0 then .Write "<tr class=datagrid_buttonrow><td colspan=2><input type=button value='" & String("system", "common", "delete") & "' onclick=""if (confirm('" & String("system", "common", "confirmdelete") & "')) { document.forms[0].elements['process'].value = 'do_delete_poll';document.forms[0].submit();}""></td></tr>"
			.Write "</form>"
			.Write "</table>"
			
		End With
	
	End Sub
	
	
	Sub do_update_poll
		Dim pollID : pollID = getParam("pollID")
		
		'-- update index
		Call UpdateNode (polls_xml, "/polls/poll[@id='" & pollID & "']", Array("question"), Array(getParam("question")))
		
		Dim oXML, oRoot
		Set oXML = CreateDomDocument
		if not oXML.Load (DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION) then
			LogIt "tool_polls.asp", "do_Edit_poll", ERROR, oXML.ParseError.Reason, oXML.url
			Exit sub
		end if
		
		Set oRoot = oXML.documentelement
		
		'-- update question
		SetChildNodeValue oRoot, "attribute", "question", getparam("question"), true
		
		'-- INSERT/UPDATE CHOICES
		
		'-- if it's an update, get the nb of choices
		Dim oList, nb
		Set oList = oRoot.SelectNodes("choice")
		nb = cint(oList.length)
		
		
		'-- loop on each posted value
		Dim i, choices : choices = split(getparam("choices"), ",")
		For i=LBound(choices) to UBound(choices)
			
			'-- trim the current element value (remove space
			choices(i) = trim(choices(i))
			
			'-- if the choice is already there
			If i<nb then
				
				'-- UPDATE
				If Len(choices(i))>0 Then
					SetChildNodeValue oList.item(i), "attribute", "value", choices(i), true
					'oList.item(i).firstchild.text = choices(i)
				'-- DELETE
				Else
					oRoot.removechild(oList.item(i))
				End If
			
			'-- Choice is not there -> append
			Elseif len(choices(i))>0 Then
			
				'-- create the choice
				Dim oChoice
				Set oChoice = oXML.createelement("choice")
				AppendAttribute oChoice, "id", cstr(i)
				AppendAttribute oChoice, "count", "0"
				AppendAttribute oChoice, "value", cstr(choices(i))
				
				'-- append the choice
				oRoot.appendCHild(oChoice)	
			End If
		Next
		
		oXML.save DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION
		
		
		'-- 
		if len(Request.QueryString("referer"))>0 then
			Response.Redirect Request.QueryString("referer")
		else
			Response.Redirect g_sScriptName & "?webform=webform_poll_polls"
		end if
		
	End Sub
	
	
	
	
	'-- Create effectively the new poll
	Sub do_insert_poll
		Dim pollID
		pollID = InsertNode (polls_xml, "/polls" , "poll", Array("question"), Array(getParam("question")), true, "")
		
		Dim oXML, oRoot
		Set oXML = CreateDomDocument
		
		'-- doc elem
		set oRoot = oXML.CreateElement("poll")
		AppendAttribute oRoot, "id", pollID
		AppendAttribute oRoot, "question", getParam("question")
		
		'-- loop on choices
		Dim i, choices : choices = split(getparam("choices"), ",")
		For i=LBound(choices) to UBound(choices)
			'-- trim the current element value (remove space
			choices(i) = trim(choices(i))
			
			'-- create the choice
			if len(choices(i))>0 Then
				Dim oChoice
				Set oChoice = oXML.createelement("choice")
				AppendAttribute oChoice, "id", cstr(i)
				AppendAttribute oChoice, "count", "0"
				AppendAttribute oChoice, "value", cstr(choices(i))
				
				'-- append the choice
				oRoot.appendCHild(oChoice)	
			End If
		Next
		
		oXML.appendChild(oRoot)		
		oXML.save DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION
		
		
		if len(Request.QueryString("referer"))>0 then
			Response.Redirect Request.QueryString("referer")
		else
			Response.Redirect g_sScriptName & "?webform=webform_poll_polls"
		end if
	End Sub
	
	
	'-------------------
	'-- Delete a user --
	'-------------------
	Sub Do_Delete_Poll
		Dim pollID : pollID = getParam("pollID")
		Dim poll_xml : poll_xml = DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION
		
		Call DeleteNode (polls_xml, "/polls/poll[@id='" & pollID & "']")
		Call DeleteFile(poll_xml)
		
		Response.Redirect g_sScriptName & "?webform=webform_poll_polls"			
	End Sub
%>