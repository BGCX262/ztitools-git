<%
	DIM POLLS_FOLDER	: POLLS_FOLDER		= appSettings("POLLS_FOLDER")
	DIM POLLS_INDEXFILE	: POLLS_INDEXFILE	= appSettings("POLLS_INDEXFILE")	
	
	
	Function polls_xml
		polls_xml = DATA_FOLDER & POLLS_INDEXFILE & XMLFILE_EXTENSION
	end Function
	

	Sub webform_poll_edit_module
		'here the code to configure this module
	End sub
	
	Sub webform_poll_help_module
	End sub
	
	
	'-------------------------------
	'-- Select box with poll list --
	'-------------------------------
	Function HtmlComponent_PollsSelect(sName, sValue)
		Dim oDOM, oItem
		Set oDOM = CreateDomDocument
		
		'-- load polls index file
		if not oDOM.Load (polls_xml) then
			LogIt "contenttype_poll.asp", "HtmlComponent_PollsSelect", ERROR, oDOM.ParseError.reason, polls_xml
		end if
		
		HtmlComponent_PollsSelect = HtmlComponent_PollsSelect & "<select name='" & sName & "' onChange=""if (pollID.options.selectedIndex==0) {document.all.msg.innerText='" & String("poll", "contenttype_poll", "newpoll") & "';} else {document.all.msg.innerText='" & String("poll", "contenttype_poll", "editpoll") & "';}"">" &_
									"<option>---------------</option>"
		
		For each oItem in oDOM.SelectNodes("polls/poll")
			HtmlComponent_PollsSelect = HtmlComponent_PollsSelect & "<option value='" & getAttribute(oItem, "id", "") & "'" & iff(sValue = getAttribute(oItem, "id", ""), " selected", "") & ">"
			HtmlComponent_PollsSelect = HtmlComponent_PollsSelect & getAttribute(oItem, "question", "")
			HtmlComponent_PollsSelect = HtmlComponent_PollsSelect & "</option>"
		Next
		
		HtmlComponent_PollsSelect = HtmlComponent_PollsSelect & "</select>" &_
									"<a href=# onclick=""document.location='popup.asp?webform=webform_Poll_EditPoll&id=' + pollID.options[pollID.options.selectedIndex].value + '&referer=" & server.URLEncode(g_sURL) & "';"">" &_
									"&nbsp;<span id=msg>" & String("poll", "contenttype_poll", "newpoll") & "</span>" &_
									"</a>"
		
	End Function
	
	
	'--------------------------
	'-- Return the poll form --
	'--------------------------
	Function Poll_Form(oContent)
		dim pollID		: pollID = getAttribute(oContent, "id", "")
		
		'-- define the template source
		Dim oTemplate
		Set oTemplate = new AspTemplate
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\poll\"
		oTemplate.Template = "poll_form.html"
		
		oTemplate.Slot("poll_formaction") = g_sUrl
		oTemplate.Slot("poll_ID") = pollID
		oTemplate.Slot("poll_question") = GetAttribute(oContent, "question", "")
		
		oTemplate.clearBlock "poll_item_block"
		
		'loop on choice
		Dim oList, oChoice
		Set oList = oContent.SelectNodes("choice")
		
		For Each oChoice in oList
			oTemplate.Slot("poll_ID") = pollID
			oTemplate.Slot("choice_ID") = getAttribute(oChoice, "id", "")
			oTemplate.Slot("choice_text") = getAttribute(oChoice, "value", "")
			oTemplate.RepeatBlock "poll_item_block"
		Next
		
		'-- text label on the buttons
		oTemplate.Slot("poll_vote") = String("poll", "contenttype_poll", "pollvote")
		
		Poll_Form = oTemplate.GetOutput
		
		set oTemplate = Nothing	
	End Function
	
	
	'-----------------------------
	'-- Return the poll results --
	'-----------------------------
	Function Poll_Results(oContent)
		dim pollID		: pollID = getAttribute(oContent, "id", "")
		dim choice : choice = getParam(pollID&"_choice")

		''' Insert the new vote if this poll is submited with an answer, and user has not already voted
		if len(choice)>0 then 'and len(request.cookies("poll_"&pollID))=0 then
			if oContent.SelectNodes("choice[@id='"&choice&"']").length=1 then
				oContent.SelectSingleNode("choice[@id='"&choice&"']").attributes.getnameditem("count").value = cstr(clng(oContent.SelectSingleNode("choice[@id='"&choice&"']").attributes.getnameditem("count").value) + 1)
				oContent.OwnerDocument.save DATA_FOLDER & POLLS_FOLDER & pollID & XMLFILE_EXTENSION
				response.cookies("poll_"&pollID) = "1"
				response.cookies("poll_"&pollID).expires = now+30
			end if
		end if

		'-- define the template source
		Dim oTemplate
		Set oTemplate = new AspTemplate
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\poll\"
		oTemplate.Template = "poll_results.html"
			
		oTemplate.Slot("poll_question") = GetChild(oContent, "question", "")
		oTemplate.Slot("skin_name") = g_sSkin
		
		oTemplate.clearBlock "poll_item_block"
		
		'loop on choice
		Dim oChoice, iVote, oList
		Dim totalvote : totalvote = 0
		Set oList = oContent.SelectNodes("choice")
		
		for each oChoice in oList
			totalvote = totalvote + getAttribute(oChoice, "count", "0")
		next

		for each oChoice in oList
			if totalvote=0 then 
				iVote = 0
			else
				iVote = round(getAttribute(oChoice, "count", "")/totalvote, 2)*100
			end if
			
			oTemplate.Slot("choice_vote") = cstr(iVote)
			oTemplate.Slot("choice_text") = getAttribute(oChoice, "value", "")
			oTemplate.RepeatBlock "poll_item_block"
		Next
		
		oTemplate.Slot("total_votes") = totalvote
		
		Poll_Results = oTemplate.GetOutput		
		set oTemplate = Nothing	
	End Function
%>