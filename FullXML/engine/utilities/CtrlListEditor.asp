<%
	Class CtrlListEditor
		
		'-- for internal use
		private m_sScriptName
		private m_sAction
		private m_sItemID
		
		'-- User Input
		'private m_sCtrlName 
		private m_arrayLabel
		private m_arrayValue
		private m_iWidth
		private m_iHeight
		private m_sTitle
		private m_sBgColor
		
		private m_sAddUrl
		private m_sEdtUrl
		private m_sDelUrl
		private m_sMupUrl
		private m_sMdwUrl
		
		private m_AddButtonLabel
		private m_EditButtonLabel
		private m_DeleteButtonLabel
		private m_MoveUpButtonLabel
		private m_MoveDownButtonLabel
		
		private m_ConfirmDeleteWarning
		
		
		'-- title of the control
		public property Let Title(sTitle)		: m_sTitle = sTitle			: End Property
		
		'-- Size of the control
		public property Let Width(iWidth)		: m_iWidth = iWidth			: End Property
		public property Let Height(iHeight)		: m_iHeight = iHeight		: End Property
				
		'-- The arraies  that define the list
		public property Let Labels(arrayLabel)	: m_arrayLabel = arrayLabel : End Property
		public property Let Values(arrayValue)	: m_arrayValue = arrayValue : End Property
		
		'-- the asp function that will be called when the form is processed
		public property Let AddUrl(sValue)			: m_sAddUrl = sValue	: End Property
		public property Let DeleteUrl(sValue)		: m_sDelUrl = sValue	: End Property
		public property Let MoveUPUrl(sValue)		: m_sMupUrl = sValue	: End Property
		public property Let MoveDownUrl(sValue)		: m_sMdwUrl = sValue	: End Property
		public property Let EditUrl(sValue)			: m_sEdtUrl = sValue	: End Property
		
		'-- the button labels
		public property Let AddButtonLabel(sValue)		: m_AddButtonLabel = sValue		: End Property
		public property Let EditButtonLabel(sValue)		: m_EditButtonLabel = sValue	: End Property
		public property Let DeleteButtonLabel(sValue)	: m_DeleteButtonLabel = sValue	: End Property
		public property Let MoveUpButtonLabel(sValue)	: m_MoveUpButtonLabel = sValue	: End Property
		public property Let MoveDownButtonLabel(sValue)	: m_MoveDownButtonLabel = sValue: End Property
		
		'-- the delete warning message
		public property Let ConfirmDeleteWarning(sValue)	: m_ConfirmDeleteWarning = sValue	: End Property
		
		
		'-- Constructor
		private Sub Class_Initialize
			
			m_sTitle = "My editor title"
			m_iWidth = 640
			m_iHeight = 200
			m_sBgColor = "#8cade7"
			
			m_AddButtonLabel = "Add"
			m_EditButtonLabel = "Edit"
			m_DeleteButtonLabel = "Delete"
			m_MoveUpButtonLabel = "Move Up"
			m_MoveDownButtonLabel = "Move Down"
			
			m_ConfirmDeleteWarning = "Really delete this item ?"
			
			'- internal process
			m_sAction = Request.Form("action")
			m_sItemID = Request.Form("item_id")
			m_sScriptName = Request.ServerVariables("SCRITP_NAME") & "?" & Request.QueryString
		End Sub
		
		
		'-- render the Control
		Public Default Sub Display
						
			with Response
				RenderCSS
				
				.Write "<table height=" & m_iHeight & " bgcolor=" & m_sBgColor & " cellpadding=0 cellspacing=0>"
				
				.Write "<input type=hidden name=ctrlList_itemID id=ctrlList_itemID value="&m_sItemID&">"
				
				.Write "<tr><td colspan=2><span class=ctrltitle>" & m_sTitle & "</span></td></tr>"
				
				.Write "<tr valign=top>"
				.Write "	<td width=" & m_iWidth & ">"
								RenderList
				.Write "	</td>"
				.Write "	<td>" 
								RenderCrontrols
				.Write "	</td>"
				.Write "</tr>"
				.Write "</table>"
				RenderJavascript
			end with
		End Sub
		
				
		'-----------------
		' draw the list
		Private Sub RenderList
			Dim index
			with Response
				.Write "<select name=ctrlList_list id=ctrlList_list style='width: " & m_iWidth & "px; height: " & m_iHeight & "px' size=10 onchange='EventListChange();' ondblclick=""Do('edit');"">"
				
				for index = LBound(m_arrayLabel) to UBound(m_arrayLabel)
					.Write "<option value='" & m_arrayValue(index) & "'>" & m_arrayLabel(index) & "</option>"
				next
				
				.Write "</select>"
			end with
		End sub
		
		
		'-----------------
		' draw the buttons
		Private Sub RenderCrontrols
			with Response
				.Write "<table width=100% height=100% >"
				.Write "<tr valign=top><td>"
					.Write "<input type=button class=ctrlbutton id=ctrlList_add name=ctrlList_add value='" & m_AddButtonLabel & "' onClick=""Do('add');""><br><table cellpadding=0 cellspacing=0><tr><td height=4></td></tr></table>"
					.Write "<input type=button class=ctrlbutton id=ctrlList_edit name=ctrlList_edit value='" & m_EditButtonLabel & "' onClick=""Do('edit');"" disabled><br><table cellpadding=0 cellspacing=0><tr><td height=4></td></tr></table>"
					.Write "<input type=button class=ctrlbutton id=ctrlList_delete name=ctrlList_delete value='" & m_DeleteButtonLabel & "' disabled onClick=""if (confirm('" & m_ConfirmDeleteWarning & "')) Do('delete');""><br><br>"
				.Write "</td></tr>"
				.Write "<tr valign=bottom><td>"
					.Write "<input type=button class=ctrlbutton id=ctrlList_moveup name=ctrlList_moveup value='" & m_MoveUpButtonLabel & "' disabled onClick=""Do('moveup');""><br><table cellpadding=0 cellspacing=0><tr><td height=4></td></tr></table>"
					.Write "<input type=button class=ctrlbutton id=ctrlList_movedown name=ctrlList_movedown value='" & m_MoveDownButtonLabel & "' disabled onClick=""Do('movedown');""><br>"
				.Write "</td></tr></table>"
			end with
		End sub
		
		
		'-----------------
		' print the css
		Private Sub RenderCSS
			with Response
				.Write "<style>"
				.Write "input.ctrlbutton {width: 74px; height: 21px; font: messagebox;  border: 1px outset; }"
				.Write "span.ctrltitle {font: small-caption; padding-left: 4px;}"
				.Write "</style>"
			end with
		End sub
		
		
		'-----------------
		' print the javascript
		Private Sub RenderJavascript
			with Response
				.Write "<script>"
				
				
				'-- This function is executed when a liste element is clicked
				'-- we use it to update the buttons state and the 'itemID' value
				.Write "function EventListChange(){" & vbCrLf
				
				'--
				.Write "if (document.getElementById('ctrlList_list').options.selectedIndex>-1) document.getElementById('ctrlList_itemID').value = document.getElementById('ctrlList_list').options[document.getElementById('ctrlList_list').options.selectedIndex].value;" & vbCrLf
				
				'-- enable delete and edit
				.Write "document.getElementById('ctrlList_delete').disabled = false;" & vbCrLf
				.Write "document.getElementById('ctrlList_edit').disabled = false;" & vbCrLf
					
					'-- State of MOVEUP button 
					.Write " if (document.getElementById('ctrlList_list').options.length>1 && document.getElementById('ctrlList_list').options.selectedIndex>0  ) " & vbCrLf
					.Write " 		document.getElementById('ctrlList_moveup').disabled = false; " & vbCrLf
					.Write " else " & vbCrLf
					.Write " 		document.getElementById('ctrlList_moveup').disabled = true; " & vbCrLf
										
					'-- State of MOVEDOWN button
					.Write " if (document.getElementById('ctrlList_list').options.length>1 && document.getElementById('ctrlList_list').options.selectedIndex != (document.getElementById('ctrlList_list').options.length-1) ) " & vbCrLf
					.Write " 		document.getElementById('ctrlList_movedown').disabled = false; " & vbCrLf
					.Write " else " & vbCrLf
					.Write " 		document.getElementById('ctrlList_movedown').disabled = true; " & vbCrLf
									
				.Write "}" & vbCrLf
				
				'-- Fonction appellé lors du click sur le bouton add
				.Write "function Do(p_sAction){" & vbCrLf
					.Write "var itemID = document.getElementById('ctrlList_itemID').value;"
					.Write "	switch(p_sAction)"
					.Write "	{"
					.Write "		case 'edit':"
					.Write "			document.location = '"&m_sEdtUrl&"' + itemID;"
					.Write "			break;"
					.Write "		case 'add':"
					.Write "			document.location = '" & m_sAddUrl & "';"
					.Write "			break;"
					.Write "		case 'delete':"
					.Write "			document.location = '"&m_sDelUrl&"' + itemID;"
					.Write "			break;"
					.Write "		case 'moveup':"
					.Write "			document.location = '"&m_sMUpUrl&"' + itemID;"
					.Write "			break;"
					.Write "		case 'movedown':"
					.Write "			document.location = '"&m_sMDwUrl&"' + itemID;"
					.Write "			break;"
					.Write "	}"
										
				.Write "}" & vbCrLf
						
				.Write "</script>"
			end with
		End sub
		
		
	'	Private Sub ProcessAction
	'		select case m_sAction
	'			case "add":
	'				execute "Call " & m_sAddUrl & "(m_sID)"
	'				
	'			case "edit":
	'				execute "Call " & m_sEdtUrl & "(m_sID)"
	'				
	'			case "delete":
	'				execute "Call " & m_sDelFunction & "(m_sID)"
	'			
	'			case "moveup":
	'				execute "Call " & m_sMupFunction & "(m_sID)"
	'			
	'			case "movedown":
	'				execute "Call " & m_sMdwFunction & "(m_sID)"
	'				
	'		end select
	'		
	'		if len(m_sAction)>0 then
	'			Response.Redirect m_sScriptName
	'		end if
	'	End Sub
		
	End Class
	
	
'	'+-----------------------------------------------+
'	'| Sample use of the class
'	'+-----------------------------------------------+
'	
'	function fake(s_id)
'		Response.Write "fake function is working"
'	end function
'	
'	
'	Dim o
'	Set o = new CtrlListEditor
'	
'	o.Title = "External tools"
'	o.Labels = Array("Active x control", "texteditor", "xsl editor", "Spy++", "valid xml")
'	o.Values = Array("a", "b", "c", "d", "e")	
'	o.EditFunction = "fake"
'	o.AddFunction = "fake"
'	o.DeleteFunction = "fake"
'	o.MoveDownFunction = "fake"
'	o.MoveUPFunction = "fake"
'	o.Display	
'	
'	Set o = Nothing
%>