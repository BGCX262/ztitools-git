<%
	'--------------------------------------------
	'-- THIS CLASS REPRESENTs THE CURRENT PAGE --
	'--------------------------------------------
	Class CPage
		
		private m_tStart	' used to measure execution time
				
				
		'-- Constructor ------------------------------
		' Init timer
		' Check user authorisation to view this page
		'---------------------------------------------
		Private Sub class_initialize
			m_tStart = Timer()
			
			If NOT g_oUser.isGranted(CONST_ACCESS_LEVEL_VIEWER) Then
				FatalError String("system", "permissions", "accessdenied"), String("system", "permissions", "accessdeniedmsg")
			End IF
		End sub
		
		
		'------------------------------------------------
		'-- Release objects and display execution time --
		'------------------------------------------------
		Private Sub class_terminate
			Call LogVisit(int( (Timer()-m_tStart)*1000))
			Call Response.Write("<script>window.status = 'Server-side execution time: " & cstr(round(( Timer()-m_tStart)*1000, 2)) & " ms';</script>")			
		End Sub
		
		
		'------------------------------------------------
		'-- Render the page
		'-- @@todo: split into more functions
		'------------------------------------------------
		public sub Display
			Dim oPlaceHolder, placeholderID, width, height, i, sPHCnt
			Dim objMatches, match, submatch
			Dim oXML : Set oXML = CreateDomDocument
			
			'-- Load the current page skeleton
			Dim skeleton : skeleton = TemplateFileContent(SKINS_FOLDER & g_sSkin & "\templates\" & g_sTemplate)
			
			'-- Loop on each placeholder
			Dim g_oRegExp
			Set g_oRegExp = New RegExp
			g_oRegExp.IgnoreCase = True
			g_oRegExp.Global = True
						
			
			'-- Loop on each placeholder founded in the skeleton			
			g_oRegExp.Pattern = "(<placeholder[^\/>*].*/>)"
			Set objMatches = g_oRegExp.Execute(skeleton)
			For each Match in objMatches
				For each submatch in Match.SubMatches					
					if oXML.LoadXML(submatch) Then
							
							'-- this variable will hold the placeholder content --
							sPHCnt = ""
							
							'-- read the placeholder attributes
							width = GetAttribute(oXML.DocumentElement, "width", "100%")
							height = GetAttribute(oXML.DocumentElement, "height", "10px")
							placeholderID = GetAttribute(oXML.DocumentElement, "id", "")
														
							
							'-- add the div around the placeholder
							sPHCnt = sPHCnt & IFF(g_oUser.isGranted(CONST_ACCESS_LEVEL_MODERATOR), "<div class='placeholder"& iff(g_oUser.ViewPlaceholders, "", "_hide") &"' id='" & placeholderID & "' contextmenu='placeholdermenu' style='width: " & width & "; height:" & height & "; '  ondragenter='cancelEvent()' ondragover='cancelEvent()' ondrop=""InsertContent(this.id, '', '" & g_sPageID & "')"" ondblclick22=""PasteContent(this.id, '', '" & g_sPageID & "')"">", "")
							
							'-- Search for this placehodler in the page.xml, and render it
							For each oPlaceHolder in  g_oWebPageXML.documentElement.SelectNodes("/page/placeholders/placeholder[@id = '" & placeholderID & "']") 
								sPHCnt = sPHCnt &  RenderPlaceholder(oPlaceHolder)
							Next
														
							'-- close the div around the placeholder
							sPHCnt = sPHCnt & IFF(g_oUser.isGranted(CONST_ACCESS_LEVEL_MODERATOR), "</div>", "")
							
							'-- fill the skeleton
							Skeleton = Replace(Skeleton, submatch, sPHCnt)
												
					else
						Logit "CPage.asp", "Display", ERROR, "A placeholder tag is not wellformed.", submatch
					end if					
				Next
			Next

			Set oXML = Nothing
			
			
			'-- Print the page --
			With Response
				.Write "<HTML>"
				.Write HtmlHeader
				.write AdminTools
				.Write skeleton
				.Write "</HTML>"	
			End With
		End Sub
	
			
		':::::::::::::::::::::::::::::::::::::::::::::::::
		':: Rendering one placeholder
		':::::::::::::::::::::::::::::::::::::::::::::::::
		Private Function RenderPlaceholder(p_oPlaceHolder)
			Dim placeholder	: placeholder = GetAttribute(p_oPlaceHolder, "id", "")
			Dim hspace		: hspace = GetAttribute(p_oPlaceHolder, "hspace", "4")
			Dim pagesize	: pagesize = GetAttribute(p_oPlaceHolder, "pagesize", "10")
			Dim paging		: paging = GetAttribute(p_oPlaceHolder, "paging", false)
			Dim page		: page = iff(len(getparam("pa"))>0 and getparam("ph")=placeholder, getparam("pa"), 0)
			Dim iCount, iMin, iMax
			Dim oContentList, i, tmp, curdate
			curdate = YYYYMMDDHHNN(now)
			
			'-- get the list of contents
			Set oContentList = p_oPlaceHolder.SelectNodes("content[@published='true' and (@startdatetime <=" & curdate & ") and (@enddatetime='' or @enddatetime>" & curdate & " )]")
			
			'-- calculate variables for paging
			if paging then 
				iCount = (oContentList.Length / pagesize)
				if int(icount)<iCount then iCount = int(iCount)+1
				iMin = pagesize*page
				iMax = iff(oContentList.Length>iMin+pagesize, iMin+pagesize, oContentList.Length) - 1
			else
				iMin = 0
				iMax = oContentList.Length-1
			end if
						
			'-- render the contents
			For i=iMin to iMax
				Dim moveup, movedown, box, contentID, contentType, cachetimeout, contentName
				DIM arrCT


				if i<iMax then : movedown = "true" : else : movedown = "false" : end if
				if i>iMin then : moveup = "true" : else : moveup = "false" : end if
				
				contentID 	= getAttribute(oContentList(i), "id", "")
				contentName = getAttribute(oContentList(i), "name", "")
				contenttype = getAttribute(oContentList(i), "contenttype", "")
				cachetimeout = getAttribute(oContentList(i), "cachetimeout", "")
				box = getAttribute(oContentList(i), "box", "")
                arrCT = split(lcase(contenttype), "_")
				
				'response.write arrCT(0) & " - " & arrCT(1) & " :" & Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[@name='"&arrCT(0)&"']/contenttypes/contenttype[@name='"&arrCT(1)&"']").length & "<br>"
				
				Dim caching, boxing
				caching 	= cbool(GetAttribute(Application(APPVAR_DOM_MODULES).DocumentElement.SelectSingleNode("/modules/module[@name='"&arrCT(0)&"']/contenttypes/contenttype[@name='"&arrCT(1)&"']"), "caching", "false"))
				boxing 		= cbool(GetAttribute(Application(APPVAR_DOM_MODULES).DocumentElement.SelectSingleNode("/modules/module[@name='"&arrCT(0)&"']/contenttypes/contenttype[@name='"&arrCT(1)&"']"), "boxing", "false"))

				
				'-- open div content
				if g_oUser.isGranted(CONST_ACCESS_LEVEL_MODERATOR) then	RenderPlaceholder = RenderPlaceholder & "<div class='content' contextmenu='contentmenu' id='" & contentID & "' moveup='"&moveup&"' movedown='"&movedown&"' box='"&box&"'>"
							
				
				'-- get the content from the cache
				Dim reload : reload = true
				
				if USE_CACHE then
					Dim cachename : cachename = "content_" & contentID
					tmp = Application(cachename)
					dim cachedate : cachedate = iff(len(tmp)>=12, mid(tmp, 1, 12), "")
							
					If len(tmp)>0 and len(cachename)>0 and len(cachetimeout)>0 and len(cachedate)=12 then
						if int(YYYYMMDDHHNN(now)-cachedate) <= int(cachetimeout) then reload = false
					End if
				End If
				
				
				'-- Check the cases when cache is disabled
				If NOT caching or NOT USE_CACHE then reload = true
									
							
				'-- if true, execute render
				if reload then
					
					Execute ("tmp = Render_" & contenttype & "(oContentList.item(i))")

					' put abox around if boxing is allowed
					if boxing and len(box)>0 then
						tmp = RenderContent(contentName, tmp, box)
				    End IF
				    
				    '-- put module in cache
					if USE_CACHE Then
						Application(cachename) = cstr(YYYYMMDDHHNN(now)) & tmp
					end if
				Else
					'-- read from cache					
					tmp = mid(tmp, 13)
				end if
				
				
				'--put the data in the placeholder
				RenderPlaceholder = RenderPlaceholder & tmp 				
				
				
				'-- old method without cache				
				'Execute ("tmp = Render_" & oContentList.item(i).attributes.GetNamedItem("contenttype").Value & "(oContentList.item(i))")
				'RenderPlaceholder = RenderPlaceholder & RenderContent(oContentList(i).attributes.GetNamedItem("name").Value, tmp, oContentList(i).attributes.GetNamedItem("box").Value)
				
				
				'-- without boxxx
				''Execute ("RenderPlaceholder = RenderPlaceholder & Render_" & oContentList.item(i).attributes.GetNamedItem("type").Value & "(oContentList.item(i))")
				
				
				'-- Close the content DIV 
				if g_oUser.isGranted(CONST_ACCESS_LEVEL_MODERATOR) then	RenderPlaceholder = RenderPlaceholder & "</div>"
										
								
				'-- put a separator beetween content
				if i<oContentList.Length-1 and isnumeric(hspace) and hspace>0 then
					RenderPlaceholder = RenderPlaceholder & "<table cellpadding=0 cellspacing=0 border=0><tr><td height=" & hspace & "></td></table>"
				end if
			Next	
			
			'-- render the paging
			if paging and iCount>1 then
				
				RenderPlaceholder = RenderPlaceholder & String("system", "placeholders", "pages") & ": "
				
				For i=0 to iCount-1
					if i<>cint(page) then
						RenderPlaceholder = RenderPlaceholder & "<a href='" & g_sScriptName & "?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&ph=" & placeholder & "&pa=" & i & "' class=paging>" & i & "</a>"
					else
						RenderPlaceholder = RenderPlaceholder & "<b>" & i & "</b>"
					end if
					if i<> icount-1 then
						RenderPlaceholder = RenderPlaceholder & "  "
					end if
				next				
			End if
			
			Set oContentList = Nothing
		End Function
		
		
		'--------------------------------------------
		'-- Put a box around a title and a content --
		'--------------------------------------------
		Function RenderContent(sTitle, sContent, sBox)
			if len(sBox)=0 then
				RenderContent = sContent
			else
				Dim oTemplate
				Set oTemplate = new AspTemplate
				oTemplate.Debug = false
				oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\boxes\"
				oTemplate.Template = sBox
				oTemplate.Slot("title") = sTitle
				oTemplate.Slot("content") = sContent
				RenderContent = oTemplate.GetOutput
				Set oTemplate = Nothing
			end if
		End Function
		
					
		
		'---------------------------------------------------
		'-- Return the HTML head part of the current page --
		'---------------------------------------------------
		Function HtmlHeader()
			
			'-- first, set the metas
			HtmlHeader =	"<head>" &_
							"<TITLE>" & g_oWebPageXML.documentElement.SelectSingleNode("/page/metas/title").text & "</TITLE>" &_
							"<meta http-equiv=""content-type"" content=""text/html;charset=" & g_sEncoding  & """>" &_
							"<META NAME=""Description"" CONTENT=""" & g_oWebPageXML.documentElement.SelectSingleNode("/page/metas/description").text & """>" &_
							"<META NAME=""Keywords"" CONTENT=""" & g_oWebPageXML.documentElement.SelectSingleNode("/page/metas/keywords").text & """>"
			
			'-- Append each css of the current theme
			Dim oCss
								
			For each oCss in Application(APPVAR_DOM_SKINS).DocumentElement.SelectNodes("/skins/skin[@id='" & g_sSkin & "']/theme[@id='" & g_sTheme & "']/css")
				HtmlHeader = HtmlHeader & "<LINK HREF=""skins/" & g_sSkin & "/themes/" & g_sTheme & "/" & getAttribute(oCss, "id","") & """ REL=""stylesheet"" TYPE=""text/css"">"
			Next			
			
			HtmlHeader = HtmlHeader & "</head>"			
		End Function
				
		
		':::::::::::::::::::::::::::::::::::::::::
		':: Display the tools for page edition		
		':::::::::::::::::::::::::::::::::::::::::
		Public Function AdminTools
			Dim t
			
			'response.Write g_oUser.Group 'PagePermissionLevel
			
			
			'-- css and js libraries
			'If g_oUser.isGranted(CONST_ACCESS_LEVEL_CONTRIBUTOR) Then
			If g_oUser.group <> "anonymous" Then
				
				'-- Include the CSS-es
				t = t & "<LINK HREF='Engine/admin/templates/default/wysiwyg.css' REL='stylesheet' TYPE='text/css'>" & vbcrlf
				t = t & "<LINK HREF='Engine/admin/templates/default/contextmenu.css' REL='stylesheet' TYPE='text/css'>" & vbcrlf
				
				'-- include JS library
				t = t & "<scr" & "ipt language='JavaScript1.2' src='Engine/admin/contextmenu.js'></sc" & "ript>" & vbcrlf
				't = t & "<scr" & "ipt type='text/javascript' src='Engine/admin/drag.js'></sc" & "ript>" & vbcrlf
				
				'-- draw the user toolbar
				t = t & "<table class=admintoolbar cellpadding=0 cellspacing=0><tr><td  >"
					
				
				Select case g_oUser.PagePermissionLevel
					
					case CONST_ACCESS_LEVEL_ADMINISTRATOR
					
						'-- Console
						t = t & "<a title='Enter into the administrative console.' href=admin.asp>" & String("system", "interface", "goadmin") & "</a>&nbsp;"
						
						
						'-- page settings				
						t = t & "<a title='Edit the properties of the current page' href=admin.asp?webform=webform_update_page&mID=" & g_sMenuID & "&pID=" & g_sPageID & ">" & String("system", "webpages", "pagesettings") & "</a>&nbsp;"
									
											
						'-- Page contents
						t = t & "<a title='Display the list of content that belong to this page.' href=admin.asp?webform=webform_update_page_contents&pID="&g_sPageID&"&mID="&g_sMenuID&">" & String("system", "webpages", "pagecontents") &"</a>&nbsp;"
						
						'-- Page permissions
						t = t & "<a title='Display the list of authorisations on this page.' href=admin.asp?webform=webform_update_page_permissions&pID="&g_sPageID&"&mID="&g_sMenuID&">" & String("system", "webpages", "pagepermissions") &"</a>&nbsp;"
						
						
						'-- insert content
						t = t & "<a title='Insert a new content in this page.' href=""javascript: var p=window.open('popup.asp?webform=webform_insert_content&mID=" & g_sMenuID & "&pID=" & g_sPageID & "','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "insertcontent") & "</a>&nbsp;"
						
						
						'-- View placeholder
						If g_oUser.ViewPlaceholders then
							t = t & "<a title='Hide the placeholders highlight style.' href="& g_sScriptName &"?mID="&g_sMenuID&"&pID="&g_sPageID&"&placeholders=off><b>" & String("system", "interface", "placeholder_off") &"</b></a>&nbsp;"
						Else
							t = t & "<a title='Show the placeholders highlight style.' href="& g_sScriptName &"?mID="&g_sMenuID&"&pID="&g_sPageID&"&placeholders=on><b>" & String("system", "interface", "placeholder_on") &"</b></a>&nbsp;"
						End If
										
					Case CONST_ACCESS_LEVEL_CONTRIBUTOR
						g_oUser.ViewPlaceholders = false
									
						'-- my contents
						t = t & "<a title='Submit a content for this page' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_insert_submited_content','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "submitcontent") &"</a>&nbsp;"
						t = t & "<a title='View the list of my submited contents' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_list_submited_contents','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "mycontents") &"</a>&nbsp;"
								
					Case CONST_ACCESS_LEVEL_AUTHOR
						'g_oUser.ViewPlaceholders = false
									
						'-- my contents
						t = t & "<a title='Insert a content for this page' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_insert_content','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "insertcontent") &"</a>&nbsp;"
						t = t & "<a title='View the list of my authored contents' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_list_authored_contents','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "mycontents") &"</a>&nbsp;"
				
					Case CONST_ACCESS_LEVEL_MODERATOR
									
						'-- my contents
						t = t & "<a title='Insert a content for this page' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_insert_content','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "insertcontent") &"</a>&nbsp;"
						t = t & "<a title='View the list of my pending contents' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_list_pending_contents','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "pendingcontents") &"</a>&nbsp;"
						t = t & "<a title='View the list of all the contents of this page' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_list_contents','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "contents", "contents") &"</a>&nbsp;"
				
				
				End Select
				
				If g_oUser.CanUpload then				
					t = t & "<a title='Upload files' href=""javascript: var p=window.open('popup.asp?mID=" & g_sMenuID & "&pID=" & g_sPageID & "&webform=webform_System_FileExplorer','fx4popup', 'width=200,height=200, resizable=1, status=0');"">" & String("system", "tool_fileexplorer", "fileupload") &"</a>&nbsp;"
				End If
				
				t = t & "<a title='View your profile' href=""javascript:var reg = window.open('popup.asp?webform=webform_authentication_profile', 'profile', 'width=400, height=400');"">" & String("system", "common", "myprofile") &"</a>&nbsp;"
				t = t & "</td><td width=100 align=right><a href=default.asp?process=Do_Authentication_LogOff>" & String("system", "common", "logoff") &" <img src='engine/admin/media/close.png' align=middle border=0 alt='logoff'></a></td></tr></table>" & vbcrlf
				
				'-- USER TOOLBAR END

				
				t = t & "<body onload='fnInit();' onclick='fnDetermine()' contextmenu='pagemenu'>" & vbcrlf
				t = t & "<div status='false' onmouseover='fnChirpOn()' onmouseout='fnChirpOff()' id='oContextMenu' class='contextmenu'></div>" & vbcrlf
			
				'-- the definition of the contextuals menus
				t = t & "<xml id='contextDef'>"
				t = t & "	<xmldata>"
				
				t = t & "	<contextmenu id='pagemenu' pageid='" & g_sPageID & "'>"
				't = t & "		<item id='title0' value='Page edition' type='title'/>"
				t = t & "		<item id='menu1' value='Edit...' type='module' cmd='editpage'/>"
				t = t & "		<item id='menu1' value='View source' type='module' cmd='source'/>"
				t = t & "	</contextmenu>"
				
				t = t & "	<contextmenu id='placeholdermenu' pageid='" & g_sPageID & "'>"
				't = t & "		<item id='title0' value='" & String("system", "contents", "insertcontent") & "' type='title'/>"
				t = t & 			ModulesAsNodeList
				t = t & "	</contextmenu>"
				
				'-- the edit content menu is only available for moderator+
				If g_oUser.isGranted(CONST_ACCESS_LEVEL_MODERATOR) Then
							
					t = t & "	<contextmenu id='contentmenu' pageid='" & g_sPageID & "'>"
					't = t & "		<item id='title0' value='" & String("system", "contents", "contenttypes") & "' type='title'/>"
					t = t & "		<item id='content1' value='" & String("system", "common", "edit") & "' type='content' cmd='editcontent' icon='engine/admin/media/edit.png'/>"
					t = t & "		<item id='content2' value='" & String("system", "common", "moveup") & "' type='content' cmd='moveupcontent' icon='engine/admin/media/moveup.png'/>"
					t = t & "		<item id='content3' value='" & String("system", "common", "movedown") & "' type='content' cmd='movedowncontent' icon='engine/admin/media/movedown.png'/>"
					t = t & "		<item id='content4' value='-' type='separator' cmd=''/>"
					t = t & "		<item id='content5' value='" & String("system", "contents", "refresh") & "' type='content' cmd='refreshcontent' icon='engine/admin/media/refresh.png'/>"
					t = t & "		<item id='content6' value='" & String("system", "common", "delete") & "' type='content' cmd='deletecontent' icon='engine/admin/media/delete.png'/>"
					t = t & "		<item id='content7' value='-' type='separator' cmd=''/>"
					t = t & "		<item id='' value='" & String("system", "contents", "nobox") & "' type='content' cmd='changebox'/>"	
					t = t & 			BoxesAsNodeList
					t = t & "	</contextmenu>"
				
				End If
				
				t = t & "	</xmldata>"
				t = t & "</xml>" & vbcrlf
				
				
				'<!-- the form used by the page to post action -->
				t = t & "<table><form method=post name=myForm action='" & g_sURL & "'>"
				t = t & "	<input type=hidden name=process value=do_nothing ID=process>"
				t = t & "	<input type=hidden name=pageID id=pageID value=" & g_sPageID & ">"
				t = t & "	<input type=hidden name=boxID ID=boxID>"
				t = t & "	<input type=hidden name=contentID id=contentID>"
				t = t & "	<input type=hidden name=placeholder id=placeholder>"
				t = t & "	<input type=hidden name=nexturl id=nexturl value='" & g_sURL & "'>"
				t = t & "</form></table>" & vbcrlf
				
				AdminTools = t
				
				
				
				
			End If
			
		End Function
		
		
		'---------------------------------------------------------------------
		'-- Return the list of contenttype, grouped by modules as xml  nodes
		'-- this is used by the contextual menu
		'---------------------------------------------------------------------
		Function ModulesAsNodeList
			Dim oModuleList, oModule, modulenum : modulenum = 0
			Dim oContentType, oContentTypeList
			
			Set oModuleList = Application(APPVAR_DOM_MODULES).DocumentElement.SelectNodes("/modules/module[contenttypes/contenttype]")
			For Each oModule in oModuleList
				modulenum = modulenum + 1
				dim modulename : modulename = getAttribute(oModule, "name", "")

				Set oContentTypeList = oModule.SelectNodes("contenttypes/contenttype")
				
				'-- Loop on each content type of the current module
				For each oContentType in oContentTypeList
					dim ctname : ctname = getAttribute(oContentType, "name", "")
					ModulesAsNodeList = ModulesAsNodeList & "<item id=""" & modulename & "_" & ctname & """ module='" & modulename  & "' contenttype='" & ctname  & "' value=""" & String(modulename, "contenttypes", ctname) & """ type='placeholder' cmd='insertcontent'/>"
				Next
				
				'-- Add a separator if there is a following module
				if modulenum<oModuleList.length then
					ModulesAsNodeList = ModulesAsNodeList & "<item id='sep' value='' type='separator' cmd=''/>"
				End if							
			Next
		End Function
		
		
		'---------------------------------------------------------
		'-- Return the list of boxes as xml (4 contextual menu) --
		'---------------------------------------------------------
		Private Function BoxesAsNodeList
			Dim oSkin
			For each oSkin in Application(APPVAR_DOM_SKINS).DocumentElement.SelectNodes("/skins/skin[@id='" & g_sSkin & "']/box")
				BoxesAsNodeList = BoxesAsNodeList & "<item id='" & oSkin.Attributes.GetNamedItem("id").Value & "' value='" & RemoveExtension(oSkin.Attributes.GetNamedItem("id").Value) & "' type='content' cmd='changebox'/>"
			Next
		End Function
				
	End Class
%>