<%
	'---------------------------------------------
	' This class handle the current logged user --
	'---------------------------------------------
	Class CUser		
		'private m_bCanUserEdit	' indicates wether current user can edit the page
		
		private m_iPagePermissionLevel
		'private m_CanUpload
		'private m_UploadLimit
				
		'-----------------
		'-- Constructor --
		'-----------------
		Private Sub Class_initialize
			'-- User information is loaded only the first time
			'-- Singleton pattern
			If len(Session(g_sServerName & "_loaded"))=0 then								
				Session(g_sServerName & "_IP") = Request.ServerVariables("REMOTE_ADDR")
				Culture = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
				set Session(g_sServerName & "_xmldoc") = CreateFreeDomDocument
				Session(g_sServerName & "_xmldoc").LoadXml("<user group='anonymous' login='' screenname='' culture='' />")
				Session(g_sServerName & "_loaded") = "loaded"
				
				Session(g_sServerName & "_placeholders") = iff(Request.Cookies("fx4")("placeholders")="on", true, false)
				
								
			End If
			
			'-- Switch the wysiwyg mode
			If lcase(request("placeholders"))="on" Then
				Response.Cookies("fx4")("placeholders") = "on"				
				ViewPlaceholders = True				
			
			ElseIf lcase(request("placeholders"))="off" Then
				Response.Cookies("fx4")("placeholders") = "off"
				ViewPlaceholders = False
			End If
						
						
			'-- get the permission level of this user, for the current page
			m_iPagePermissionLevel = readUserPermission("//page[@id='"&g_sPageID&"']", Login, Group)
		End sub
		
		
		'------------------------------------------------
		'-- Update cookie expiration date to one month --
		'------------------------------------------------
		Private Sub Class_Terminate
			response.cookies("fx4").expires = now + 30
		End Sub
		
		
		Public Sub Reset
			
			'clean old if any
			Set Session(g_sServerName & "_xmldoc") = Nothing
			Session.Abandon
			
			Session(g_sServerName & "_IP") = Request.ServerVariables("REMOTE_ADDR")
			Culture = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
			set Session(g_sServerName & "_xmldoc") = CreateFreeDomDocument
			Session(g_sServerName & "_xmldoc").LoadXml("<user group='anonymous' email='' screenname='' culture='' />")
			Session(g_sServerName & "_loaded") = "loaded"			
			
		End Sub
		
		
		'--------------------------
		'-- Return the sessionID --
		'--------------------------
		Public property get SID 
			Sid = Session.SessionID
		End Property
		
		
		'-----------
		'-- Login --
		'-----------
		Public default Property  get  Login
			Login = getAttribute(XmlDoc.documentelement, "email", "")
		End Property
		Public Property let Login(sValue)
			setAttribute XmlDoc.documentelement, "email", sValue
		End Property
		
		
		'--------------
		'-- Password --
		'--------------
		Public Property get Password
			Password = getAttribute(XmlDoc.documentelement, "password", "")
		End Property
		Public Property let Password(sValue)
			setAttribute XmlDoc.documentelement, "password", sValue
		End Property
		
		
		'--------------------------
		'-- the user screen name --
		'--------------------------
		Public Property  get  ScreenName
			ScreenName = getAttribute(XmlDoc.documentelement, "screenname", "")
		End Property
		Public Property let ScreenName(sValue)
			setAttribute XmlDoc.documentelement, "screenname", sValue
		End Property
		
		
		'--------------------
		'-- The user group --
		'--------------------
		Public Property get Group
			'TODO : Set the default group as a website option
			Group = getAttribute(XmlDoc.documentelement, "group", "anonymous")
		End Property
		Public Property let Group(sValue)
			setAttribute XmlDoc.documentelement, "group", sValue
		End Property
		
		
		'----------------------------
		'-- is Wysiwyg activated ? --
		'----------------------------
		Public Property Get ViewPlaceholders
			if Session(g_sServerName & "_placeholders") then
				ViewPlaceholders = True
			else
				ViewPlaceholders = False
			End iF
		End Property
		Public Property let ViewPlaceholders(bValue)
			Session(g_sServerName & "_placeholders") = bValue
		End Property
		
		
		'--------------------------
		'-- the user DOM --
		'--------------------------
		Public Property  get XmlDoc
			Set XmlDoc = Session(g_sServerName & "_xmldoc")
		End Property
		Public Property let XmlDoc(oValue)
			Set Session(g_sServerName & "_xmldoc") = oValue
		End Property
				
		'-----------------------------
		'-- Return the user culture --
		'-----------------------------
		Public Property get Culture
			Culture = Session(g_sServerName & "_culture")
		End Property
		
		
		'---------------------------------------------------------
		'-- Return the level of permission for the current page --
		'---------------------------------------------------------
		Public Property Get PagePermissionLevel
			PagePermissionLevel  = m_iPagePermissionLevel
		End Property
		
		
		'-----------
		'-- Can Upload --
		'-----------
		Public Property  get  CanUpload
			if len(Session(g_sServerName & "_canupload"))=0 Then
				
				Dim oXml : Set oXml = CreateDomDocument
				if oXML.Load(groups_xml) then
					Dim oGroup
					Set oGroup = oXML.SelectSingleNode("/groups/group[@id='"&Group&"']")
					Session(g_sServerName & "_canupload") = cbool(getAttribute(oGroup, "canupload", false ))
				else
					Session(g_sServerName & "_canupload") = false
				end if
			end if
			CanUpload = cbool(Session(g_sServerName & "_canupload"))
		End Property
				
	
	'	Public Property get Debug 
	'		Debug = Session(g_sServerName & "_debug")
	'	End Property
		
		
		
				
		
		
		'--------------------------
		'-- Affect the culture setting
		'-- Session [ _originalculture] = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
		'-- Session [ _culture] = the real language that will be used
		Public Property let Culture(sValue)
						
			'-- set the browser language 
			Session(g_sServerName & "_originalculture")=sValue
			
			
			'-- set the used language (defined at website level)				
			Session(g_sServerName & "_culture") = g_sCulture
			
			'-- Set the lcid
			on error resume next
			Session.LCID = g_iLCID
			if err<>0 then
				LogIt "CUser.asp", "Culture", ERROR, "Can't switch to specified lcid. You may have to add it in your regional settings (into windows control panel)", g_iLCID
				Session.LCID = 1033
				err.clear
			end if
			on error goto 0
			
			
			'-- set the used language (defined at website level)
'			Dim oLanguageList
'			Dim sCulture : sCulture = getAttribute(g_oWebSiteXML.documentElement, "culture", "en")
'			
'			Set oLanguageList = g_oLanguagesXML.SelectNodes("languages/language[@id='" & sCulture & "']")
'								
'			If oLanguageList.Length=1 then
'				
'				Session(g_sServerName & "_culture") = getAttribute(oLanguageList.item(0), "file", "en")
'				on error resume next
'				Session.LCID = cint(getAttribute(oLanguageList.item(0), "lcid", "1033"))
'				if err<>0 then
'					LogIt "CUser.asp", "Culture", ERROR, "Can't switch to specified lcid. You may have to add it in your regional settings (into windows control panel)", getAttribute(oLanguage.item(0), "lcid", "1033")
'					Session.LCID = 1033
'					err.clear
'				end if
'				on error goto 0
'			Else
'				Session(g_sServerName & "_culture") = "en-us"
'				Session.LCID = 1033
'			End If
					
			
			
			'-- set the lcid
'			Dim oXML, oLanguage
'			Set oXML = CreateDomDocument		
'					
			'-- try to load the languages list
			'-- @todo: freethreaded
'			if not oXML.Load(languages_path) then
'				LogIt "CUser.asp", "Culture", ERROR, oXML.parseError.reason, oXML.url
'				Session(g_sServerName & "_culture") = "en-us"
'				Session.LCID = 1033
			'-- loaded!
'			else
'				Set oLanguage=oXML.SelectNodes("languages/language[@id='" & sValue & "']")
'								
'				if oLanguage.Length=1 then
'					Session(g_sServerName & "_culture") = getAttribute(oLanguage.item(0), "file", "en")
'					on error resume next
'					Session.LCID = cint(getAttribute(oLanguage.item(0), "lcid", "1033"))
'					if err<>0 then
'						LogIt "CUser.asp", "Culture", ERROR, "Can't switch to specified lcid", getAttribute(oLanguage.item(0), "lcid", "1033")
'						Session.LCID = 1033
'						err.clear
'					end if
'					on error goto 0
'				else
'					Session(g_sServerName & "_culture") = "en-us"
'					Session.LCID = 1033
'				end if
'			end if
'			Set oXML = nothing
		End Property
			
		
		'-----------------------------------------------------------------------------
		'-- Used to see if user have the requested access level on the current page --
		'-----------------------------------------------------------------------------
		Public Function isGranted(p_RequestedAccessLevel)
			If Group="administrator" or Group="webmaster" then
				isGranted = true
			ElseIf m_iPagePermissionLevel >= p_RequestedAccessLevel Then
				isGranted = true
			Else
				isGranted = False
			End If
		End Function
		
	End Class
%>