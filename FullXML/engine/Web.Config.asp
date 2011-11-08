<%
	'-----------------------------------------
	'-- Load the website configuration file --
	'-----------------------------------------
	Public Sub LoadWebConfig(bForceUpdate)
				
		If bForceUpdate OR (not isObject(Application(g_sServerName & "_appSettings"))) Then
			'-- clean old datas
			Set Application(g_sServerName & "_appSettings") = Nothing
			Application(g_sServerName & "_appSettings") = Empty
			
			'-- Create a free threaded DOM and put the config file in it	
			Dim oXML
			Set oXML = CreateFreeDomDocument
			If NOT oXML.load(g_sServerMapPath & "\" & g_sServerName & ".config.xml.asp") Then
				If NOT oXML.load(g_sServerMapPath & "\web.config.xml.asp") Then
					LogIt "web.config.asp", "LoadWebConfig", FATAL, "No configuration file", oXML.Parseerror.reason
					FatalError "Fatal Error", "This website encounter a fatal error, please check back later."
				End If
			End If
				
			Set Application(g_sServerName & "_appSettings") = oXML	
		End If		
	End Sub
	
	
	'---------------------------------------------------------------------------------
	'-- Return the value of the appSettings Key, or 'empty' if the key is not found --
	'---------------------------------------------------------------------------------
	Public Function AppSettings(p_sKey)
		Dim oNodeList
		
		'-- load conf if not there
		LoadWebConfig False
		
		Set oNodeList = Application(g_sServerName & "_appSettings").SelectNodes("configuration/appSettings/key[@add='" & p_sKey & "']")
		
		If oNodeList.Length=1 Then
			AppSettings = oNodeList.Item(0).Attributes.GetNamedItem("value").Text
		Else
			AppSettings = empty
		End If
		
		
		Set oNodeList = Nothing
	End Function
%>