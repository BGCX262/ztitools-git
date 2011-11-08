<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_Search_Results(oContent)
			
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_Search_Results(oNode)
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_Search_Results(oContent)
		Dim tStartSearch : tStartSearch = timer()
		Dim nbres : nbres = 0
		Dim s : s = getParam("s")
		
		Dim oTemplate
		Set oTemplate = new AspTemplate
		oTemplate.TemplateDir = SKINS_FOLDER & g_sSkin & "\modules\search\"
		oTemplate.Template = "search_results.html"
		
		oTemplate.Slot("search_results_title") = String("search", "contenttype_results", "search_results_title")
		oTemplate.Slot("search_button_label") = String("search", "contenttype_box", "search_button_label")
		oTemplate.Slot("search_value") = getparam("s")
		
		
		oTemplate.ClearBlock "search_item_block"

		if len(s)>2 then
			Dim oPageList
			Set oPageList = g_oWebSiteXML.selectNodes("website/menus/menu/page")
			For Each oPage in oPageList
				Dim l_pagepath
				l_pagepath = DATA_FOLDER & PAGES_FOLDER & getAttribute(oPage, "id", "") & XMLFILE_EXTENSION
				Dim oXML
				Set oXML = CreateDomDocument
				if oXML.Load( l_pagepath ) then
					Dim oResults, oResult
					Set oResults = oXML.SelectNodes("//content[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZéèà', 'abcdefghijklmnopqrstuvwxyzeea'), '"& lcase(s) &"')]")
					
					if oResults.length>0 then						
						oTemplate.Slot("search_item_name") =  getAttribute(oPage, "name", "")
						oTemplate.Slot("search_item_link") = g_sBaseUrl & "?pID="& getAttribute(oPage, "id", "")
						oTemplate.RepeatBlock "search_item_block"

						nbres = nbres + 1
					end if				
				end if
				Set oXML = Nothing	
								
			Next

			' no results
			if nbres=0 then
				oTemplate.Slot("search_results_message") = String("search", "contenttype_results", "noresults") & " """&s&"""."
			else
				oTemplate.Slot("search_results_message") = nbres & " " & String("search", "contenttype_results", "pagesfound") & " '" & s & "' [" & string("search", "contenttype_results", "searchtook") & " " & round(timer()-tStartSearch, 3) & " " & string("search", "contenttype_results", "seconds") & ".]"
			end if

		' to short search
		elseif len(s)>0 then
			oTemplate.Slot("search_results_message") = String("search", "contenttype_results", "search_too_short")
		
		' no search -> no message
		else
			oTemplate.Slot("search_results_message") = ""
		end if


		Render_Search_Results = oTemplate.GetOutput
		Set oTemplate = Nothing

	End Function
%>