<%
	'-------------------------------------------
	'-- Display Fullxml Update welcome screen --
	'-------------------------------------------
	Sub webform_fullxmlupdate()					
		Response.Write "<h3>" & String("system", "fullxmlupdates", "welcomefullxmlupdate") & "</h3>"
		Response.Write "<p>" & String("system", "fullxmlupdates", "fullxmlupdatemsg") & "</p>"
		Response.Write "<img src=engine/admin/media/update.png align=middle>&nbsp;<a href=admin.asp?webform=webform_scan_updates>" & String("system", "fullxmlupdates", "scanupdates") & "</a>"
	End Sub
	
	
	'-------------------------------------------
	'-- Scan fullxml.com for availables updates
	'-------------------------------------------
	Sub webform_scan_updates()
		Dim oFullxmlUpdate
		Set oFullxmlUpdate = CreateDomDocument
		
		Response.Write "<h3>" & String("system", "fullxmlupdates", "welcomefullxmlupdate") & "</h3>"
		
		If oFullxmlUpdate.LoadXML(GetHttp("http://www.fullxml.com/fullxml.updates.xml", "fullxml.updates.xml", 0)) then
			Dim i : i =0
			Dim oApp
			Dim d_version
			Dim l_version : l_version = getAttribute(g_oLocalVersion.documentelement, "major", "0")*1000 + getAttribute(g_oLocalVersion.documentelement, "minor", "0")*100 + getAttribute(g_oLocalVersion.documentelement, "revision", "0")
						
			'-- check each application node to search for a newer version
			For each oApp in oFullxmlUpdate.SelectNodes("applications/application[@name='fullxml4']")
				d_version = getAttribute(oApp, "major", "0")*1000 + getAttribute(oApp, "minor", "0")*100 + getAttribute(oApp, "revision", "0")
				
				if d_version>l_version then
					Response.Write "<b>" & getChild(oApp, "friendlyname", "") & "</b><br>"
					Response.Write String("system", "fullxmlupdates", "downloadsize") & " " & int(getChild(oApp, "size", "")/1024) & String("system", "fullxmlupdates", "kilobyte") & "<br>" 
					Response.Write getChild(oApp, "description", "") & "<br>"
					Response.Write "<a href=""" & getChild(oApp, "path", "") & """ target=_blank>" & String("system", "fullxmlupdates", "startdownload") & "</a><br>"					
					Response.Write "<p>"
					i = i + 1
				end if
							
			Next
			
			if i=0 then Response.Write "<p>" & String("system", "fullxmlupdates", "noupdate")
			
		Else
			Response.Write "<br>" & String("system", "fullxmlupdates", "fullxmlupdatedown")
		End if
		
		
		'-- link to older versions
		Response.Write "<br><a href=admin.asp?webform=webform_list_UpdatesHistory>" & String("system", "fullxmlupdates", "formerversions") & "</a>"		
	End sub
	
	
	'------------------------------------------
	'-- Display every former fullxml version --
	'------------------------------------------
	Sub webform_list_UpdatesHistory()
		Dim oFullxmlUpdate
		Set oFullxmlUpdate = CreateDomDocument
		
		Response.Write "<h3>" & String("system", "fullxmlupdates", "welcomefullxmlupdate") & "</h3>"
		
		If oFullxmlUpdate.LoadXML(GetHttp("http://www.fullxml.com/fullxml.updates.xml", "fullxml.updates.xml", 0)) then
			Dim i : i =0
			Dim oApp
						
			'-- check each application node to search for a newer version
			For each oApp in oFullxmlUpdate.SelectNodes("applications/application[@name='fullxml4']")
				
				Response.Write "<b>" & getChild(oApp, "friendlyname", "") & "</b><br>"
				Response.Write String("system", "fullxmlupdates", "downloadsize") & " " & int(getChild(oApp, "size", "")/1024) & String("system", "fullxmlupdates", "kilobyte") & "<br>" 
				Response.Write getChild(oApp, "description", "") & "<br>"
				Response.Write "<a href=""" & getChild(oApp, "path", "") & """ target=_blank>" & String("system", "fullxmlupdates", "startdownload") & "</a><br>"					
				Response.Write "<p>"
			Next
		Else
			Response.Write "<br>" & String("system", "fullxmlupdates", "fullxmlupdatedown")
		End if
	End Sub
%>