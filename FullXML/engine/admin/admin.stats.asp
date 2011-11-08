<%
	Sub webform_webtraffic_settings
		Response.Write ". black listed IP<br>"
		Response.Write ". main referer list<br>"
	End Sub


	'--------------------
	Sub webform_webtraffic_summary
		Dim oConn, oRS, file
		Dim filename 
		Dim pageviews
		Dim totalusers
		Set oConn = server.CreateObject("adodb.connection")
		oConn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & DATA_FOLDER & STATS_FOLDER & ";Extensions=asc,csv,txt;FIL=text"
		'oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DATA_FOLDER & STATS_FOLDER & ";Extended Properties='text;FMT=Delimited;HDR=YES;'"
		oConn.Open
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.CursorType = 3
		oRS.LockType = 1
		
		
		'-- DISPLAY THE SUMMARY
		Response.Write "<h3>" & String("system", "webtraffic", "webtraffic") & " :: " & String("system", "webtraffic", "summary") & "</h3>"
		Response.Write "<table class=datagrid cellpadding=0 cellspacing=0>"
		Response.Write "<caption>" & String("system", "webtraffic", "webtrafficsummarydesc") & "</caption>"		
		Response.Write "<tr class=datagrid_column><th>&nbsp;</th><th>" & String("system", "webtraffic", "date") & "</th><th>" & String("system", "webtraffic", "pageviews") & "</th><th>" & String("system", "webtraffic", "totalusers") & "</th></tr>"
		
		'-- each stat file
		For each file in g_oFSO.getfolder(DATA_FOLDER & STATS_FOLDER).files
			if right(file.name, 3) = "csv" then
				filename = file.name 
				'-- page view
				oRS.Open "SELECT count(*) FROM [" & filename & "]", oConn
				pageviews = oRS(0)
				oRS.Close
				
				'-- total users
				oRS.Open "SELECT ipaddress FROM [" & filename & "] GROUP BY ipaddress", oConn
				totalusers = oRS.RecordCount
				oRS.Close
								
				Response.Write "<tr class=datagrid_row><th>&nbsp;</th><td>" & left(filename, instr(1,filename, ".")-1) & "</td><td>" & pageviews & "</td><td>" & totalusers & "</td></tr>"
				
			end if
		next
		
		Response.Write "</table>"
		
		oConn.Close
		Set oRs = Nothing
		set oConn = nothing
	End Sub
	
	
	'----------------------
	'-- Pages statistics --
	'----------------------
	Sub webform_webtraffic_pages
		Dim oConn, oRS
		Dim filename
		Dim pageviews
		Dim totalusers
				
		Set oConn = server.CreateObject("adodb.connection")
		oConn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & DATA_FOLDER & STATS_FOLDER & ";Extensions=asc,csv,txt;Persist Security Info=False;"
		oConn.Open
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.CursorType = 3
		oRS.LockType = 1
		
		'-- filename
		filename = iff(len(getParam("date"))>0, getparam("date"), Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2)) & ".csv"
		
		
		'-- query
		oRS.Open "SELECT top 50 pageid, count(*) FROM [" & filename & "] group by pageid ORDER BY count(*) desc", oConn
		
				
		'-- DISPLAY THE SUMMARY
		Response.Write "<h3>" & String("system", "webtraffic", "webtraffic") & " [" & String("system", "webtraffic", "pages") & "] " & left(filename, instr(1,filename, ".")-1) & "</h3>"
		Response.Write "<table class=datagrid cellpadding=0 cellspacing=0>"
		Response.Write "<caption>" & String("system", "webtraffic", "mostpopularpages") & "</caption>"		
		Response.Write "<tr class=datagrid_column><th>&nbsp;</th><th>" & String("system", "common", "name") & "</th><th>" & String("system", "webtraffic", "hits") & "</th></tr>"
		
		Dim pagename, oNodeList
		
		Do While not oRS.Eof
			Set oNodeList = g_oWebSiteXML.documentelement.SelectNodes("//page[@id='" & oRS(0) & "']")
			
			if oNodeList.Length>0 then
				pagename = getAttribute(oNodeList.item(0), "name", "")
			else
				pagename = oRS(0)
			end if
			Response.Write "<tr class=datagrid_row><th>&nbsp;</th><td>" & pagename & "</td><td>" & cstr(oRS(1)) & "</td></tr>"
		oRS.movenext
		Loop
		Response.Write "</table>"
		
		Response.Write HTMLComponents_SelectDateStats
		
		
		oRS.Close
		oConn.Close
		Set oRs = Nothing
		set oConn = nothing
	End Sub
	
	
	'----------------------
	'-- Users statistics --
	'----------------------
	Sub webform_webtraffic_users
		Dim oConn, oRS
		Dim filename
		Dim pageviews
		Dim totalusers
		Set oConn = server.CreateObject("adodb.connection")
		oConn.ConnectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & DATA_FOLDER & STATS_FOLDER & ";Extensions=asc,csv,txt;Persist Security Info=False;"
		oConn.Open
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.CursorType = 3
		oRS.LockType = 1
		
		'-- filename
		filename = iff(len(getParam("date"))>0, getparam("date"), Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2)) & ".csv"
		
		
		'-- query
		oRS.Open "SELECT top 50 ipaddress, count(*) FROM [" & filename & "] group by ipaddress ORDER BY count(*) desc", oConn
		'oRS.Open "SELECT * FROM [" & filename & "]", oConn
		
				
		'-- DISPLAY THE SUMMARY
		Response.Write "<h3>" & String("system", "webtraffic", "webtraffic") & " [" & String("system", "webtraffic", "users") & "] " & left(filename, instr(1,filename, ".")-1) & "</h3>"
		Response.Write "<table class=datagrid cellpadding=0 cellspacing=0>"
		Response.Write "<caption>" & String("system", "webtraffic", "mostactiveusers") & "</caption>"		
		Response.Write "<tr class=datagrid_column><th>&nbsp;</th><th>" & String("system", "webtraffic", "ipaddress") & "</th><th>" & String("system", "webtraffic", "hits") & "</th></tr>"
				
		Do While not oRS.Eof
			Response.Write "<tr class=datagrid_row><th>&nbsp;</th><td>" & oRS(0) & "</td><td>" & cstr(oRS(1)) & "</td></tr>"
		oRS.movenext
		Loop
		Response.Write "</table>"
		
		Response.Write HTMLComponents_SelectDateStats
		
		oRS.Close
		oConn.Close
		Set oRs = Nothing
		set oConn = nothing
	End Sub
	
	
	'--------------------------
	'-- Languages statistics --
	'--------------------------
	Sub webform_webtraffic_languages
		Dim oConn, oRS
		Dim filename
		Dim pageviews
		Dim totalusers
		Set oConn = server.CreateObject("adodb.connection")
		oConn.ConnectionString = "Driver={Microsoft Text-Treiber (*.txt; *.csv)};Dbq=" & DATA_FOLDER & STATS_FOLDER & ";Extensions=asc,csv,txt;Persist Security Info=False;"
		oConn.Open
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.CursorType = 3
		oRS.LockType = 1
		
		'-- filename
		filename = iff(len(getParam("date"))>0, getparam("date"), Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2)) & ".csv"
		
		
		'-- query
		oRS.Open "SELECT top 50 language, count(*) FROM [" & filename & "] group by language ORDER BY count(*) desc", oConn
		
				
		'-- DISPLAY THE SUMMARY
		Response.Write "<h3>" & String("system", "webtraffic", "webtraffic") & " [" & String("system", "webtraffic", "languages") & "] " & left(filename, instr(1,filename, ".")-1) & "</h3>"
		Response.Write "<table class=datagrid cellpadding=0 cellspacing=0>"
		Response.Write "<caption>" & String("system", "webtraffic", "mostusedlanguages") & "</caption>"		
		Response.Write "<tr class=datagrid_column><th>&nbsp;</th><th>" & String("system", "webtraffic", "language") & "</th><th>" & String("system", "webtraffic", "hits") & "</th></tr>"
		
		Do While not oRS.Eof
			Response.Write "<tr class=datagrid_row><th>&nbsp;</th><td>" & oRS(0) & "</td><td>" & cstr(oRS(1)) & "</td></tr>"
		oRS.movenext
		Loop
		Response.Write "</table>"
		
		Response.Write HTMLComponents_SelectDateStats
		
		oRS.Close
		oConn.Close
		Set oRs = Nothing
		set oConn = nothing
	End Sub
	
	
	'-------------------------
	'-- Browsers statistics --
	'-------------------------
	Sub webform_webtraffic_browsers
		Dim oConn, oRS
		Dim filename
		Dim pageviews
		Dim totalusers
		Set oConn = server.CreateObject("adodb.connection")
		oConn.ConnectionString = "Driver={Microsoft Text-Treiber (*.txt; *.csv)};Dbq=" & DATA_FOLDER & STATS_FOLDER & ";Extensions=asc,csv,txt;Persist Security Info=False;"
		oConn.Open
		
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.CursorType = 3
		oRS.LockType = 1
		
		'-- filename
		filename = iff(len(getParam("date"))>0, getparam("date"), Year(now) & "-" & Right("0" & Month(now), 2) & "-" & Right("0" & Day(now), 2)) & ".csv"
		
		
		'-- query
		oRS.Open "SELECT top 50 browser, count(*) FROM [" & filename & "] group by browser ORDER BY count(*) desc", oConn
		
				
		'-- DISPLAY THE SUMMARY
		Response.Write "<h3>" & String("system", "webtraffic", "webtraffic") & " [" & String("system", "webtraffic", "browsers") & "] " & left(filename, instr(1,filename, ".")-1) & "</h3>"
		Response.Write "<table class=datagrid cellpadding=0 cellspacing=0>"
		Response.Write "<caption>" & String("system", "webtraffic", "mostusedbrowsers") & "</caption>"		
		Response.Write "<tr class=datagrid_column><th>&nbsp;</th><th>" & String("system", "webtraffic", "browser") & "</th><th>" & String("system", "webtraffic", "hits") & "</th></tr>"
		
		Dim pagename, oNodeList
		
		Do While not oRS.Eof
			set oNodeList = g_oWebSiteXML.documentelement.SelectNodes("//page[@id='" & oRS(0) & "']")
			if oNodeList.Length>0 then
				pagename = getAttribute(oNodeList.item(0), "name", "")
			else
				pagename = oRS(0)
			end if
			Response.Write "<tr class=datagrid_row><th>&nbsp;</th><td>" & pagename & "</td><td>" & cstr(oRS(1)) & "</td></tr>"
		oRS.movenext
		Loop
		Response.Write "</table>"
		
		Response.Write HTMLComponents_SelectDateStats
		
		oRS.Close
		oConn.Close
		Set oRs = Nothing
		set oConn = nothing
	End Sub
	
	
	'--------------------------------------------
	'-- Dsiplay a select of the available days --
	'--------------------------------------------
	Function HTMLComponents_SelectDateStats()
		Dim file
		
		HTMLComponents_SelectDateStats = "<p>" & String("system", "webtraffic", "selectadate") &_
			": <select id=date name=date onchange=""document.location='admin.asp?webform="&g_sWebform&"&date=' + this.options[this.selectedIndex].value ;"">"
		
		HTMLComponents_SelectDateStats = HTMLComponents_SelectDateStats & "<option value=>----------</option>"
		For each file in g_oFSO.getfolder(DATA_FOLDER & STATS_FOLDER).files
			if right(file.name, 3) = "csv" then
				dim tmp : tmp = left(file.name, instr(1,file.name, ".")-1)
				HTMLComponents_SelectDateStats = HTMLComponents_SelectDateStats & "<option value="&tmp & iff(Request.QueryString("date")=tmp, " selected", "") & ">" & tmp & "</option>"
			end if
		Next
		HTMLComponents_SelectDateStats = HTMLComponents_SelectDateStats & "</select>"
	End Function
	
%>