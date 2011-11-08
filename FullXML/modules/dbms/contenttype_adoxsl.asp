<!-- #include file="CAdoXml.asp" -->
<%
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Execute the insert/update of specific data of this modules
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub InsertUpdate_Dbms_AdoXsl(oContent)
	
		Call InsertUpdateExtraContent( oContent, array("connectionstring") )
		Call InsertUpdateExtraContent( oContent, array("sqlquery") )
		Call InsertUpdateExtraContent( oContent, array("xsltemplate") )
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: This function write the part of the INSERT/UPDATE FORM
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Sub Edit_Dbms_AdoXsl(oContent)
		Dim connectionstring
		Dim sqlquery
		Dim xsltemplate
								
		'-- We try to get the value in the case of an update
		If not isempty(oContent) Then 
			connectionstring	= GetChild(oContent, "connectionstring", "")
			sqlquery 			= GetChild(oContent, "sqlquery", "")
			xsltemplate			= GetChild(oContent, "xsltemplate", "")
		End If
		
		'-- Print The form
		With Response
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("dbms", "contenttype_adoxsl", "connectionstring") & "</th><td><input type=text class=large name=connectionstring value='" & connectionstring & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("dbms", "contenttype_adoxsl", "sqlquery") & "</th><td><input type=text class=large name=sqlquery value='" & sqlquery & "'></td></tr>"
			.Write "<tr class=datagrid_editrow valign=top><th>" & String("dbms", "contenttype_adoxsl", "xsltemplate") & "</th><td><input type=text class=large name=xsltemplate value='" & xsltemplate & "'></td></tr>"
		End WIth
		
	End Sub
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	':: Module rendering function
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function Render_Dbms_AdoXsl(oContent)
		Dim connectionstring
		Dim sqlquery
		Dim xsltemplate
		
		connectionstring	= replace(GetChild(oContent, "connectionstring", ""), "{{servermappath}}", g_sServerMappath)
		sqlquery 			= GetChild(oContent, "sqlquery", "")
		xsltemplate			= replace(GetChild(oContent, "xsltemplate", ""), "{{servermappath}}", g_sServerMappath)
		
		if Len(connectionstring)=0 and len(sqlquery)=0 or len(xsltemplate)=0 then
			Exit Function
		End If
			
	
		Dim oAdoXML
		set oAdoXML = new CAdoXML
			oAdoXML.DebugMode = False
			oAdoXML.ConnectionString = connectionstring
			oAdoXML.Query = sqlquery
			oAdoXML.XSLTemplate = xsltemplate
			Render_Dbms_AdoXsl = oAdoXML.Process
		set oAdoXML = nothing
		
	End Function
%>