<% 'Response.Write AdoXml("Driver={SQL Server};server=DEVSQL1;uid=sa;password=somedev;database=modules", "procTInterviewsSelect", "adoxml.xsl") %>

<%
	
	' Shortcut function
	Function AdoXml(sConnectionString, sQuery, sXSLTemplate)
		Dim oTest
		set oTest = new CAdoXML
			oTest.DebugMode = False
			oTest.ConnectionString = sConnectionString
			oTest.Query = sQuery
			oTest.XSLTemplate = sXSLTemplate
			AdoXml = oTest.Process
		set oTest = nothing
	End Function


	'----------------------------------------------------------------------
	'
	'----------------------------------------------------------------------
	Class CAdoXML
	
		'members
		private m_sConnectionString
		private m_sQuery
		private m_sTemplate
		
		'debug
		Private m_bDebugMode
		Private m_tStart 		
		Private adPersistXML
		
		
		
		'----------------------------------------------------------------------
		'			CLASS's INTERNAL SUBS
		'----------------------------------------------------------------------				
		Private Sub Class_Initialize
			m_bDebugMode = false
			m_tStart = Timer()
			adPersistXML = 1		
		End Sub
		
		Private Sub Class_Terminate
			
			'debug cstr((Timer()-m_tStart)*1000) & "ms"
		End Sub
		
		
		'----------------------------------------------------------------------
		'			PROPERTIES
		'----------------------------------------------------------------------
		Property Let ConnectionString(sValue)
			m_sConnectionString = sValue
		End Property
		
		Property Let Query(sValue)
			m_sQuery = sValue
		End Property
		
		Property Let XSLTemplate(sValue)
			m_sTemplate = sValue
		End Property
		
		Property Let DebugMode(sValue)
			if lenb(sValue)>0 then : m_bDebugMode = true : Else : m_bDebugMode = false : End If
		End Property
		
		'----------------------------------------------------------------------
		'			PUBLIC METHODES
		'----------------------------------------------------------------------
		'(description=dsdsds')
		Public Function Process()
			Dim objCnn
			Dim objRst
			Dim objXSL
			Dim objXML
			Dim bLoad
			on error resume next
			if lenb(m_sConnectionString)=0 or lenb(m_sQuery)=0 or lenb(m_sTemplate)=0 then
				debug ("<b>Connection string:</b> " & m_sConnectionString & "<br>" & "<b>SQL Query:</b> " & m_sQuery & "<br>" & "<b>XSL Template:</b> " & m_sTemplate & "<br>")
				exit function
			end if
			
			Set objCnn = Server.CreateObject("ADODB.Connection")
			Set objXSL = Server.CreateObject("MSXML2.FreeThreadedDOMDocument.4.0") : objXSL.async = false
			Set objXML = Server.CreateObject("MSXML2.FreeThreadedDOMDocument.4.0") : objXML.async = false
			
		
			
			'we load the template
			'on charge le XSL
			if mid(m_sTemplate,1,1)="<" then
				bLoad = objXSL.loadXML(m_sTemplate)
			elseif mid(m_sTemplate,1,4)="http" then
				bLoad = objXSL.load(m_sTemplate)
			elseif mid(m_sTemplate,2,1)=":" then
				bLoad = objXSL.load(m_sTemplate)
			else
				bLoad = objXSL.load(Server.Mappath(m_sTemplate))
			end if
			
			if not bLoad then 
				debug("<b>XSL parseError</b><br>Line: " & objXSL.parseError.line & "<br>Description: " & objXSL.parseError & " " & objXSL.parseError.reason)	
				exit function
			end if	
			
			
			'-- open db						
			objCnn.Open m_sConnectionString
			
			set objRst = objCnn.Execute(m_sQuery)
			objRst.Save objXML,  adPersistXML
			Process = objXML.transformNode(objXSL)
			
			if err<>0 then
				Process = err.number & ": " & err.Description
				err.Clear
			end if
			
			'Close Database, recordsets, objects
			objRst.Close
			objCnn.Close

			Set objXSL = nothing
			Set objXML = nothing
			Set objCnn = nothing
			Set objRst = nothing			
		End Function
				
				
		'----------------------------------------------------------------------
		'			PRIVATE METHODES
		'----------------------------------------------------------------------
		private function debug(sText)
			if m_bDebugMode then Response.Write sText & "<br>"
		end function
	End Class
%>