<%
	
	'CONST DOMDOCUMENT_PROGID				= "MSXML2.DOMDocument.3.0"
	'CONST FreeThreadedDOMDOCUMENT_PROGID	= "MSXML2.FreeThreadedDOMDocument.3.0"
	'CONST XSLTEMPLATE_PROGID				= "MSXML2.XSLTemplate.4.0"

	'==============================
	'Shortcut and example function
	'==============================
	Function Transform(pXMLSource, pXSLSource)
		Dim oTrans 
		set oTrans = new CTransform
		
		oTrans.DebugMode = false
		Call oTrans.Sources(pXMLSource, pXSLSource)		
		Call oTrans.Process()
		Transform = oTrans.Content
		set oTrans = nothing
	End Function
	
	
	Sub TransformShow(pXMLSource, pXSLSource)
		Dim oTrans 
		set oTrans = new CTransform
		oTrans.Sources pXMLSource, pXSLSource
		oTrans.show
		set oTrans = nothing
	End Sub
	
	
	Function TransformDebug(pXMLSource, pXSLSource)
		Dim oTrans 
		set oTrans = new CTransform
		
		oTrans.DebugMode = true
		Call oTrans.Sources(pXMLSource, pXSLSource)		
		Call oTrans.Process()
		TransformDebug = oTrans.Content & oTrans.Message
		set oTrans = nothing
	End Function
	
%>
<%
	'==============================================================================
	' Permet de faire une transformation entre des données XML et un template XSL.
	'	Les fonctions interressantes sont: la possibilité d'ajouter des parametres 
	'	(addParam) et l'ajout automatique du context querystring en tant que parameter 
	'	XSL.
	'	modified: 13/09/2002
	'==============================================================================
	Class CTransform

		'==============================================================================
		' Membres
		'==============================================================================
		Private m_strContent                                              
		Private m_strError                                                
		Private m_intError 
		Private m_sXMLContext   
		
		Private oXML
		Private oXSL   
		Private oXSLTemplate    
		Private oXSLProcessor                                    

		private m_bDebugMode

		' Constructeur
		Private Sub Class_Initialize
			Set oXML = server.CreateObject(DOMDOCUMENT_PROGID)
				oXML.async = false
				call oXML.setProperty ("ServerHTTPRequest", true)
				If DOMDOCUMENT_PROGID = "MSXML2.DOMDocument.4.0" then 
					Call oXML.setProperty("NewParser", True)
				End If
			
			Set oXSl = server.CreateObject(FreeThreadedDOMDOCUMENT_PROGID)
				oXSL.async = false
				call oXSL.setProperty ("ServerHTTPRequest", true)	
				
				If FreeThreadedDOMDOCUMENT_PROGID = "MSXML2.FreeThreadedDOMDocument.4.0" then 
					Call oXSL.setProperty("NewParser", True)
				End If
				
			Set oXSLTemplate  = server.createobject(XSLTEMPLATE_PROGID)
			
			m_bDebugMode = false
				
		End Sub
		
		' Destructeur
		Private Sub Class_Terminate
			Set oXML = Nothing
			Set oXSl = Nothing	    
		End Sub

		'==============================================================================
		' Property en Ecriture
		'==============================================================================

		Property Let DebugMode(bValue)
			m_bDebugMode = bValue
		End Property

		'==============================================================================
		' Property en Lecture
		'==============================================================================
		
		' Contains the result of the transformation
		Property Get Content() : Content = m_strContent : End Property	
		' Contains the XML data
		Property Get XML() : XML = oXML.xml : End Property
		' Contains the XSL template
		Property Get XSL() : XSL = oXSL.xml : End Property
		' Contains the error message
		Property Get Message() : Message = m_strError : End Property
		' Contains the debugmode
		Property Get DebugMode() : DebugMode = m_bDebugMode : End Property	

		'==============================================================================
		' Cette méthode permet d'affecter au processeur les sources XML et XSL
		' Elles peuvent être sous forme d'URL, de chemin relatif/absolu ou de Chaine
		'==============================================================================
		Public Sub Sources(sXMLSource, sXSLSource)
			Dim bLoad : bLoad=false
			
			
			
			'on charge le XML
			if mid(sXMLSource,1,1)="<" then
				debug "XML string"
				bLoad = oXML.loadXML(sXMLSource)				 
			elseif mid(sXMLSource,1,4)="http" then
				debug "XML http"
				bLoad = oXML.load(sXMLSource)
			elseif mid(sXMLSource,2,1)=":" then
				debug "XML file"
				bLoad = oXML.load(sXMLSource)
			else
				debug "XML relative file"
				bLoad = oXML.load(Server.Mappath(sXMLSource))
			end if
			
			if (oXML.parseError.errorCode <> 0) or bLoad=false  then
				m_strError = m_strError & "XML parseError on line " & oXML.parseError.line & ", descr: " & oXML.parseError & " " & oXML.parseError.reason & "<br>"
				debug m_strError
				Exit Sub
			End If
			
		
			'on charge le XSL
			if mid(sXSLSource,1,1)="<" then
				bLoad = oXSL.loadXML(sXSLSource)
			elseif mid(sXSLSource,1,4)="http" then
				bLoad = oXSL.load(sXSLSource)
			elseif mid(sXSLSource,2,1)=":" then
				bLoad = oXSL.load(sXSLSource)
			else
				bLoad = oXSL.load(Server.Mappath(sXSLSource))
			end if
			
			if (oXSL.parseError.errorCode <> 0) or bLoad=false then
				m_strError = m_strError & "XSL parseError on line " & oXSL.parseError.line & ", descr: " & oXSL.parseError & " " & oXSL.parseError.reason & "<br>"
				debug m_strError
			End If
			
			'tout a reussi, on crée le XSLTemplate, et le XSLProcessor afin de pouvoir ajouter des parameters
			Set oXSLTemplate.stylesheet = oXSL				
			Set oXSLProcessor = oXSLTemplate.createProcessor()
			oXSLProcessor.input = oXML		
			
			'response.Write (m_strError)
			
		End Sub


		'==============================================================================
		' Effectue la transformation et renvoie le resultat dans la propriété content
		'==============================================================================
		Public Sub Process()
			if lenb(m_strError)>0 then exit sub
			Call AddContext()
			oXSLProcessor.Transform
			m_strContent = oXSLProcessor.output	
			
			Set oXSLTemplate.stylesheet = nothing
			Set oXSLProcessor = Nothing
		End Sub

		'-- 
		Public Sub Show()
			if lenb(m_strError)>0 then exit sub
			Call AddContext()
			oXSLProcessor.output = Response
			oXSLProcessor.Transform
			
			Set oXSLTemplate.stylesheet = nothing
			Set oXSLProcessor = Nothing					
		End Sub
		
		
		'-- Save result in a file :: @todo
		Public sub Save(sFileName)
			
		end sub

		'==============================================================================
		' Add a parameter to the XSL
		'==============================================================================	
		Public Function AddParam(sParamName, sParamValue)
			oXSLProcessor.addParameter sParamName, sParamValue		
		End Function
	  
	  
		'==============================================================================
		' Automatically add the Querystring context as parameters into the XSL
		'==============================================================================	
		Private Sub AddContext()
  			Dim Item
		  	
  			For each Item in request.querystring
  				oXSLProcessor.addParameter Cstr(Item), Cstr(request.querystring(Item))
  			Next  	  	
		End Sub
		
		'==============================================================================
		' Print "sText" to screen if DebugMode = True
		'==============================================================================
		private function debug(sText)
			if m_bDebugMode then Response.Write sText & "<br>"
		end function
	  
	End Class
%>
