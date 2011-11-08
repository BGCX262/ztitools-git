<%
	'----------------------------------------
	'-- This Function create a DomDocument --
	'----------------------------------------
	Function CreateDomDocument
		Set CreateDomDocument = server.CreateObject(DOMDOCUMENT_PROGID)
					
		'-- Dom properties
		CreateDomDocument.async = false
		
		If DOMDOCUMENT_PROGID = "MSXML2.DOMDocument.4.0" then 
			CreateDomDocument.setProperty "NewParser", True
		End If
		
		CreateDomDocument.setProperty "SelectionLanguage", "XPath"
		CreateDomDocument.setProperty "ServerHTTPRequest", true
		CreateDomDocument.validateOnParse = false
		CreateDomDocument.resolveExternals = false
		CreateDomDocument.preserveWhiteSpace = True
		
		'-- Add the processing instruction
		Dim pi : Set pi = CreateDomDocument.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
		CreateDomDocument.appendChild(pi)
		
		'-- Adding asp comment to stop execution
		Dim oComment : Set oComment = CreateDomDocument.CreateComment(" <% Response.End %"&" >")
		CreateDomDocument.appendChild oComment
		
	End Function
	
	
	'------------------------------------------------------
	'-- This Function create a Free threaded DomDocument --
	'------------------------------------------------------
	Function CreateFreeDomDocument
		Set CreateFreeDomDocument = server.CreateObject(FreeThreadedDOMDOCUMENT_PROGID)
		
		'-- Dom properties
		CreateFreeDomDocument.async = false
		If FreeThreadedDOMDOCUMENT_PROGID = "MSXML2.FreeThreadedDOMDocument.4.0" then 
			CreateFreeDomDocument.setProperty "NewParser", True
		End If
		CreateFreeDomDocument.setProperty "SelectionLanguage", "XPath"
		CreateFreeDomDocument.setProperty "ServerHTTPRequest", true
		CreateFreeDomDocument.validateOnParse = false
		CreateFreeDomDocument.resolveExternals = false
		CreateFreeDomDocument.preserveWhiteSpace = True
		
		'-- Add the processing instruction
		Dim pi : Set pi = CreateFreeDomDocument.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
		CreateFreeDomDocument.appendChild(pi)
		
		'-- Adding asp comment to stop execution
		Dim oComment : Set oComment = CreateFreeDomDocument.CreateComment(" <% Response.End %"&" >")
		CreateFreeDomDocument.appendChild oComment
		
	End Function
	
	
	'-- Return the attribute value of an node
	Function GetAttribute(oNode, sName, sDefaultValue)
		on error resume next
		
		Dim oAttribute
		GetAttribute = sDefaultValue
		
		For each oAttribute in oNode.Attributes
			if (oAttribute.name = sName) Then
				GetAttribute = oAttribute.Value
				'Exit Function
			End If
		Next
		
		if err<>0 then
			response.write "<b>GetAttribute</b> :" & sName & "<br>"
			err.clear
		end if
		
		on error goto 0
		
	End Function
	
	
	'-- Sets the value of a node attribute
	Sub SetAttribute(oNode, sName, sValue)
		Dim oAttribute		
		For each oAttribute in oNode.Attributes
			if (oAttribute.name = sName) Then
				oAttribute.Value = sValue
				Exit Sub
			End If
		Next		
	End Sub
		
	
	'-- Return a child node value
	Function GetChild(oNode, sName, sDefaultValue)
		Dim oChild
		For each oChild in oNode.ChildNodes
			'if oChild.nodeType=1 then
				if (oChild.nodeName = sName) Then
					GetChild = oChild.text
					Exit Function
				End If
			'End If
		Next
		GetChild = sDefaultValue	
	End Function
	
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	'::: SetChildNodeValue, This function set the value of a node child.
	'	- pParentNode: the XMLNode
	'	- the type of the child to update ("cdata", "node", "attribute")
	'	- pName: the name of the child
	'	- pCreateIfNotExists: indicates wether or not the child has to be created if it is not there
	':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Function SetChildNodeValue(pParentNode, pType, pName, pValue, pCreateIfNotExists)
		Dim oNode
		
		Select case pType
			case "cdata"
				if pParentNode.SelectNodes(pName).Length=0 then
					set oNode = pParentNode.ownerDocument.CreateElement(cstr(pName))
					oNode.appendChild(pParentNode.ownerDocument.createCDATASection(cstr(pValue)))
					pParentNode.appendChild(oNode)
				Else
					pParentNode.SelectNodes(pName).item(0).firstchild.text = pValue
				End If
				
			case "node"
				set oNode = pParentNode.ownerDocument.CreateElement(cstr(pName))
				oNode.text = cstr(pValue)
				pParentNode.appendChild(oNode)
			
			case else
				'-- todo: create missing attribute
				set oNode = pParentNode.ownerDocument.CreateAttribute(cstr(pName))
				oNode.value = cstr(pValue)
				pParentNode.attributes.SetNamedItem(oNode)
				
		end select
	End Function
	
	
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	'-- This function displayt a datagrid from an xml file
	' todo: add paging
	'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
	Public Sub XmlDatagrid(sName, p_XMLData, sXPath, arrCols, arrFields, myaction, primarykey, primarykeyname, allowcreate)
		Dim i
		Dim oXML, oNodeList, oNode, sLink
		
		
		'-- work with dom or 
		If IsObject(p_XMLData) Then
			Set oXML = p_XMLData
		else
			Set oXML = CreateDomDocument
			p_XMLData = CheckFileName(p_XMLData)
			if not oXML.Load(p_XMLData) then
				LogIt "XmlUtilities.asp", "XmlDatagrid", ERROR, oXML.ParseError.Reason, oXML.url
				Exit Sub
			end if
		end if
		
		
		Dim oTemplate
		set oTemplate = new ASPTemplate
		oTemplate.TemplateDir = g_sServerMapPath & ADMIN_FOLDER & "templates/"
		oTemplate.Template = "datagrid.html"
		
		
		'-- caption
		oTemplate.Slot("caption") = sName
		
		
		'-- Header columns
		oTemplate.ClearBlock "HeadersColumns"
		For i=LBound(arrCols) to UBound(arrCols)
			oTemplate.Slot("header_rank") = i+1
			oTemplate.Slot("header_name") = arrCols(i)
			oTemplate.RepeatBlock "HeadersColumns"
		Next
				
		
		'-- Data rows
		oTemplate.ClearBlock "DatasRows"
		
		'-- Loop on the rows
		For each oNode in oXML.SelectNodes(sXPath)
			'-- link to edit the row
			if len(myaction)>0 then
            	sLink = g_sScriptName & "?webform=" & myaction & "&"&primarykeyname&"=" & oNode.attributes.getNamedItem(primarykey).Value
			end if
            oTemplate.Slot("editrow_link") = sLink
			
			'-- loop on each field
			oTemplate.ClearBlock "DatasColumns"
			For i=LBound(arrFields) to UBound(arrFields)
				if arrFields(i) = "password" then
					oTemplate.Slot("data_value") = "******"
				elseif i=0 and len(myaction)>0 Then
					oTemplate.Slot("data_value") = "<a href='" & sLink & "'>" & oNode.attributes.getNamedItem(arrFields(i)).Value & "</a>"
				else
					Dim tmp : tmp = getAttribute(oNode, arrFields(i), "")
					oTemplate.Slot("data_value") = iff(len(tmp)>0, tmp, "&nbsp;")  'oNode.attributes.getNamedItem(arrFields(i)).Value
				end if
				oTemplate.RepeatBlock "DatasColumns"
			Next
						
			oTemplate.RepeatBlock "DatasRows"
		Next
		
		'-- add the "new row" if needed
		if allowcreate then
			oTemplate.Slot("newrow") = "<tr class=datagrid_newrow><th>*</th><td colspan=50> <a href="""&g_sScriptName&"?webform=" & myaction & """>" & String("system", "common", "create") & "</a></td></tr>"
		end if
		
		oTemplate.Generate		
		Set oTemplate = Nothing
				
		if NOT IsObject(p_XMLData) Then Set oXML = Nothing
	End Sub
	
	

	
	'''================================================================
	''' Create a SELECT from a querystring
	'''================================================================
	''' sListName : the html name of each CHECKBOX item
	''' sFieldVALUE : the bdd field used for the html value of each element
	''' sFieldLABEL : the bdd field displayed near each element
	''' sQuery : the SQL query used to retrive the list of element
	''' sSelectedValue : the list of selected VALUE, sperated by a coma . for example for example "2" or "2,4,78"
	'--------------------------------------------------------------------
	Function XMLListBox(sListName, sFieldVALUE, sFieldLABEL, sXMLPath, sXpath, sSelectedValue, arrayValue, arrayLabel)
		Dim oXML
		Dim oNodeList, oNode
		if IsNull(sSelectedValue) or LenB(sSelectedValue) = 0 then sSelectedValue = ""
		Dim arrSel : arrSel = split(sSelectedValue, ",")
		Dim j
		
		XMLListBox = ""
		
		if Not IsObject(sXMLPath) then
			set oXML = CreateDomDocument
			if not oXML.Load(sXMLPath) Then
				LogIt "XmlUtilities", "XMLListBox", ERROR, oXML.ParseError.reason, CheckFileName(sXMLPath)
				Exit Function
			End If
		Else
			Set oXML = sXMLPath
		End If
		
	'	if not oXML.load(CheckFileName(sXMLPath)) then
	'		LogIt "XmlUtilities", "XMLListBox", ERROR, oXML.ParseError.reason, CheckFileName(sXMLPath)
	'		Exit Function
	'	end if
		
		'-- execution of the query
		Set oNodeList = oXML.SelectNodes(sXPath)
		
		'--Start of the select
		XMLListBox = "<select name='" & sListName & "' id='" & sListName & "'>"
		
		If isarray(arrayValue) and IsArray(arrayLabel) then
			For j = LBound(arrayValue) to UBound(arrayValue)
				XMLListBox = XMLListBox & "<option value='" & arrayValue(j) & "'>&nbsp;" & arrayLabel(j) & "</option>" 
			Next
		End If
		
		'-- les npdes
		For each oNode in oNodeList
			XMLListBox = XMLListBox & "<option value='" & oNode.Attributes.GetNamedItem(sFieldVALUE).value & "'" 
			For j=lBound(arrSel) to uBound(arrSel) : If cstr( oNode.Attributes.GetNamedItem(sFieldVALUE).value)=cstr(trim(arrSel(j))) Then : XMLListBox = XMLListBox & " selected" : End If : Next
			XMLListBox = XMLListBox & ">&nbsp;" & oNode.Attributes.GetNamedItem(sFieldLABEL).value & "</option>"
		Next
		XMLListBox = XMLListBox & "</select>"
		
		set oNodeList = nothing
		'Set oXML = Nothing		
	End Function
	
	
'	'''================================================================
'	''' Create a SELECT from a querystring
'	'''================================================================
'	''' sListName : the html name of each CHECKBOX item
'	''' sFieldVALUE : the bdd field used for the html value of each element
'	''' sFieldLABEL : the bdd field displayed near each element
'	''' sQuery : the SQL query used to retrive the list of element
'	''' sSelectedValue : the list of selected VALUE, sperated by "," . for example for example "2" or "2,4,78"
'	'--------------------------------------------------------------------
'	Function XMLListBox(sListName, sFieldVALUE, sFieldLABEL, sXMLPath, sXpath, sSelectedValue, arrayValue, arrayLabel)
'		Dim oXML : set oXML = CreateDomDocument
'		Dim oNodeList, oNode
'		if IsNull(sSelectedValue) or LenB(sSelectedValue) = 0 then sSelectedValue = ""
'		Dim arrSel : arrSel = split(sSelectedValue, ",")
'		Dim j
'		
'		XMLListBox = ""
'		
'		'-- execution of the query
'		if not oXML.load(sXMLPath) then
'			XMLListBox = "can't load xml."
'			exit function
'		end if
'		
'		Set oNodeList = oXML.SelectNodes(sXPath)
'		
'		XMLListBox = "<select name='" & sListName & "' id='" & sListName & "'>"
'		
'		'-- array de valeurs statiques
'		'if lenb(sDefaultValue)>0 and lenb(sDefaultLabel)>0 then
'		'	XMLListBox = XMLListBox & "<option value='" & sDefaultValue & "'>" & sDefaultLabel & "</option>" 
'		'end if
'		
'		'-- les npdes
'		For each oNode in oNodeList
'			XMLListBox = XMLListBox & "<option value='" & oNode.Attributes.GetNamedItem(sFieldVALUE).value & "'" 
'			'on test s'il fait parti des selections
'			For j=lBound(arrSel) to uBound(arrSel) : If cstr( oNode.Attributes.GetNamedItem(sFieldVALUE).value)=cstr(trim(arrSel(j))) Then : XMLListBox = XMLListBox & " selected" : End If : Next
''			XMLListBox = XMLListBox & ">&nbsp;" & oNode.Attributes.GetNamedItem(sFieldLABEL).value & "</option>"
'		Next
'		XMLListBox = XMLListBox & "</select>"
'		
'		set oNodeList = nothing
'		Set oXML = Nothing		
'	End Function
	
	
	'------------------------------------------
	'-- Create an XML file if it not present --
	'------------------------------------------
	Sub CreateXmlFile(path, rootname)
		Dim oXML
		Set oXML = CreateDomDocument
		
		If Not g_oFso.FileExists(path) Then
			Dim oRoot
			Set oRoot = oXML.createElement(rootname)
			
			oXML.appendchild(oRoot)
			
			on error resume next
				oXML.Save(path)
				
				IF ERR<>0 Then 
					LogIt "XmlUtilities.asp", "CreateXmlFile", ERROR, "Can't save file " & path, Err.number & ": " & err.Description
					err.Clear
				End if
			on error goto 0
				
		Else
			LogIt "XmlUtilities.asp", "CreateXmlFile", INFO, "File already exists.", path
		End If
				
		Set oXML = nothing	
	End Sub
	
	
	'+---------------------------------------------------+
	'| Delete a node from an xml file
	'|---------------------------------------------------+
	'| sXMLPath : path to the xml file, or dom
	'| sXPath	: the xpath to the node(s) to be deleted 
	'+---------------------------------------------------+
	Sub DeleteNode (p_XML, p_sXPath)
				
		Dim l_oXML, oDeletedNodeList
		
		
		'-- If the 1st param is a path, load the File
		if Not IsObject(p_XML) then
			Set l_oXML = CreateDomDocument
			if not l_oXML.Load(p_XML) Then
				LogIt "xmlutitlities.asp", "DeleteNode", ERROR, l_oXML.ParseError.reason, p_XML
				Exit sub
			end if
		Else
			Set l_oXML = p_XML
		End If
		
		
		'-- Log
		LogIt "XmlUtilities.asp", "DeleteNode", INFO, "XML Nodes are deleted", l_oXML.url
				
		
		'-- Remove nodes that matching the xpath
		Set oDeletedNodeList = l_oXML.SelectNodes(p_sXPath)
		oDeletedNodeList.removeAll
		
		'-- Save the data
		l_oXML.Save(l_oXML.url)
				
		
		'-- release created object
		Set oDeletedNodeList = nothing
		
		'If Not IsObject(p_XML) then
		'	Set oXML = nothing
		'End If	
	End Sub
	
	
	'----------------------------------------------------------
	' Insert some attributes to node(s) into an xml document 
	'	Inputs:
	'	p_XML: xml path or object
	'   sXpath: Xpath to the node(s) to update 
	'	arrayNode: array of attribute's name
	'	arrayValue: array of attribute's value
	'	bAddAutoID: indicates if a autoid will be added
	'	insertBeforeXpath: specify a xpath, relative to sXpath param, to put node before
	'----------------------------------------------------------
	Function InsertNode (p_XML, sXPath, sNodeName, arrayNode, arrayValue, bAddAutoID, insertBeforeXpath)
		Dim oXML, oNode, oAtt, i, oRefNode, oRefNodeList
		
		if Not IsObject(p_XML) then
			Set oXML = CreateDomDocument
			if not oXML.Load(p_XML) Then
				LogIt "XMLUtilities.asp", "InsertNode", ERROR, oXML.url, oXML.parseError.reason
				Exit Function
			end if
			LogIt "XmlUtilities.asp", "UpdateNode", INFO, "A Node is inserted in : " & p_xml, sXPath
		Else
			Set oXML = p_XML
			LogIt "XmlUtilities.asp", "UpdateNode", INFO, "A Node is inserted in : " & p_xml.url, sXPath
		End If
		
		
		
		'-- create the new node		
		Set oNode = oXML.CreateElement(sNodeName)		
		
		'-- Add numauto
		if bAddAutoID then
			Set oAtt = oXML.createAttribute("id")
			InsertNode = GetGuid
			oAtt.Value = InsertNode
			oNode.Attributes.SetNamedItem(oAtt)
			Set oAtt = Nothing
		end if
		
		'-- Add attributes
		For i=lBound(arrayNode) to uBound(arrayNode)			
			Set oAtt = oXML.createAttribute(arrayNode(i)) 
			'if arrayNode(i)="password" then
			'	oAtt.Value = Md5(cstr(arrayValue(i)))	'Md5(GetParam(arrayValue(i)))
			'else
				oAtt.Value = cstr(arrayValue(i)) 'GetParam(arrayValue(i))
			'end if
			
			oNode.Attributes.SetNamedItem(oAtt)
			Set oAtt = Nothing
		Next
				
		'-- Append the new node to each node that match the Xpath
		Dim oNodeChild
		For each oNodeChild in oXML.SelectNodes(sXPath)
			
			'-- Case of a refNode :: use to insert a node at the right place
			If Len(insertBeforeXpath)>0 Then
				Set oRefNodeList = oXML.SelectNodes(insertBeforeXpath)
								
				'-- insert the node before specified one
				if oRefNodeList.length=0 then
					Call oNodeChild.insertBefore(oNode.CloneNode(true), null)
				else
					For each oRefNode in oRefNodeList
						Call oNodeChild.insertBefore(oNode.CloneNode(true), oRefNode)
					Next
				end if					
			Else
				'-- insert child at the end of the list
				Call oNodeChild.insertBefore(oNode.CloneNode(true), null)
			End If			
		Next
				
		'-- Save the data
		oXML.Save(oXML.url)
				
		Set oNode = nothing
	End Function
	
	
	'----------------------------------------------------------
	' Update some attributes of node(s) into an xml document 
	'	Inputs:
	'	p_XML: xml path or object
	'   sXpath: Xpath to the node(s) to update 
	'	arrayAttributeName: array of attribute's name
	'	arrayAttributeValue: array of attribute's value
	'----------------------------------------------------------
	Sub UpdateNode (p_XML, sXPath, arrayAttributeName, arrayAttributeValue)
		Dim oXML, oNode, i
		Set oXML = CreateDomDocument
		Dim tmpVal
	
		'-- the first parameter is either a path or a domdocument object
		'-- We load the file in both case
		if Not IsObject(p_XML) then
			Set oXML = CreateDomDocument
			if not oXML.Load(p_XML) Then
				LogIt "xmlutitlities.asp", "UpdateNode", ERROR, oXML.ParseError.reason, p_XML
				Exit sub
			end if
			LogIt "XmlUtilities.asp", "UpdateNode", INFO, "A Node is updated in " & p_xml, sXPath
		Else
			Set oXML = p_XML
			LogIt "XmlUtilities.asp", "UpdateNode", INFO, "A Node is updated in " & p_xml.url, sXPath
		End If
		
				
		'update all the nodes that matching the xpath
		For each oNode in oXML.SelectNodes(sXPath)
			For i=lBound(arrayAttributeName) to uBound(arrayAttributeName)
				
				'-- in a case of a field called PASSWORD then encrypt				
				'if arrayAttributeName(i)="password" then
				'	tmpVal = Md5(arrayAttributeValue(i))'Md5(GetParam(arrayAttributeValue(i)))
				'else
				'tmpVal = arrayAttributeValue(i)		'GetParam(arrayAttributeValue(i))
				'end if
				
				'-- we try to update the attribute, and if it fails we append it :-o
				on error resume next
				oNode.Attributes.GetNamedItem(arrayAttributeName(i)).Value = arrayAttributeValue(i)
				if err<>0 then
					err.Clear : on error goto 0
					Dim att
					set att = oXML.CreateAttribute(arrayAttributeName(i))
					att.value = arrayAttributeValue(i)
					oNode.Attributes.SetNamedItem(att)
				end if
				on error goto 0
			Next
		Next
		
		'save the data
		oXML.Save (oXML.url)
				
		Set oNode = nothing
		'Set oXML = nothing		
	End Sub
	
	
'	Function CreateXmlFile(path, rootname)
'		Dim oXML
'		Set oXML = CreateDomDocument
'		oXML.LoadXML ("<" & rootname & "/>")
'		oXML.Save path
'	End Function
	
	'-- append an attribute to a node
	Sub AppendAttribute(oNode, attName, attValue)
		Dim att
		set att = oNode.ownerDocument.CreateAttribute(attName)
		att.value = attValue
		oNode.Attributes.SetNamedItem(att)
	End Sub
	
	
	'
	
	
	'+---------------------------------------------------+
	'| Evaluate a XPath
	'|---------------------------------------------------+
	'| Inputs
	'|	sXMLPath : path to the xml file
	'|	sXPath	: the xpath to the node(s) to be moved 
	'|
	'|Ouputs:
	'|	The number of matching nodes
	'+---------------------------------------------------+
	Function XPathChecker(p_XML, sXPath)
		Dim oXML, oNodeList, i
		
		if Not IsObject(p_XML) then
			Set oXML = CreateDomDocument
			if not oXML.Load(p_XML) Then
				LogIt "XMLUtilities.asp", "XPathChecker", ERROR, oXML.url, oXML.parseError.reason
				Exit Function
			end if
		Else
			Set oXML = p_XML
		End If
		
		XPathChecker = oXML.SelectNodes(sXPath).Length
	End Function
	
	'+---------------------------------------------------+
	'| Move up a node from an xml file
	'|---------------------------------------------------+
	'| sXMLPath : path to the xml file
	'| sXPath	: the xpath to the node(s) to be moved 
	'+---------------------------------------------------+
	Sub MoveUpNode (p_XML, sXPath)
		Dim oXML, oNodeList, i
		Dim oNodeToMoveUp, oPreviousNode, oNewNode
		        		
        if Not IsObject(p_XML) then
			Set oXML = CreateDomDocument
			if not oXML.Load(p_XML) Then
				LogIt "XMLUtilities.asp", "MoveUpNode", ERROR, oXML.url, oXML.parseError.reason
				Exit Sub
			end if
		Else
			Set oXML = p_XML
		End If
        
		'get nodes that matching the xpath
		Set oNodeList = oXML.SelectNodes(sXPath)
		
		'check that the node is here
		if oNodeList.Length<>1 then exit sub
		
		set oNodeToMoveUp = oNodeList.item(0)
		set oPreviousNode = oNodeToMoveUp.previousSibling
		set oNewNode = oNodeToMoveUp.CloneNode(true)
		
		'switch
		'-- add the 
		call oNodeToMoveUp.parentNode.insertBefore(oNewNode, oPreviousNode)
		call oNodeToMoveUp.parentNode.replaceChild(oPreviousNode, oNodeToMoveUp)
		
		'save the data
		oXML.Save (oXML.url)
				
		Set oNodeList = nothing
		'Set oXML = nothing		
	End Sub
	
	
	'+---------------------------------------------------+
	'| Move down a node from an xml file
	'|---------------------------------------------------+
	'| sXMLPath : path to the xml file
	'| sXPath	: the xpath to the node(s) to be moved 
	'+---------------------------------------------------+
	Sub MoveDownNode (p_XML, sXPath)
		Dim oXML, oNodeList, i
		Dim oNodeToMoveDown, oNextNode, oBackupNode
		
        if Not IsObject(p_XML) then
			Set oXML = CreateDomDocument
			if not oXML.Load(p_XML) Then
				LogIt "XMLUtilities.asp", "MoveUpNode", ERROR, oXML.url, oXML.parseError.reason
				Exit Sub
			end if
		Else
			Set oXML = p_XML
		End If
		
		'get nodes that matching the xpath
		Set oNodeList = oXML.SelectNodes(sXPath)
		
		'check that the node is here
		if oNodeList.Length<>1 then exit sub
		set oNodeToMoveDown = oNodeList.item(0)
					
		set oNextNode = oNodeToMoveDown.nextSibling
		
		
		'-- check if there is a node after
	'	if oNextNode<>NULL then
		
			Set oBackupNode = oNextNode.CloneNode(true)
			
			'switch :
			call oNextNode.parentNode.insertBefore(oBackupNode, oNextNode)
			call oNextNode.parentNode.replaceChild(oNodeToMoveDown, oNextNode)
			
			'save the data
			oXML.Save (oXML.url)
		
	'	end if
				
		Set oNodeList = nothing
		'Set oXML = nothing		
	End Sub
%>