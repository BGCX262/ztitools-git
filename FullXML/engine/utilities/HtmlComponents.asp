<%
	' THIS FILE CONTAINS VARIOUS HIGH LEVEL HTML FORM COMPONENTS
	
	
	'-- Create a listbox -----------------------------------------------------------------------------------
	' Inputs:
	'		- sControlName: name of the control
	'		- sControlValue: pre-selected value
	'		- arrayLabel: array that contains the Labels
	'		- arrayValue: array that contains the values
	'-------------------------------------------------------------------------------------------------------
	Function HtmlComponent_Select(sControlName, sControlValue, arrayLabel, arrayValue)
		Dim i, t
		
		t = t & "<select name='" & sControlName & "'>"		
		for i = LBound(arrayLabel) to UBound(arrayLabel)					
			t = t & "<option value='" &  arrayValue(i) & "'" & IFF(arrayValue(i) = sControlValue, " SELECTED", "") & ">" 
			't = t & getString(arrayLabel(i))
			t = t & arrayLabel(i)
			t = t & "</option>"
		next
		t = t & "</select>"
		
		HtmlComponent_Select = t
	End Function
	
	
	
	
	'-- Display a control that allow number selection ------------------------------------------------------
	' Inputs:
	'		- sFormName: the name of the parent form
	'		- sControlName: name of the control
	'		- sControlValue: pre-selected value
	'		- iMinRange: the lower authorized value
	'		- iMaxRange: the higher authorized value
	'-------------------------------------------------------------------------------------------------------
	Function HtmlComponent_Number(sFormName, sControlName, sControlValue, iMinRange, iMaxRange)
		Dim t
		t = t & "<input type=text class=small name='" & sControlName & "' value='" & sControlValue & "'>"
		t = t & " <a href='#' onclick=""if (window.document.forms['" & sFormName & "'].elements['" & sControlName & "'].value > " & iMinRange & ") window.document.forms['" & sFormName & "'].elements['" & sControlName & "'].value--;""><b>-</b></a>"
		t = t & " / "
		t = t & " <a href='#' onclick=""if (window.document.forms['" & sFormName & "'].elements['" & sControlName & "'].value < " & iMaxRange & ") window.document.forms['" & sFormName & "'].elements['" & sControlName & "'].value++;""><b>+</b></a>"
		HtmlComponent_Number = t
	End Function
	
	
	
	
	'-- Display a control that allow to choose between true and false --------------------------------------
	'		- sFormName: the name of the parent form
	'		- sControlName: name of the control
	'		- bControlValue: preselected value
	'-------------------------------------------------------------------------------------------------------
	Function HtmlComponent_Bool(sFormName, sControlName, bControlValue)
		Dim index, t
		if len(bControlValue)>0 then
			if cBool(bControlValue) then : index = 0 : else : index = 1 : end if
		else
			index = 1
		end if
		t = t & "<input type=Radio name='" & sControlName & "' id='" & sControlName & "_true' value='true'><label for='" & sControlName & "_true'>" & String("system", "common", "true") & "</label>"
		t = t & "<input type=Radio name='" & sControlName & "' id='" & sControlName & "_false' value='false'><label for='" & sControlName & "_false'>" & String("system", "common", "false") & "</label>"
		t = t & "<script>" & sFormName & "." & sControlName & "[" & index & "].checked = true</script>"	
		
		HtmlComponent_Bool = t	
	End Function
	
	
	'-- Display a control that allow to choose between true and false --------------------------------------
	'		- sFormName: the name of the parent form
	'		- sControlName: name of the control
	'		- iControlValue: preselected value: 0: offline, 1: in progress, 2: online
	'-------------------------------------------------------------------------------------------------------
	Function HtmlComponent_PublicationState(sFormName, sControlName, iControlValue)
		
		'-- read the default publication state from the web.config
		
		HtmlComponent_PublicationState = "" &_
			"<input type=Radio name='" & sControlName & "' id='" & sControlName & "_0' value='0'><label for='" & sControlName & "_0'>" & String("system", "common", "false") & "</label>" &_
			"<input type=Radio name='" & sControlName & "' id='" & sControlName & "_1' value='1'><label for='" & sControlName & "_1'>" & String("system", "common", "inprogress") & "</label>" &_
			"<input type=Radio name='" & sControlName & "' id='" & sControlName & "_2' value='2'><label for='" & sControlName & "_2'>" & String("system", "common", "true") & "</label>" &_
			"<script>" & sFormName & "." & sControlName & "[" & iControlValue & "].checked = true</script>"	
	End Function
	
	
	'-- Display an input bow for url, with a browse link ---------------------------------------------------
	'-------------------------------------------------------------------------------------------------------
	Function HtmlComponent_URL(sName, sValue)
		HtmlComponent_URL = "<input type=text class=large name=" & sName & " value='" & sValue & "'> <a href=# onclick=""window.open(document.all." & sName & ".value,'browse','');"">" & String("system", "common", "browse") & "</a>" 
	End Function
	
	
'	Function HtmlComponent_TextEditor(sName, sValue, sType)
'		Dim t : t = ""
'		
'		Select Case sType
'			case "small":
'				t = t & "<textarea name='" & sName & "' class=small>" & sValue & "</textarea>"
'				
'			case "medium":
'				t = t & "<textarea name='" & sName & "' class=medium>" & sValue & "</textarea>"
'			
'			case "large":
'				t = t & "<textarea name='" & sName & "' class=large>" & sValue & "</textarea>"
'		End Select
'		
'		't = t & "<textarea name='" & sName & "' cols=40 row=10>" & sValue & "</textarea>"
'		
'		HtmlComponent_TextEditor = t
'	End Function
	
	
	Sub HtmlComponent_TextEditor(sName, sValue, width, height)
		'-- on utilise un editeur FCK 
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.Value = sValue
		oFCKeditor.CreateFCKeditor sName, width, height
		Set oFCKeditor = Nothing
	End Sub
	
	
	
	
	'-- composant HTML de selection de DateTime
	'-- le parametre dDate doit etre au format YYYYMMDDHHNN
	Function HtmlComponent_DateTime(sName, sDate)
		Dim i, t : t = ""
		Dim arrayMonth
		arrayMonth = Array("Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre")
		
		Response.Write "<input type=hidden name=" & sName & " value='" & sDate & "'>"
		
		
		'-- liste des jours
		t = t & "<select name='" & sName & "_day' onChange='Build" & sName & "()'><option>dd</option>"
		for i = 1 to 31 
			t = t & "<option value='" & right("0" & i, 2) & "'"		
			If mid(sDate, 7, 2) = right("0" & i, 2) Then : t = t & " selected" : End if
			t = t & ">" & right("0" & i, 2) & "</option>"
		Next
		t = t & "</select>"
		
		
		'-- liste des mois		
		t = t & "<select name=" & sName & "_month onChange='Build" & sName & "()'><option>MM</option>"
		for i = 1 to 12 
			t = t & "<option value=" & right("0" & i, 2)
			If mid(sDate, 5, 2) = right("0" & i, 2) Then : t = t & " selected" : End If
			t = t & ">" & arrayMonth(i-1) & "</option>"
		Next
		t = t & "</select>"
		
		
		'-- liste des années		
		t = t & "<select name=" & sName & "_year onChange='Build" & sName & "()'><option>YYYY</option>"
		for i = year(date()) to year(date())+5 
			t = t & "<option value=" & i 
			if mid(sDate, 1, 4) = cstr(i) then
				t = t & " selected"
			end if
			t = t & ">" & i & "</option>"
		Next
		t = t & "</select>&nbsp;&nbsp;&nbsp;"
		
		
		'-- liste des heures		
		t = t & "<select name=" & sName & "_hour onChange='Build" & sName & "()'><option>hh</option>"
		for i = 0 to 23
			t = t & "<option value=" & right("0" & i, 2)
			If mid(sDate,9,2) = right("0" & i, 2) Then : t = t & " selected" : End If
			t = t & ">" & right("0" & i, 2) & "</option>"
		Next
		t = t & "</select>"
		
		
		'-- liste des heures		
		t = t & "<select name=" & sName & "_minute onChange='Build" & sName & "()'><option>mm</option>"
		
		for i = 0 to 59
			t = t & "<option value=" & right("0" & i, 2)
			if mid(sDate, 11, 2) = right("0" & i, 2) Then : t = t & " selected" : End If
			t = t & ">" & right("0" & i, 2) & "</option>"
		Next
		t = t & "</select>"
		
		'-- Function javascript qui permet de recreer le hidden a partir des different select
		t = t & "<script language=javascript>" &_
				"function Build" & sName & "(){" &_
					"document.all." & sName & ".value = " &_
					"+ document.all." & sName &"_year.options[document.all." & sName &"_year.options.selectedIndex].value" &_
					"+ document.all." & sName &"_month.options[document.all." & sName &"_month.options.selectedIndex].value" &_
					"+ document.all." & sName &"_day.options[document.all." & sName &"_day.options.selectedIndex].value" &_
					"+ document.all." & sName &"_hour.options[document.all." & sName &"_hour.options.selectedIndex].value" &_
					"+ document.all." & sName &"_minute.options[document.all." & sName &"_minute.options.selectedIndex].value ;" &_
					"/*alert(document.all." & sName & ".value);*/" &_
				"}</script>"
		
		'-- return value
		HtmlComponent_DateTime = t
		
	End Function
	
	
	'-----------------------------------------------
	'-- Display a selectbox for image align
	'-----------------------------------------------
	Private Function HtmlComponent_ImageAlign(name, value)
		Dim arr : arr = array(" ", "center", "left", "right", "top", "bottom", "middle", "absmiddle")
		Dim i		
		
		HtmlComponent_ImageAlign = "<select name=" & name & ">"
		
		For i=lbound(arr) to ubound(arr)
			HtmlComponent_ImageAlign = HtmlComponent_ImageAlign & "<option value='" & arr(i) & "'" & IFF(value=arr(i), " selected", "") & ">" & arr(i) & "</option>" 
		Next
		
		HtmlComponent_ImageAlign = HtmlComponent_ImageAlign & "</select>"
	End Function
	
	
	'-----------------------------------------------
	'-- Display a selectbox for image align
	'-----------------------------------------------
	Private Function HtmlComponent_ImageVAlign(name, value)
		Dim arr : arr = array(" ", "top", "middle", "bottom")
		Dim i		
		
		HtmlComponent_ImageVAlign = "<select name=" & name & ">"
		
		For i=lbound(arr) to ubound(arr)
			HtmlComponent_ImageVAlign = HtmlComponent_ImageVAlign & "<option value='" & arr(i) & "'" & IFF(value=arr(i), " selected", "") & ">" & arr(i) & "</option>" 
		Next
		
		HtmlComponent_ImageVAlign = HtmlComponent_ImageVAlign & "</select>"
	End Function
	
	
	'-----------------------------------------------
	'-- Display a selectbox for link target
	'-----------------------------------------------
	Private Function HtmlComponent_LinkTarget(name, value)
		Dim arr : arr = array("_self", "_parent", "_blank", "_top")
		Dim i
		
		HtmlComponent_LinkTarget = "<select name=" & name & ">"
		
		For i=lbound(arr) to ubound(arr)
			HtmlComponent_LinkTarget = HtmlComponent_LinkTarget & "<option value='" & arr(i) & "'" & IFF(value=arr(i), " selected", "") & ">" & arr(i) & "</option>" 
		Next
		
		HtmlComponent_LinkTarget = HtmlComponent_LinkTarget & "</select>"
	End Function
	
	
	
	
%>