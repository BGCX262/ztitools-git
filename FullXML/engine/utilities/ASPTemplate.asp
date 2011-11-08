<script language="VBScript" runat="server">
' Copyright 2002 Sven Axelsson
class BlockInfo

	public Template
	public Show
	
	private data_array()
	private data_len
	
	private sub class_Initialize
		data_len = -1
		Show = false
	end sub
	
	public property get Data
		Data = data_array
	end property
	
	public property get HasData
		HasData = (data_len > -1)
	end property
	
	public sub AddData(value)
		data_len = data_len + 1
		redim preserve data_array(data_len)
		data_array(data_len) = value
	end sub
	
	public sub ClearData
		data_len = -1
		redim data_array(0)
	end sub
	
end class

class ASPTemplate

	public Debug
	
	private is_template_loaded
	private token_list, block_list, block_names
	private template_dir, output
	private fso, re
	private template_count

	private sub class_Initialize
		set token_list = CreateObject("Scripting.Dictionary")
		token_list.CompareMode = 1
		set block_list = CreateObject("Scripting.Dictionary")
		block_list.CompareMode = 1
		set block_names = CreateObject("Scripting.Dictionary")
		block_names.CompareMode = 1
		set fso = CreateObject("Scripting.FileSystemObject")
		set re = new regexp : re.IgnoreCase = true : re.Global = true
		template_dir = "templates/"
		Debug = false : is_template_loaded = false
	end sub
	
	private sub class_Terminate
		set token_list = nothing
		set block_list = nothing
		set block_names = nothing
		set re = nothing
		set fso = nothing
	end sub
	
	public property let Template(filename)
		dim path : path = template_dir & Replace(filename, "\", "/")
		if fso.FileExists(path) then
			dim f : set f = fso.OpenTextFile(path)
			re.Pattern = "[\r\n\t]+"
			output = re.Replace(f.ReadAll, " ")
			f.Close : set f = nothing
			template_count = 1
			output = ParseTemplate(output)
			is_template_loaded = true
		else
			ErrMsg "Can't find template file <b>" & path & "</b>."
		end if
	end property
	
	public property let TemplateDir(dirname)
		template_dir = Replace(dirname, "\", "/")
		if template_dir <> "" and not Right(template_dir, 1) = "/" then 
			template_dir = template_dir & "/"
		end if
	end property
	
	public property let Slot(name, value)
		if IsNull(value) then value = ""
		if token_list.Exists(name) then
			token_list.Item(name) = value
		else
			token_list.Add name, value
		end if
	end property
	
	public property get Slot(name)
		if token_list.Exists(name) then
			Slot = token_list.Item(name)
		else
			Slot = ""
		end if
	end property
	
	public property let ShowBlock(name, value)
		dim real_name
		for each real_name in Split(block_names.Item(name))
			block_list.Item(real_name).Show = value
		next
	end property
	
	public sub RepeatBlock(name)
		if UBound(Split(block_names.Item(name))) = 0 then
			with block_list.Item(block_names.Item(name))
				.Show = true
				.AddData GenerateHTML(.Template)
			end with
		else
			ErrMsg "Block names used with RepeatBlock must be unique.<br>Error on <b>" & name & "</b>"
		end if
	end sub
	
	public sub ClearBlock(name)
		if UBound(Split(block_names.Item(name))) = 0 then
			block_list.Item(block_names.Item(name)).ClearData
		else
			ErrMsg "Block names used with ClearBlock must be unique.<br>Error on <b>" & name & "</b>"
		end if
	end sub
	
	public sub Generate
	'	output = GenerateHTML(output)
	'	' Needed since we can't have nested comments in SGML/HTML
	'	re.Pattern = "(-->){2,}"
	'	Response.Write re.Replace(output, "-->")
		Response.Write GetOutput
	end sub
	
	public function GetOutput
		output = GenerateHTML(output)
		' Needed since we can't have nested comments in SGML/HTML
		re.Pattern = "(-->){2,}"
		GetOutput = re.Replace(output, "-->")
	end Function
	
	private function GenerateHTML(text)
		if is_template_loaded then
			dim re2 : set re2 = new regexp
			re2.IgnoreCase = true : re2.Global = true
			' Update all slots not included in a block
			text = UpdateSlots(text)
			' Update all shown blocks
			re2.Pattern = "{%{([a-z0-9_]+)}%}"
			dim matches : set matches = re2.Execute(text)
			dim match, name
			for each match in matches
				name = match.SubMatches(0)
				with block_list.Item(name)
					re2.Pattern = "{%{" & name & "}%}"
					if .Show then
						if .HasData then
							text = re2.Replace(text, Join(.Data, vbCrLf))
						else
							text = re2.Replace(text, GenerateHTML(.Template))
						end if
					end if
				end with
			next
			' Hide all unused blocks and slots
			if not Debug then
				re2.Pattern = "({%?{[a-z0-9_]+}%?})"
				text = re2.Replace(text, "<!" & "--$1-->")
			end if
			set re2 = nothing
		end if
		GenerateHTML = text
	end function
	
	private function ParseTemplate(text)
		' Find all blocks in the template text
		re.Pattern = "<!" & "--#BeginBlock\s+(\w+)\s*-->(.*?)" & _
			"<!" & "--#EndBlock\s+\1\s*-->"
		dim matches : set matches = re.Execute(text)
		dim match, block, name, decorated_name
		for each match in matches
			' Save the enclosed text in a BlockInfo structure
			set block = new BlockInfo
			name = match.SubMatches(0)
			decorated_name = AddTemplateName(name)
			block.Template = ParseTemplate(match.SubMatches(1))
			block_list.Add decorated_name, block
			' Replace the block with a marker {%{Name_Index}%}
			re.Pattern = "<!" & "--#BeginBlock\s+" & name & "\s*-->.*?" & _
				"<!" & "--#EndBlock\s+" & name & "\s*-->"
			re.Global = 0
			text = re.Replace(text, "{%{" & decorated_name & "}%}")
			re.Global = 1
		next
		ParseTemplate = text
	end function
	
	private function AddTemplateName(name)
		dim decorated_name : decorated_name = name & "_" & CStr(template_count)
		if block_names.Exists(name) then
			block_names.Item(name) = block_names.Item(name) & " " & decorated_name
		else
			block_names.Add name, decorated_name
		end if
		template_count = template_count+1
		AddTemplateName = decorated_name
	end function
	
	private function UpdateSlots(text)
		dim slot
		for each slot in token_list
			re.Pattern = "{{" & slot & "}}"
			text = re.Replace(text, token_list.Item(slot))
		next
		UpdateSlots = text
	end function
	
	private sub ErrMsg(err)
		Response.Write "<html><head><title>ASPTemplate Error</title></head>" & _
			"<body><h3><font color='red'>ASPTemplate Error:</font></h3>" & _
			"<p>" & err &"</p></body></html>"
	end sub
	
end class
</script>