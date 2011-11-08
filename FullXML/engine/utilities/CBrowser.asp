<%
	Class CBrowser
		Public Browser
		Public Version
		Private UserAgent
	
		Private sub Class_Initialize
			Call Sniff
		End Sub
		
		Private sub Class_Terminate
			
		End Sub
		
		
		'''---'''
		Private Sub Sniff			
			Dim str
			
			UserAgent = UCASE(Request.servervariables("HTTP_USER_AGENT"))
			
			if instr(UserAgent, "MSIE")>0 then
				' Opera case
				if instr(UserAgent, "OPERA")>0 then
					Browser = "OP"
					str = Mid(UserAgent, instr(UserAgent, "OPERA")+6)
					Version = Mid(str, 1, instr(str, " ")-1)
				' normal Ie
				else
					Browser = "IE"
					str = Mid(UserAgent, instr(UserAgent, "MSIE")+4)
					Version = trim(Mid(str, 1, instr(str, ";")-1))
				end if
			else
				Version = CInt(Mid(UserAgent, instr(UserAgent, "/") + 1, 1))
				if Version >= 6 then
					Browser = "NS7"
				elseif Version >= 5 then
					Browser = "NS6"
				elseif Version >= 4 then
					if instr(UserAgent, "PPC") then
						Browser = "NSMAC" 
					else
						Browser = "NSPC"
					end if
				else
					Browser = "NS3"
				end if
			end if
		End Sub
		
	End Class	
%>