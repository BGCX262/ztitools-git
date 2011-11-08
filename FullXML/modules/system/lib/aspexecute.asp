<% OPTION EXPLICIT %><!-- #include file="../../Library.asp" --><%
	Dim l_oRst, asp
	Set l_oRst = ExecuteQuery("adm_extmodule_TAspsByModuleIDSelect '" & request("moduleID")  & "'")
	asp = l_oRst("asp")
	l_oRst.close
	Set l_oRst = Nothing	
	Execute (replace (asp, vbcrlf, ":"))		
%>