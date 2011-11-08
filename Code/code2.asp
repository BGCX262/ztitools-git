<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
<% 
' This program parses either an ASP/HTML file or a VB Object Source Code        
' and then it generates an easier-to-view page accordingly        
' *************** Notice ******************************************        
' * There are two ways to handle security while using this program.        
' * The first involves setting up one directory        
' * tree that you can view the code for ASP files.        
' * All other directories will be unviewable        
' * The second involves setting up one directory         
' * that you can't see ASP files. This would allow        
' * users to view the ASP code in any other directory.        
' *         
' * Please look at the function fValidPath for further information        
' *****************************************************************
%>
<HTML>
<HEAD>
<TITLE>View Active Server Page Source</TITLE>
<META http-equiv="PICS-Label" content='(PICS-1.1 "http://www.rsac.org/ratingsv01.html" l gen true comment "RSACi North America Server" by "inet@microsoft.com" for "http://www.microsoft.com" on "1997.06.30T14:21-0500" r (n 0 s 0 v 0 l 0))'>
</HEAD>
<BODY BGCOLOR="#FFFFFF" LINK="#0000FF" VLINK="#0000FF">
<% Dim strVirtualPath  'The virtual path of the file to be viewed        
Dim strFilename        'The physical path name to the file after mapping its virtual path          
Dim FileObject         'FileSystemObject        
Dim oInStream          'The In Stream for reading in the source file        
Dim strOutput          'The line that's being read from the source file        
Dim fScript            'The flag that detemines If we are currently in a script block or Not        
Dim fObjectFile        'The flag that determines If we are viewing a VB object file or Not        
Dim fFileError         'The flag that tells If there's been an file error occured        

'Defining constants        
Dim chDoubleQuote : chDoubleQuote = Chr(34)        
Dim chTab         : chTab = Chr(9)        
Dim iASPOpen      : iASPOpen = 1         
Dim iASPClose     : iASPClose = 2        
Dim iScriptOpen   : iScriptOpen = 3
Dim iScriptClose  : iScriptClose = 4
Dim iInclude      : iInclude = 5        
Dim IFRAMEOpen    : IFRAMEOpen = 6        
Dim IFRAMEClose   : IFRAMEClose = 7        
Dim iComment      : iComment = 8        'Two dimentional array that stores the ASP color legends information        

Dim arrASPColorLegend(7,1)

arrASPColorLegend(0,0) = "Server Side Script"     : arrASPColorLegend(0,1) = "#FF0000"
arrASPColorLegend(1,0) = "Client Side Script"     : arrASPColorLegend(1,1) = "#9900CC"        
arrASPColorLegend(2,0) = "Hyperlink"              : arrASPColorLegend(2,1) = "#0000FF"        
arrASPColorLegend(3,0) = "Include"                : arrASPColorLegend(3,1) = "#22AA22"        
arrASPColorLegend(4,0) = "Frames"                 : arrASPColorLegend(4,1) = "#996600"        
arrASPColorLegend(5,0) = "Comments"               : arrASPColorLegend(5,1) = "#006600"        
arrASPColorLegend(6,0) = "Object Code Link"       : arrASPColorLegend(6,1) = "#009999"        
arrASPColorLegend(7,0) = "Standard HTML and Text" : arrASPColorLegend(7,1) = "#000000"        

Dim arrASPKeyWordList(8)  'array that stores all the key words excluding comments        

'initializing all the key words that we are checking        
arrASPKeyWordList(0) = "<" & "%"                    : arrASPKeyWordList(1) = "%" & ">"        
arrASPKeyWordList(2) = "<" & "SCRIPT"               : arrASPKeyWordList(3) = "<" & "/SCRIPT"        
arrASPKeyWordList(4) = "<" & "!" & "--" & "#INCLUDE"           
arrASPKeyWordList(5) = "<" & "IFRAME "              : arrASPKeyWordList(6) = "<" & "/IFRAME"        
arrASPKeyWordList(7) = "<" & "FRAME "               : arrASPKeyWordList(8) = "<" & "/FRAME"                        

'Two dimentional array that stores the VB Object color legends information        
Dim arrVBColorLegend(5,1)        
arrVBColorLegend(0,0) = "Declaration"   : arrVBColorLegend(0,1) = "#FF0000"        
arrVBColorLegend(1,0) = "Functions"     : arrVBColorLegend(1,1) = "#9900CC"        
arrVBColorLegend(2,0) = "Procedures"    : arrVBColorLegend(2,1) = "#0000FF"        
arrVBColorLegend(3,0) = "Property Let"  : arrVBColorLegend(3,1) = "#996600"        
arrVBColorLegend(4,0) = "Property Get"  : arrVBColorLegend(4,1) = "#006600"        
arrVBColorLegend(5,0) = "Comments"      : arrVBColorLegend(5,1) = "#009999"        

'initialize all the flags to be FALSE        
fFileError = FALSE        
fScript = FALSE        
fObjectFile = FALSE            

'mapping the source file and opening the file for reading        

Call InitFileReading(fScript, fObjectFile, fFileError)                  

If Not fFileError Then                
  If FValidPath(strVirtualPath) Then %>        
  <FONT <%=Application("FontMediumHeader") %>>                
  <B>Source Code for http://<%=Request.ServerVariables("SERVER_NAME") %><%=Request.QueryString("source") %>                        
  <BR>                
  </B>        
  </FONT>        
  <BR>        COLOR LEGEND        
  <BR>        
  <TABLE BORDER=1>
  <%
  Dim arrColorLegend    'temporary array variable to store the choosen color legend array        
  Dim iCounter          'counter to iterate through and display all the color labels        
  If fObjectFile Then                
    arrColorLegend = arrVBColorLegend        
    Else                
    arrColorLegend = arrASPColorLegend        
  End If        
  For iCounter = 0 to Ubound(arrColorLegend)                
  Response.Write("<TR><TD WIDTH = ""25"" BGCOLOR=""" & arrColorLegend(iCounter,1) & """>  </TD>")                
  Response.Write("<TD><FONT " & Application("FontMainBody") & ">" & arrColorLegend(iCounter,0))                
  Response.Write("</FONT></TD></TR>")        
Next        
Response.Write("</TABLE>")        
Response.Write("<BR>")        
'If (Request("HTTP_Referer")<>"") Then                
  'Response.Write("Back to <A HREF=" & Request("HTTP_Referer") & ">" & Request("HTTP_Referer") & "</A>")        
  'End If         
Response.Write("<HR>")        
'to parse and display the processed source file        

Call ProcessFile()                
Else 
%>        
<HR>                        
View Active Server Page Source-- Access Denied<BR>                        
The code you have attempted to view has been placed in a secure directory and is Not viewable.                

<%                    
End If        
End If 
%>
</FONT></BODY></HTML>
<%        
'------------------------------------------------------------------------------------        
'This function maps the source file and initialize the file reading            
Sub InitFileReading(fScript, fObjectFile, fFileError)                    
On Error Resume Next                    
strVirtualPath = Request.QueryString("source")                    
If strVirtualPath <> "" Then                            
  fScript = FALSE                            
  fObjectFile = FALSE                            
  'Mapping the virtual file path to its physical full path                            
  strFilename = Server.MapPath(strVirtualPath)                            
  If err.number <> 0 Then                                    
    Response.Write("<HR> The path to the requested file cannot be mapped!<HR>")                                    
    fFileError = TRUE                            
    Else                                     
    'Checking to whether we are handling asp file or vb object file                                    
    If LCase(Right(strFilename, 4)) = ".vbo" Then                                            
      fObjectFile = TRUE                                    
    End If                                    
    'Opening the mapped file for reading                                    
    
    Set FileObject = Server.CreateObject("Scripting.FileSystemObject")                                    
    If err.number <> 0 Then                                            
      Response.Write("<HR>The FileSystemObject cannot be created.<HR>")                                            
      fFileError = TRUE                                    
      Else                                            
      Set oInStream = FileObject.OpenTextFile(strFilename, 1, FALSE, 0)                                            
      If err.number <> 0 Then                                                    
        Response.Write("<HR>The code you have attempted to view could not be retrieved.<HR>")
        
                                                            
                                                            fFileError = TRUE                                            
                                                          End If                                    
                                                        End If                            
                                                      End If                    
                                                      Else                            
                                                      Response.Write("<HR> There is no source file passed in! <HR>")                            
                                                      fFileError = TRUE                    
                                                    End If            
                                                  End Sub        
                                                  '------------------------------------------------------------------------------------        
                                                  'This procedure parses the source file and display the processed content        
                                                  Sub ProcessFile()                
                                                  On Error Resume Next                                                
                                                  Response.Write("<PRE>")                
                                                  'the two blocks in the following If-Else statements are containing the same                
                                                  'codes except for calling either PrintObjectLine or FPrintASPLine.                
                                                  'The reason I do Not merge the two blocks into one by using a flag is because                
                                                  'it would be too expensive to determine every time which functions to call according                 
                                                  'to the flag when this needs to be execuated numerous times.                
                                                  If fObjectFile Then                        
                                                    Do Until oInStream.AtEndOfStream                                
                                                    strOutput = oInStream.ReadLine                                      
                                                    If err.number <> 0 Then                                        
                                                      Response.Write("<HR> Error Processing File! <HR>")                                        
                                                      exit Sub                                
                                                    End If                                
                                                    Call PrintObjectLine(strOutput)                                
                                                    Response.Write("<BR>")                        
                                                    Loop                
                                                    Else                        
                                                    Do Until oInStream.AtEndOfStream                                
                                                    strOutput = oInStream.ReadLine                                
                                                    If err.number <> 0 Then                                        
                                                      Response.Write("<HR> Error Processing File! <HR>")                                        
                                                  exit Sub                                
                                                  End If                                
                                                  fScript = FPrintASPLine(strOutput, fScript)                                
                                                  Response.Write("<BR>")                        
                                                  Loop                                                        
                                                  End If                
                                                  Response.Write("</PRE>")        
                                                  End Sub        
                                                  '------------------------------------------------------------------------------------        
                                                  ' The security is currently set to the second method,        
                                                  ' with the directory /Secure/ as our private directory.        
                                                  ' You can either keep this method, keep this method and         
                                                  ' change the private directory, or add an additional directory.        
                                                  '         ' You can also switch security methods by changing the         
                                                  ' "= 0" to "<> 0", and the pointing to the directory         
                                                  ' that you want to be viewable.        
                                                  Function FValidPath (ByVal strPath)                
                                                  FValidPath = (InStr(1, strPath, "/SECURE/", 1) = 0)        
                                                  End Function        
                                                  '------------------------------------------------------------------------------------        
                                                  'Returns the minimum number greater than 0        
                                                  'If both are 0, returns -1        
                                                  'this Function is used to get the precedence of keywords        
                                                  Function IMin (iNum1, iNum2)                
                                                  If iNum1 = 0 AND iNum2 = 0 Then                        
                                                    IMin = -1                
                                                  ElseIf iNum2 = 0 Then                        
                                                    IMin = iNum1                
                                                  ElseIf iNum1 = 0 Then                        
                                                    IMin = iNum2                
                                                  ElseIf iNum1 <    iNum2 Then                        
                                                    IMin = iNum1                
                                                    Else                         
                                                    IMin = iNum2                
                                                  End If        
                                                End Function        
                                                '------------------------------------------------------------------------------------        
                                                'This Function returns the number of occurrence of strToken within the string strInString          
                                                Function CTokenOccurrence(ByVal strInString, ByVal strToken)                
                                                Dim iPos  'temporary index for doing the counting                
                                                CTokenOccurrence = 0                
                                                iPos = InStr(1, strInString, strToken, 1)                
                                                If iPos <> 0 Then                        
                                                  do until iPos = 0                                
                                                  CTokenOccurrence = CTokenOccurrence + 1                                
                                                  iPos = InStr(iPos + 1, strInString, strToken, 1)                        
                                                  loop                
                                                End If        
                                              End Function                                
                                              '------------------------------------------------------------------------------------         
                                              'Finding the virtual path excluding the filename            
                                              Function StrGetFullRelativePath(ByVal strVirtualFileName)                    
                                              Dim iPos  'temporary index for doing the mapping                    
                                              iPos = InStrRev(strVirtualFileName, "/", -1)                    
                                              If iPos <> 0 Then                            
                                                StrGetFullRelativePath = Left(strVirtualFileName, iPos)                    
                                                Else                            
                                                StrGetFullRelativePath = ""                    
                                              End If            
                                            End Function                
                                            '------------------------------------------------------------------------------------        
                                            'this function retrieves the hyperlink of a string and breaks down the string        
                                            'into 3 parts: strLeftString storing the thing left to the hyperlink        '              
                                            strRightString storing the thing right to teh hyperlink        '              
                                            strHyperLink storing the HTML tagged hyperlink        
                                          Sub GetHyperLink(ByVal strFullString, strHyperLink, strLeftString, strRightString)                
                                            Dim strPageLink         
                                            'temporary string storing the hyperlink location                
                                            Dim iStartPos           
                                            'index indicating the beginning of the hyperlink                
                                            Dim iEndPos             
                                            'index indicating the Ending of the hyperlink                
                                            strHyperLink = ""                
                                            strLeftString = ""                
                                            strRightString = ""                                 
                                            iStartPos = InStr(1, strFullString, chDoubleQuote, 1)                
                                            strLeftString = Left(strFullString, iStartPos)                            
                                            strFullString = Right(strFullString, Len(strFullString) - iStartPos)                
                                            iEndPos = InStr(1, strFullString, chDoubleQuote, 1)                
                                            strPageLink = Left(strFullString, iEndPos - 1)                
                                            'check to see If we are dealing with full relative link or Not                
                                            If Left(strPageLink,1) <> "/" Then                        
                                              strHyperLink = StrGetFullRelativePath(strVirtualPath)                        
                                            If strHyperLink = "" Then                                
                                              strHyperLink = "/"                        
                                            End If                        
                                          strHyperLink = strHyperLink & strPageLink                
                                          Else                        
                                          strHyperLink = strPageLink                
                                          End If                
                                          strRightString = Right(strFullString, Len(strFullString) - iEndPos + 1)                
                                          strHyperLink = Request.ServerVariables("SCRIPT_NAME") & "?source=" & strHyperLink                
                                          strHyperLink = "<A HREF=""" & strHyperLink & """>" & strPageLink & "</A>"        
                                          End Sub        
                                          '------------------------------------------------------------------------------------        
                                          'This Function parse a line (or a Sub-line) in asp files and look for        
                                          'keywords and set the precedence for them from left to right.        
                                          'In aNother word, the leftmost keyword found will have the highest        
                                          'precedence. This Function is returning the code which determines which        'keyword that we are working on.        Function ICheckASPLineForKeyWords (ByVal strLine, ByVal fInScript, iCurrentIndex)                
                                          Dim iPos               
                                          'variable that holds the offset of keywords in the string                
                                          Dim iKeyWord           
                                          'variable that holds the index of the keyword                        
                                          ICheckASPLineForKeyWords = 0                
                                          iCurrentIndex = 0                
                                          for iKeyWord = LBound(arrASPKeyWordList) to UBound(arrASPKeyWordList)                        
                                          iPos = InStr(1, strLine, arrASPKeyWordList(iKeyWord), 1)                        
                                          If IMin(iCurrentIndex, iPos) = iPos Then                                
                                            iCurrentIndex = iPos                                
                                          'Both IFRAME and FRAME have the same index                                
                                          If ((iKeyWord + 1) = IFRAMEOpen) Or ((iKeyWord - 1) = IFRAMEOpen) Then                                        
                                            ICheckASPLineForKeyWords = IFRAMEOpen                                
                                          ElseIf ((iKeyWord + 1) = IFRAMEClose) Or ((iKeyWord - 1) = IFRAMEClose) Then
                                            ICheckASPLineForKeyWords = IFRAMEClose                                
                                            Else                                        
                                            ICheckASPLineForKeyWords = iKeyWord + 1                                
                                          End If                        
                                        End If                
                                      next                
                                      'We are treating comments in a special way.                
                                      'We will only work on them If there's no other keywords found                
                                      'in the current line or Sub-line. This gurantee that no keyword                
                                      'will be omitted If they are after the comment tag.                
                                      iPos = InStr(1, strLine, "REM", 1)  'Testing of quote checking                
                                      If (iCurrentIndex = 0 And iPos <> 0 And fInScript) Then                        
                                        iCurrentIndex = iPos                        
                                        ICheckASPLineForKeyWords = iComment                
                                      End If                
                                      iPos = InStr(strLine, "'")  
                                      REM Testing of quote checking                
                                      'fInScript tells If the current line is within a scripting tag                
                                      If (iCurrentIndex = 0 And iPos <> 0 And fInScript) Then                        
                                        iCurrentIndex = iPos                        
                                        ICheckASPLineForKeyWords = iComment                
                                      End If        
                                    End Function        
                                    '------------------------------------------------------------------------------------        
                                    'This Function encodes and print out a HTML line         
                                    Sub PrintHTML (ByVal strLine)                Dim cSpaces  'number of spaces                Dim iPos     'index for doing the encoding                Dim iLooper  'index for doing the looping                cSpaces = Len(strLine) - Len(LTrim(strLine))                iLooper = 1                'handling tabs and we make it equal to 2 spaces                While Mid(Strline, iLooper, 1) = chTab                        cSpaces = cSpaces + 2                        iLooper = iLooper + 1                WEnd                If cSpaces > 0 Then                        For iLooper = 1 to cSpaces                                Response.Write("&nbsp;")                        Next                End If                iPos = InStr(strLine, "<")                If iPos Then                        Response.Write(Left(strLine, iPos - 1))                        Response.Write("&lt;")                        strLine = Right(strLine, Len(strLine) - iPos)                        Call PrintHTML(strLine)                Else                        Response.Write(Server.HTMLEncode(strLine))                End If        End Sub        '------------------------------------------------------------------------------------        'we are dealing with the open asp tag here        Function FHandleASPOpen(ByVal strLine, ByVal iPos)                Call PrintHTML(Left(strLine, iPos - 1))                Response.Write("<FONT COLOR=#ff0000>")                Response.Write("&lt;%")                FHandleASPOpen = FPrintASPLine(Right(strLine, Len(strLine) - (iPos + 1)), TRUE)         End Function        '------------------------------------------------------------------------------------        'we are dealing with the closing asp tag here        Function FHandleASPClose(ByVal strLine, ByVal iPos)                Call FPrintASPLine(Left(strLine, iPos - 1), TRUE)                Response.Write("%&gt;")                Response.Write("</FONT>")                FHandleASPClose = FPrintASPLine(Right(strLine, Len(strLine) - (iPos + 1)), FALSE)        End Function        '------------------------------------------------------------------------------------        'we are dealing with the open script tag here        Function FHandleScriptOpen(ByVal strLine, ByVal iPos)                Dim iTempPos1       'temporary index for checking client or server script                Dim iTempPos2       'temporary index for checking client or server script                Dim strRightString  'stores the chopped part to the right of the string                Call PrintHTML(Left(strLine, iPos - 1))                strRightString = Right(strLine, Len(strLine) - (iPos + 6))                'checking to see If the SCRIPT tag is for client side                'or for server side                iTempPos1 = InStr(1, strRightString, "server", 1)                iTempPos2 = InStr(strRightString, ">")                'determining whether we are handling server side scripting or                 'client side scripting, and rEnder the corresponding color                If iTempPos1 <> 0 Then                        If iTempPos2 <> 0 Then                                If iTempPos1 <    iTempPos2 Then                                        Response.Write("<FONT COLOR=#ff0000>")                                Else                                        Response.Write("<FONT COLOR=#9900CC>")                                End If                        Else                                Response.Write("<FONT COLOR=#ff0000>")                        End If                Else                        Response.Write("<FONT COLOR=#9900CC>")                End If                                        Response.Write("&lt;SCRIPT")                FHandleScriptOpen = FPrintASPLine(strRightString, TRUE)        End Function        '------------------------------------------------------------------------------------        'we are dealing with the closing script tag here        Function FHandleScriptClose(ByVal strLine, ByVal iPos)                Call FPrintASPLine(Left(strLine, iPos - 1), TRUE)                Response.Write("&lt;/SCRIPT&gt;")                Response.Write("</FONT>")                FHandleScriptClose = FPrintASPLine(Right(strLine, Len(strLine) - (iPos + 8)), FALSE)        End Function        '------------------------------------------------------------------------------------        'we are dealing with the include tag here        Function HandleIncludeTag(ByVal strLine, ByVal iPos, ByVal fInScript)                Dim strHyperLink   'stores the HTML tagged hyperlink                
                                                            
                                                            
       Dim strLeftString  'stores the chopped part to the left of the string                Dim strRightString 'stores the chopped part to the right of the string                Call PrintHTML(Left(strLine, iPos - 1))                Response.Write("<FONT COLOR=#22AA22>")                Response.Write("&lt;!--#INCLUDE")                'finding the hyperlink of the INCLUDE by trapping for the two double quotes around it                Call GetHyperLink(Right(strLine, Len(strLine) - iPos - 11), strHyperLink, strLeftString, strRightString)                Call FPrintASPLine(strLeftString, fInScript)                Response.Write(strHyperLink)                iPos = InStr(1, strRightString, "-->", 1)                Call FPrintASPLine(Left(strRightString, iPos - 1), fInScript)                Response.Write("-->")                Response.Write("</FONT>")                If len(strRightString) - iPos > 3 Then                        HandleIncludeTag = FPrintASPLine(Right(strRightString, Len(strRightString) - iPos - 3), fInScript)                End If        End Function        '------------------------------------------------------------------------------------        'we are dealing with the FRAME and IFRAME tags here        Function HandleFrameOpen(ByVal strLine, ByVal fInScript)                Dim iPos           'temporary index for parsing the (sub) string                Dim strHyperLink   'stores the HTML tagged hyperlink                Dim strLeftString  'stores the chopped part to the left of the string                Dim strRightString 'stores the chopped part to the right of the string                                iPos = InStr(1, strLine, "<" & "IFRAME", 1)                If iPos <> 0 Then                        'we have a IFRAME tag                        Call PrintHTML(Left(strLine, iPos - 1))                        Response.Write("<FONT COLOR=#996600>")                        Response.Write("&lt;IFRAME")                          strLine = Right(strLine, Len(strLine) - iPos - 6)                Else                        'we have a FRAME tag                        iPos = InStr(1, strLine, "<" & "FRAME", 1)                        Call PrintHTML(left(strLine, iPos - 1))                        Response.Write("<FONT COLOR=#996600>")                        Response.Write("&lt;FRAME")                        strLine = Right(strLine, Len(strLine) - iPos - 5)                End If                iPos = InStr(1, strLine, "SRC", 1)                Call FPrintASPLine(Left(strLine, iPos - 1), fInScript)                                                'finding the hyperlink of the FRAME/IFRAME by trapping for the two double quotes around it                Call GetHyperLink(Right(strLine, Len(strLine) - iPos + 1), strHyperLink, strLeftString, strRightString)                Call FPrintASPLine(strLeftString, fInScript)                Response.Write(strHyperLink)                HandleFrameOpen = FPrintASPLine(strRightString, fInScript)        End Function'------------------------------------------------------------------------------------        'we are dealing with the FRAME and IFRAME tags here        Function HandleFrameClose(ByVal strLine, ByVal fInScript)                Dim iPos           'temporary index for parsing the (sub) string                Dim strHyperLink   'stores the HTML tagged hyperlink                Dim strLeftString  'stores the chopped part to the left of the string                Dim strRightString 'stores the chopped part to the right of the string                                iPos = InStr(1, strLine, "<" & "/IFRAME>", 1)                If iPos <> 0 Then                        'we have a IFRAME tag                        Call PrintHTML(Left(strLine, iPos - 1))                        Response.Write("&lt;/IFRAME&gt;")                        Response.Write("</FONT>")                        HandleFrameClose = FPrintASPLine(Right(strLine, Len(strLine) - (iPos + 8)), FALSE)                Else                        'we have a FRAME tag                        iPos = InStr(1, strLine, "<" & "/FRAME>", 1)                        Call PrintHTML(Left(strLine, iPos - 1))                        Response.Write("&lt;/FRAME&gt;")                        Response.Write("</FONT>")                        HandleFrameClose = FPrintASPLine(Right(strLine, Len(strLine) - (iPos + 7)), FALSE)                End If        End Function        '------------------------------------------------------------------------------------        'we are dealing with commenting here        Sub HandleCommentTag(ByVal strLine, ByVal iPos, ByVal fInScript)                Dim strHyperLink   'stores the HTML tagged hyperlink                Dim strLeftString  'stores the chopped part to the left of the string                Dim strRightString 'stores the chopped part to the right of the string                strLeftString = Left(strLine, iPos - 1)                Call FPrintASPLine(strLeftString, fInScript)                strRightString = Right(strLine, Len(strLine) - iPos + 1)                'checking to see If we are inside a string bounded by 2 double quotes                If ((CTokenOccurrence(strLeftString, """") mod 2) = 0) And fInScript Then                        Response.Write("<FONT COLOR=#006600>")                        iPos = InStr(1, strRightString, "OBJECTCODE", 1)                        If iPos <> 0 Then                                CAll PrintHTML(Left(strRightString, iPos - 1))                                Response.Write("<FONT COLOR=#009999>")                                                                                'finding the hyperlink of the VB object by trapping for the two double quotes                                'around it                                Call FindHyperLink(Right(strRightString, Len(strRightString) - iPos + 1), strHyperLink, strLeftString, strRightString)                                Call PrintHTML(strLeftString)                                Response.Write(strHyperLink)                                Response.Write("</FONT>")                        End If                        Call PrintHTML(strRightString)                        Response.Write("</FONT>")                Else                        'Jump over the void quote or void REM in the string and to check to                        'see If there's commenting further down the string                        iPos = InStr(strRightString, chDoubleQuote)                        If iPos = 0 Then                                Call FPrintASPLine(strRightString, fInScript)                        Else                                Call PrintHTML(Left(strRightString, iPos - 1))                                Response.Write("""")                                Call FPrintASPLine(Right(strRightString, Len(strRightString) - iPos), fInScript)                        End If                End If        End Sub        '------------------------------------------------------------------------------------        'This Function rebuilt and prints out the line according        'to the current keyword that we are working on        Function FPrintASPLine(ByVal strLine, ByVal fInScript)                Dim iKeyWordIndex   'index of the key word that we need to work on                Dim iKeyWordPos     'index of the picked key word in the string                FPrintASPLine = fInScript                iKeyWordIndex = ICheckASPLineForKeyWords(strLine, fInScript, iKeyWordPos)                Select Case iKeyWordIndex                        Case 0                                Call PrintHTML(strLine)                        Case iASPOpen                                FPrintASPLine = FHandleASPOpen(strLine, iKeyWordPos)                        Case iASPClose                                FPrintASPLine = FHandleASPClose(strLine, iKeyWordPos)                        Case iScriptOpen                                FPrintASPLine = FHandleScriptOpen(strLine, iKeyWordPos)                        Case iScriptClose                                FPrintASPLine = FHandleScriptClose(strLine, iKeyWordPos)                        Case iInclude                                FPrintASPLIne = HandleIncludeTag(strLine, iKeyWordPos, fInScript)                            Case IFRAMEOpen                                    FPrintASPLIne = HandleFrameOpen(strLine, fInScript)                        Case IFRAMEClose                                    FPrintASPLIne = HandleFrameClose(strLine, fInScript)                        Case iComment                                    Call HandleCommentTag(strLine, iKeyWordPos, fInScript)                                Case Else                                Response.Write("Function ERROR -- PLEASE CONTACT ADMIN.")                End Select        End Function        '------------------------------------------------------------------------------------        'This Function parse a line (or a Sub-line) in VB source code and look for        'keywords and set the precedence for them from left to right.        'In aNother word, the leftmost keyword found will have the highest        'precedence. This Function is returning the code which determines which        'keyword that we are working on.        Function ICheckObjectLineForKeyWords (ByVal strLine)                Dim arrASPKeyWordList(6)  'array that stores all the key words excluding comments                Dim iTemp              'variable that holds the current leftmost keyword offset                Dim iPos               'variable that holds the offset of keywords                Dim iKeyWord           'variable that holds the index of the keyword                        ICheckObjectLineForKeyWords = 0                iTemp = 0                arrASPKeyWordList(0) = "End"      : arrASPKeyWordList(1) = "Function"                arrASPKeyWordList(2) = "Sub"      : arrASPKeyWordList(3) = "property"                arrASPKeyWordList(4) = " as "     : arrASPKeyWordList(5) = "Dim "                  arrASPKeyWordList(6) = "'"                                : arrASPKeyWordList(7) = "REM"                        for iKeyWord = LBound(arrASPKeyWordList) to UBound(arrASPKeyWordList)                        iPos = InStr(1, strLine, arrASPKeyWordList(iKeyWord), 1)                        If IMin(iTemp, iPos) = iPos Then                                iTemp = iPos                                If (iKeyWord = 5) OR (iKeyWord = 8) Then                                        ICheckObjectLineForKeyWords = iKeyWord                                Else                                        ICheckObjectLineForKeyWords = iKeyWord + 1                                End If                        End If                next        End Function        '------------------------------------------------------------------------------------        'This function prints out the lines of VB objects        'Note. Please ignore this function for now        Sub PrintObjectLine (ByVal strLine)                Dim iKeyWordIndex                iKeyWordIndex = ICheckObjectLineForKeyWords(strLine)                Select Case iKeyWordIndex                            Case 0                                Call PrintHTML(strLine)                        Case 1                                Call PrintHTML(strLine)                                If instr(1, strLine, "Function", 1) <> 0 Then                                        Response.Write("</FONT>")                                        Response.Write("<HR>")                                ElseIf instr(1, strLine, "Sub", 1) <> 0 Then                                        Response.Write("</FONT>")                                        Response.Write("<HR>")                                ElseIf instr(1, strLine, "property", 1) <> 0 Then                                        Response.Write("</FONT>")                                        Response.Write("<HR>")                                End If                        Case 2                                Response.Write("<FONT COLOR= #9900CC>")                                Call PrintHTML(strLine)                        Case 3                                Response.Write("<FONT COLOR= #0000FF>")                                Call PrintHTML(strLine)                        Case 4                                If instr(1, strLine, "let", 1) <> 0 Then                                        Response.Write("<FONT COLOR= #996600>")                                ElseIf instr(1, strLine, "get", 1) <> 0 Then                                        Response.Write("<FONT COLOR= #006600>")                                End If                                Call PrintHTML(strLine)                        Case 5                                Response.Write("<FONT COLOR= #FF0000>")                                Call PrintHTML(strLine)                                Response.Write("</FONT>")                        Case 6                                iPos = InStr(strLine, "'")                                leftString = Left(strLine, iPos - 1)                                Call PrintObjectLine(leftString, ICheckObjectLineForKeyWords(leftString))                                strLine = Right(strLine, Len(strLine) - iPos + 1)                                Response.Write("<FONT COLOR= #009999>")                                Call PrintHTML(strLine)                                Response.Write("</FONT>")                        Case Else                                Response.Write("Function ERROR -- PLEASE CONTACT ADMIN.")                End Select        End Sub      %>
