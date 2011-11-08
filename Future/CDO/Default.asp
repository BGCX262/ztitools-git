<%

Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")

objCDO.To = "someone@xyz.com (John Doe)"
objCDO.From = "me@abc.com (Jane Doe)"
objCDO.bcc = "janedoe@aol.com" 'Blind cc
objCDO.Subject = "My Resume, per Request"
objCDO.Body = "Hello John. Here is a copy of my resume"
objCDO.Importance = 2 'High importance! 0 - Low 1 - Normal 2 - High
objCDO.AttachFile("\\server\jane\resume.doc","Resume.doc")
objCDO.Send 'Send off the email!

'Cleanup
Set objCDO = Nothing

%>