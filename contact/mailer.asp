<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'**************************************************************************************
'**  Copyright Notice
'**  Websunami.Com - FORM TO MAIL SCRIPT - (f2m.asp)
'**  Copyright 2003-2004 Volkan Ozer All Rights Reserved.
'**
'**  This program is free software; you can modify (at your own risk) any part of it
'**  under the terms of the License that accompanies this software and use it both
'**  privately and commercially.
'**  This program is Puplished as is WITHOUT ANY WARRANTY !!!!! 
'**  All copyright notices must remain in tacked in the script 
'**  No official support is available for this program
'**  Support questions are NOT answered by e-mail
'**  For non-support related questions you can contact us via info@websunami.com

'**  http://www.websunami.com
'*************************************************************************************
Dim MailTo
Dim mailcomp
Dim smtp
Dim From
Dim Comments

' variables statrs HERE***********************************
If request ("email") = "" Then 'where email is the name of the filed in the form
From = "contact@coastalwebtech.com" 'type a default e-mail here. even a fake one will be ok
Else
From =  Request.Form ("Email") 'where email is the name of the field in the form
End If
' this will check the Email from the form submitted and assign the From address. 
' If Email field is left blank then it assign the fake address. YOu Must Enter an eMail address Above

MailTo = "sales@coastalwebtech.com"
' this is where you want e-mails to go to ***

mailcomp = "CDOSYS"
' You must specify a component to use
' CDONTS, CDOSYS, ASPEmail, JMail ****

smtp = "mail.coastalwebtech.com"
' If using CDONTS, this is not necessary (supported on Win2K, NT4 but not on WinXP Pro)
' If you are having problems with CDOSYS on a XP Pro, Search and Download CDONTS from the net
' and register the dll using regsvr ****
' for all others you must specify an SMTP server
%>
<%
	Comments = request.form("Comments")
	Comments = Replace(Comments ,vbcrlf, "<BR>" & vbcrlf)

'Generates the HTML formatted e-mail
    Dim html
    html = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
    html = html & "<html>"
    html = html & "<head>"
    html = html & "<meta http-equiv=""Content-Type"""
    html = html & "content=""text/html; charset=iso-8859-1"">"
    html = html & "<title>Surf Station Website Form Submit</title>"
    html = html & "</head>"
    html = html & "<body>" 'email bg color
	html = html & "<center><b>*** Coastal Web Tech Website Form Submit ***</b></center>"
	html = html & "<b>Name:</b> " & Request.Form("Name")
	html = html & "<br>"
	html = html & "<b>Organization:</b> " & Request.Form("Organization")
	html = html & "<br>"
	html = html & "<b>Phone Number:</b> " & Request.Form("Phone")
	html = html & "<br>"
	html = html & "<b>Website:</b> " & Request.Form("Website")
	html = html & "<br>"
	html = html & "<b>Email Address:</b> " & Request.Form("email")
	html = html & "<br>"
	html = html & "<b>Subject:</b> " & Request.Form("Subject")
	html = html & "<br>"
	html = html & "<b>Comments:</b> " & Comments
	html = html & "<br>"
	html = html & "<p></p>"
	html = html & "<b>Users IP:</b> " & Request.ServerVariables("REMOTE_ADDR") 
	html = html & "<p></p>"
    html = html & "</body>"
    html = html & "</html>"
'  ************************ $$$$ variables UP TO here $$$**************************
%>
<%
'  ************************   DO NOT MAKE ANY CHANGES BELOW    *******************
' Which mail component to use?
if mailcomp = "CDOSYS" then	
	set imsg = createobject("cdo.message")
    set iconf = createobject("cdo.configuration")
    Set Flds = iConf.Fields
    With Flds
   .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtp
   .Update
   End With
   With iMsg
   Set .Configuration = iConf
        .To = MailTo
        .From = From
        .Subject = "*** CWT Form Submit - " & Request.Form("Subject")
        .HTMLBody = HTML
        .fields.update
        .Send
    End With
    set imsg = nothing
    set iconf = nothing
    set HTML = nothing	
else

if mailcomp = "CDONTS" then
   Dim Avanos
   Set Avanos = CreateObject("CDONTS.NewMail")
    Avanos.From= From
    Avanos.To= MailTo
    Avanos.Subject= Request.Form("Subject")
    Avanos.BodyFormat=0
    Avanos.MailFormat=0
    Avanos.Body=HTML
    Avanos.Send
	set HTML = nothing
    set Avanos=nothing
	else
	
if mailcomp = "ASPEmail" then
    Set WSweb = Server.CreateObject("Persits.MailSender")
    WSweb.Host = smtp
    WSweb.From = From
    WSweb.AddAddress MailTo
    WSweb.Subject = Request.Form("Subject")
    WSweb.Body = HTML
    WSweb.IsHTML = True			
    WSweb.Send
    set WSweb = Nothing
    set HTML = Nothing
	else
	
if mailcomp = "JMail" then
    set msg = Server.CreateObject( "JMail.Message" )
    msg.Logging = true
    msg.ContentType = "text/html"
    msg.From = From
    msg.AddRecipient = MailTo
    msg.Subject = Request.Form("Subject")
    msg.Body = HTML
    msg.Send( smtp )
    Set msg = Nothing
    set HTML = Nothing

end if
end if
end if
end if

response.Redirect("thanks.asp")
%>
