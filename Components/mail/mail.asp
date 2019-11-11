<%@ Language=VBScript %>
<% option explicit %>

<!DOCTYPE html>
<html>
    <body>
        <p>Sending and receiveing mails</p>

        <%
            Set myMail = CreateObject("CDO.Message")
            myMail.Subject = "Sending email with CDO"
            myMail.From = "mymail@mydomain.com"
            myMail.To = "someone@somedomain.com"
            myMail.Bcc = "someoneelse@somedomain.com"
            myMail.Cc = "someoneelse2@somedomain.com"
            myMail.TextBody = "This is a message."
            myMail.AddAttachment "c:\mydocuments\test.txt"
            myMail.Send
            set myMail = nothing
        %>
    </body>
</html>
