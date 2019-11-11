<%@ Language=VBScript %>
<% option explicit %>

<!DOCTYPE html>
<html>
    <body>
        <p>Sending and receiveing mails</p>

        <%
        'Microsoft’s Collaboration Data Objects for NT Server (CDONTS)
        Dim mail
        Set mail = Server.CreateObject("CDONTS.Newmail")
        mail.From = "me@mymail.com"
        mail.To = "you@yourmail.com"
        mail.Subject = "Mastering ASP"
        mail.Body = "This mail was sent using CDONTS"
        mail.Send
        %>
    </body>
</html>