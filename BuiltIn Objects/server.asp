<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP Server Object</h2>
        <%       
            Server.ScriptTimeout = 10 'set
            Response.Write Server.ScriptTimeout 'get
        %>
        <hr>
        <%
            Response.Write "I am on server.asp"
            Server.Execute "requestobj-form.asp"
            Response.Write "I am back on server.asp"
        %>
        <hr/>
        <%
            ' Response.Write "I am on server.asp"
            ' Server.Transfer "requestobj-form.asp"
            ' Response.Write "I am back on server.asp"
        %>
        <hr/>
        <%
            response.write Server.MapPath("server.asp") & "<br/>"
            response.write(Server.MapPath("respo.asp") & "<br>")
            response.write(Server.MapPath("\respo.asp") & "<br>")
            response.write(Server.MapPath("/") & "<br>")
            response.write(Server.MapPath("./") & "<br>")
        %>
        <hr/>
        <%
            dim content
            content = "The image tag : <img>. You & me, ""OK""?"
            response.write Server.HTMLEncode(content)
        %>
    </body>
</html>
