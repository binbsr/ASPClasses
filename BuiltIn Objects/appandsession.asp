<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP Application and Session Objects</h2>
        <%            
            Application("CollegeName") = "MBM, Anamnagar"
            Session("info") = "Its Session object"

            response.write Application.Contents("CollegeName") & "<br/>"
            response.write Session.Contents("info")  & "<br/>"
            response.write Session.SessionID  & "<br/>"
        %>
        <hr>
        <p>Your Page Requests/Reloads:
        <%
            session("hits") = session("hits") + 1
            response.write session("hits")
        %>
        </p>
        <p>Total Page Requests/Reloads from all users:
        <%
            Application.Lock
                application("hits") = application("hits") + 1
            Application.Unlock

            response.write application("hits")
        %>
        </p>
        <hr>
        <form action="appandsession.asp" method="POST">
            <input type="submit" value="Destroy Session" name="destroy_session" />
        <form>
        <%
            if request.form("destroy_session") = "Destroy Session" then
                session.abandon
                response.write "Your session has been destroyed."
                response.end
            end if
        %>
        
        <p>Number of online users (active sessions) : 
        <%
            response.write Application("noofusersonline")
        %>
    </body>
</html>
