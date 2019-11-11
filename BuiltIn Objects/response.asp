<%@ Language=VBScript %>
<% option explicit %>
<%
    dim numberOfVisits
    Response.Cookies("NumberOfVisits").Expires = date() + 20
    numberOfVisits = request.cookies("NumberOfVisits")

    if numberOfVisits = "" then
        Response.Cookies("NumberOfVisits") = 1
        response.write "Welcome! this is the first time you visited us"
    else
        Response.Cookies("NumberOfVisits") = numberOfVisits + 1
        response.write "You have visited this web page "_
        & numberOfVisits & " times before."
    end if
%>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Response Object</h2>
        <%   
            'Embedded quotes         
            Response.Write "Hello ""World!"""
            Response.Write "<br/>"
            Response.Write "Hello " & chr(34) _
            & "World!" & chr(34)
            Response.Write "<br/>"
            Response.Write "Hello &quot;World!&quot;"
        %>
        <br/>
        <%            
            'Response.Redirect "https://www.google.com"
        %>
    </body>
</html>
