<%@ Language=VBScript %>
<% option explicit %>

<%            
    Response.Cookies("user")("firstname") = "Hari"
    Response.Cookies("user")("lastname") = "Anton"
    Response.Cookies("user")("age") = 23
    Response.Cookies("user")("isFemale") = false
%>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Request Object</h2>
        <%
            dim cookie
            if Request.Cookies("user").HasKeys then
                for each cookie in Request.Cookies("user")
                    response.write Request.Cookies("user")(cookie)
                    response.write "<br/>"
                next
            end if 
        %>
        <hr/>
        <%
            dim operand1, operand2, sum
            operand1 = Request.QueryString("op1")
            operand2 = Request.QueryString("op2")
            sum = cdbl(operand1) + cdbl(operand2)
        %>
        <%
            dim variable
            for each variable in Request.ServerVariables
                response.write variable & " = " _ 
                & Request.ServerVariables(variable) & "<br/>"
            next
        %>

    </body>
</html>
