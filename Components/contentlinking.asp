<%@ Language=VBScript %>
<% option explicit %>

<!DOCTYPE html>
<html>
    <body>
        <p>The example below builds a set of links</p>
        <%
        dim c
        dim i
        set nl = server.createobject("MSWC.Nextlink")
        c = nl.GetListCount("links.txt")
        i = 1
        %>
        <ul>
        <%= do while (i <= c) %>
        <li><a href="<%=nl.GetNthURL("links.txt", i)%>">
        <%=nl.GetNthDescription("links.txt", i)%></a></li>
        <%
        i = (i + 1)
        loop
        %>
        </ul>
    </body>
</html>