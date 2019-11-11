<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>VBScript random examples using built-in APIs</h2>
        <h3>Datetime</h3>
        <p>
            Today is <% Response.Write date()%>
            <br/>
            And its <% Response.Write time()%>
            <br/>
            <%
            response.write(FormatDateTime(date(),vbgeneraldate))
            response.write("<br>")
            response.write(FormatDateTime(date(),vblongdate))
            response.write("<br>")
            response.write(FormatDateTime(date(),vbshortdate))
            response.write("<br>")
            response.write(FormatDateTime(now(),vblongtime))
            response.write("<br>")
            response.write(FormatDateTime(now(),vbshorttime))
            %>
        </p>
        <p>
            Date after 45 months from today: 
            <% response.write DateAdd("m", 45, date())%>
        </p>
        <%
            dim year2500
            year2500 = cdate("1/1/2500 00:00:00")
        %>
        It is <%response.write datediff("yyyy", now(), year2500) %> years to year 2500!
        <br>
        It is <%response.write datediff("m", now(), year2500) %> months to year 2500!
        <br>
        It is <%response.write datediff("ww", now(), year2500) %> weeks to year 2500!
        
        <h3>Strings</h3>
        <% 
            dim sometext
            sometext = "Madan Bhandari Memorial College"
            sometext = trim(sometext)
            response.write strReverse(sometext) 
            response.write "<br>" 
            response.write mid(sometext, 16) 
        %>
        <br>
        <% 
            randomize()
            response.write cint(100 * rnd())
        %>
    </body>
</html>
