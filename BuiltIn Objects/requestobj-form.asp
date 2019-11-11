<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Form inputs</h2>
        <form action="requestobj-form.asp" method="POST">
            <P>First Name : <input type="text" name="firstname" value=<%=Request.Form("firstname")%>></p>
            <P>Last Name  : <input type="text" name="lastname" value=<%=Request.Form("lastname")%>></p>
            <P>Select your fav. Color : 
                <select name="favColor">
                    <option>Blue</option>
                    <option>Yellow</option>
                    <option>Green</option>
                    <option>Pink</option>
                </select>
            </p>
            <p> Select your gender:
                <input type="radio" name="gender" value="Male" <%if gender="Male" then response.write "checked" end if%>>Male
                <input type="radio" name="gender" value="FeMale">Female
                <input type="radio" name="gender" value="Others">Others
            </p>   
            <input type="submit" value="Add User" />         
        </form>
        <%
            dim firstname, lastname, color, gender
            firstname = Request.Form("firstname")
            lastname = Request.Form("lastname")
            color = Request.Form("favColor")
            gender = Request.Form("gender")

            if firstname <> "" and lastname <> "" and color <> "" then
                Response.Write "Hi, " & firstname & " " _
                & lastname & "<br>"
                Response.Write "You chose " & color & ", Great color!" & "<br>"            
            end if    
        %>
    </body>
</html>
