<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Control Structures</h2>
        <%            
            dim i : i = hour(time)
            if i = 10 then
                Response.Write "Class Started..."
            elseIf i = 12 then
                Response.Write "Afternoon, lunch time..."
            elseIf i = 3 then
                Response.Write "Time to go home..."
            else
                Response.Write "Whatever..."
            end if

            'select case
            Dim hairColor
            hairColor = "Yellow"
            Select Case hairColor
            Case "Black"
                Response.Write("Same as the cat")
            Case "Blue"
                Response.Write("Same as the parrot")
            Case "Yellow"
                Response.Write("Same as the taxi")
            Case Else
                Response.Write("Ah well, whatever...")
            End Select
        %>
    </body>
</html>
