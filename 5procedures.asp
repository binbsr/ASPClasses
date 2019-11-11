<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>VBScript Procedures</h2>
        <%   
            'Sub procedure:     
            'series of statements, enclosed by the Sub and End Sub statements
            'can perform actions, but does not return a value         
            sub WriteHello(helloString)
                Response.Write helloString
            end sub
            
            WriteHello "Hello World!"

            'function :
            'is a series of statements, enclosed by the 
            'Function and End Function statements
            'can perform actions and can return a value     
            'returns a value by assigning a value to 
            'its name
            function add(a, b)
                add = a + b
            end function
            dim sum : sum = add(213, 343)
            Response.Write sum

        %>
    </body>
</html>
