<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Loops</h2>
        <%   
            'For...Next - Uses a counter to run 
            'statements a specified number of times.
            dim  i
            for i = 0 to 10 step 2
                response.write "Value of i : " & i
                response.write "<br/>"
            next
            
            'For Each...Next - Repeats a group of 
            'statements for each item in a 
            'collection or each element of an array. 
            dim fruits, item
            fruits = Array("Apple", "Orange", "Cherries")
            for each item in fruits
                response.write item & "<br/>"
            next

            'do...Loop - If you don't know how many repetitions you want.
            
            'Repeat code while a condition is true
            dim  ii
            do while ii < 10
                ii = ii + 2
                response.write "Value of ii : " & ii
                response.write "<br/>"
            loop
            dim  iii
            do 
                iii = iii + 2
                response.write "Value of iii : " & iii
                response.write "<br/>"
            loop while iii < 10

            'Repeat code until a condition becomes true
            dim j
            do until j = 50
                j = j + 1
                if j = 30 then exit do
                response.write "Value of j : " & j
                response.write "<br/>"
            loop
        %>
    </body>
</html>
