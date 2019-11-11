<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>Error Handling</h2>
        <%            
            dim fso, textstream, allcontent, filePath
            filePath = "D:\sdfsf.txt"
            on error resume next
                set fso = server.createobject("scripting.filesystemobject")
                set textstream = fso.opentextfile(filePath, 1)
                allcontent = textstream.readall()
                textstream.close()
            if Err.Number <> 0 then
                response.write "Error : " & Err.Number & "<br>"
                response.write "Source : " & Err.Source   & "<br>"
                response.write "Description : " & Err.Description  & "<br>"
                response.write filePath & " does not exist." & "<br>"
            end if
        %>
    </body>
</html>
