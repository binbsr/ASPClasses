<%@ Language=VBScript %>
<% option explicit %>

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP - Read text file</h2>
        <%      
            dim fileSystem, file 
            set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
            '.OpenTextFile(fname,mode,create,format)
            'mode
            '1=ForReading - Open a file for reading. You cannot write to this file.
            '2=ForWriting - Open a file for writing.
            '8=ForAppending - Open a file and write to the end of the file.
            set file = fileSystem.OpenTextFile("D:\Papito jokes.txt", 1)
            response.write file.ReadAll()
            
            set file = nothing
        %>
    </body>
</html>
