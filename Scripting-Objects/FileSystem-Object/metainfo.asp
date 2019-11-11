<%@ Language=VBScript %>
<% option explicit %>

<html>  
    <head>
        <title>ASP Introduction</title>
        <style>
            table, th, td {
                border: 2px solid blue;
            }
        </style>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP - File/Folder/Drive meata information</h2>
        <%      
            dim fileSystem, f 
            set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
            set f = fileSystem.GetFile("D:\Papito jokes.txt")
            Response.Write("File Name: " & f.ShortName & "<br>")
            Response.Write("File created: " & f.DateCreated & "<br>")
            Response.Write("File Last modified: " & f.DateLastModified & "<br>")
            Response.Write("File Type: " & f.Type & "<br>")
            Response.Write("File size (bytes): " & f.Size & "<br>")
            set f = nothing
        %>        
        <hr/>
        <%  
            dim fo, x
            set fo = fileSystem.GetFolder("D:\Submissions")
            Response.Write("Folder Name: " & fo.ShortName & "<br>")
            Response.Write("Created: " & fo.DateCreated & "<br>")
            Response.Write("Last modified: " & fo.DateLastModified & "<br>")
            Response.Write("Size (bytes): " & fo.Size & "<br><br>")
                       
            'set fo = nothing
        %>

        <h4>Files in this folder:</h4>
        <table>
            <tr>
                <th>Name</th>
                <th>File Type</th>
                <th>Size</th>
                <th>Created Date</th>
                <th>Last Modified Date</th>
            </tr>

            <%for each x in fo.Files%>
                <tr>
                    <td><%=x.Name%></td>
                    <td><%=x.Type%></td>
                    <td><%=x.Size%></td>
                    <td><%=x.DateCreated%> </td>
                    <td><%=x.DateLastModified%></td>
                </tr>
            <%next%>
        </table>
    </body>
</html>
