<%@ Language=VBScript %>
<% option explicit %>

<html>  
    <head>
        <title>ASP Introduction</title>        
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ASP - Drive meata information</h2>
        <%      
            dim fileSystem, drive, driveletter
            set fileSystem = Server.CreateObject("Scripting.FileSystemObject")
            
            ' for each driveletter in fileSystem.drives
            '     set drive = fileSystem.GetDrive(driveletter)
            '     if drive.IsReady then
            '         response.write drive.DriveLetter & "<br>"
            '         response.write drive.DriveType & "<br>"
            '         response.write drive.FileSystem & "<br>"
            '         response.write drive.TotalSize & "<br>"
            '         response.write drive.AvailableSpace & "<br><hr>"   
            '     end if             
            ' next
            '0 = unknown
            '1 = removable
            '2 = fixed
            '3 = network
            '4 = CD-ROM
            '5 = RAM disk
        %>    
        <table>
            <tr>
                <th>Drive</th>
                <th>Type Type</th>
                <th>FileSystem</th>
                <th>Total Space</th>
                <th>Available Space</th>
            </tr>

            <%
            for each driveletter in fileSystem.Drives
                set drive = fileSystem.GetDrive(driveletter)
                if drive.IsReady then
            %>
                <tr>
                    <td><%=drive.DriveLetter%></td>
                    <td><%=GetDriveType(drive.DriveType)%></td>
                    <td><%=drive.FileSystem%></td>
                    <td><%=FormatNumber(drive.TotalSize / (1024 * 1024 * 1024)) & "GB"%> </td>
                    <td><%=FormatNumber(drive.AvailableSpace / (1024 * 1024 * 1024)) & "GB"%></td>
                </tr>                
            <%
                end if
            next
            function GetDriveType(driveletter)
                select case driveletter
                    case 0: GetDriveType = "Unknown"
                    case 1: GetDriveType = "Removable"
                    case 2: GetDriveType = "Fixed"
                    case 3: GetDriveType = "Network"
                    case 4: GetDriveType = "CD-ROM"
                    case 5: GetDriveType = "RAM"
                end select
            end function
            %>
        </table>
    </body>
</html>
