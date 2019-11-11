<%@ Language=VBScript %>
<% option explicit %>
<!-- #INCLUDE FILE="adovbs.inc" -->

<html>
    <head>
        <title>ASP Introduction</title>
    </head>
    <body>
        <h1>Welcome to ASP Programming!</h1>
        <h2>ADO - Connecting to Database</h2>
        <%            
            dim Conn, sql, rowsAffected
            set conn = Server.CreateObject("ADODB.Connection")
            conn.Mode = adModeReadWrite
            conn.ConnectionString = "Provider=SQLOLEDB;Server=.\SQLEXPRESS;" &  _
            "Database=MBMDB2;User Id=bishnu;Password=123456;"
            conn.Open
            response.write "Connected Successfully."
            
            sql = "Insert into Student (FirstName, Surname, Address, Phone)" & _
            " Values ('Ramesh', 'Elton','AnamNagar, Ktm', '789809809')"
            
            conn.Execute sql, rowsAffected

            response.write rowsAffected & " row inserted."
            conn.close
        %>
    </body>
</html>
