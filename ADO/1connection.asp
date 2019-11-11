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
            dim Conn, rs, x
            set conn = Server.CreateObject("ADODB.Connection")
            conn.Mode = adModeReadWrite
            conn.ConnectionString = "Provider=SQLOLEDB;Server=.\SQLEXPRESS;" &  _
            "Database=MBMDB2;User Id=bishnu;Password=123456;"
            conn.Open
            response.write "Connected Successfully."

            set rs = Server.CreateObject("ADODB.Recordset")
            rs.Open "Select * From Student", conn

            do until rs.EOF
                for each x in rs.Fields
                    response.write x.name & " = "
                    response.write x.value & "<br/>"
                next
                rs.MoveNext
            loop            
            rs.close
            conn.close
        %>
    </body>
</html>
