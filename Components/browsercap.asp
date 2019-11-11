<%@ Language=VBScript %>
<% option explicit %>
<!DOCTYPE html>
<html>
    <body>
    <%
        dim browser
        Set browser = Server.CreateObject("MSWC.BrowserType")
    %>

    <table border="0" width="100%">
        <tr>
        <th>Client OS</th><th><%=browser.platform%></th>
        </tr><tr>
        <td >Web Browser</td><td ><%=browser.browser%></td>
        </tr><tr>
        <td>Browser version</td><td><%=browser.version%></td>
        </tr><tr>
        <td>Frame support?</td><td><%=browser.frames%></td>
        </tr><tr>
        <td>Table support?</td><td><%=browser.tables%></td>
        </tr><tr>
        <td>Sound support?</td><td><%=browser.backgroundsounds%></td>
        </tr><tr>
        <td>Cookies support?</td><td><%=browser.cookies%></td>
        </tr><tr>
        <td>VBScript support?</td><td><%=browser.vbscript%></td>
        </tr><tr>
        <td>JavaScript support?</td><td><%=browser.javascript%></td>
        </tr>
    </table>

</body>
</html>