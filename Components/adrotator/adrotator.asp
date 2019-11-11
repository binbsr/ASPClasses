<%@ Language=VBScript %>
<% option explicit %>
<%
    dim url
    url = Request.QueryString("url")
    If url <> "" then 
        Response.Redirect(url)
    end if
%>
<!DOCTYPE html>
<html>
    <body>
    <%
        dim adrotator
        set adrotator = server.createobject("MSWC.AdRotator")
        adrotator.TargetFrame="target='_blank'"
        response.write(adrotator.GetAdvertisement("ads.txt"))
    %>
    </body>
</html>