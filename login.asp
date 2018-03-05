<html>
<head>
<title>ÕıÔÚµÇÂ¼</title>
</head>
<% 
Dim connString
Dim rs
string username,password
username=cstr(trim(Request("username")))
password=cstr(trim(Request("password")))
set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
if username <> "" or password <> "" then
rs.Source ="select* from [consumer$] where uid='"&username&"'and pwd='"&password&"'"
rs.Open()
if rs.eof or rs.bof then 
   response.redirect("login-no.asp")
   else 
   Response.Cookies("guango")("id")= username
   response.redirect("login-yes.asp")
   rs.close
   end if
else
if request.cookies("guango").haskeys then
response.redirect("personal.html")
else 
end if
response.redirect("login-no.asp")
end if
%>

</html>