<html>
<head>
<title>е§дкзЂВс</title>
</head>
<% 
Dim connString
Dim rs
dim username,password
username=cstr(trim(Request("username")))
password=cstr(trim(Request("password")))
set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
if username <> "" or password <> "" then
rs.Source ="insert into [consumer$](uid,pwd) values('"&username&"','"&password&"')"
rs.Open()
response.redirect("login.html")
else
response.write("зЂВсЪЇАм")
response.redirect("register.html")
end if
%>
</html>