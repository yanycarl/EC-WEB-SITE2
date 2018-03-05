<center>
<br>
</br>
µИґэЦ§ё¶ЦёБоЎЈЎЈЎЈ
<br>
</br>
<br>
</br>
<br>
</br>
<% 
string username,password
username=cstr(trim(Request("oid")))
password=cstr(trim(Request("oid")))

Dim connString
Dim rs

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
if username <> "" then
	rs.Source ="select* from [order$] where oid='"&username&"'"
	rs.Open()
	if rs.eof or rs.bof then 
   		response.redirect("orders.asp")
   	else 
   		Response.Cookies("guango")("oid")=username
   		response.redirect("pay.asp")
   		rs.close
   	end if
else
response.redirect("orders.asp")
end if
%>
</center>