<center>
<br>
</br>
商品购买结果
<br>
</br>
<br>
</br>
<br>
</br>
<% 

if request.cookies("guango")("id")="0" then

    response.redirect("login.html")
else

Dim connString
Dim rs,rs2,rs3,rs4
string PRICE,name
int oid
Randomize Timer

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs3=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
Set rs4=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString
rs3.ActiveConnection = connString
rs4.ActiveConnection = connString
        
        
response.write(request.cookies("guango")("name"))

rs2.Source ="select * from [consumer$] where uid='"&request.cookies("guango")("id")&"'"
rs2.Open()

response.write("用户ID："+rs2("uid"))
response.write("<br>")
response.write("用户昵称："+rs2("name"))
response.write("<br>")
response.write("余额："+rs2("balance"))
response.write("<br>")


rs3.Source ="select * from [products$] where name='"&request.cookies("guango")("name")&"'"
rs3.open()
if rs3.eof or rs3.bof then
response.write("查无此商品")
else
response.write("商品名："+rs3("name"))
response.write("<br>")
response.write("单价："+rs3("price"))
response.write("<br>")
response.write("产地："+rs3("country"))
response.write("<br>")


price=rs3("price")
name=rs3("name")
oid=Int(899999 * Rnd + 100000)

rs.Source ="insert into [order$] values('"&rs2("uid")&"','"&oid&"','0','"&name&"','"&price&"','1','地址未填写','"&price&"','订单成功，未支付');"
rs.Open()


rs4.Source ="select * from [order$] where oid='"&oid&"'"
rs4.open()


response.write("订单号："+rs4("oid"))
response.write("<br>")
response.write("状态："+rs4("situaction"))
response.write("<br>")
response.cookies("guango")("oid")=rs4("oid")
response.write("<br><a href="&"pay.asp"&">立即支付</a></br>")

end if
end if

%>

<br>
<a href="default.html">返回首页</a>
</br>
</center>