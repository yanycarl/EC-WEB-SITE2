<center>
<br><a href="default.html">返回首页</a></br>
<br>
</br>
订单查询结果：
<br>
</br>
<br>
</br>
<br>
</br>



<% 


if request.cookies("guango").haskeys then

Dim connString
Dim rs,rs2

set connString =Server.CreateObject("ADODB.Connection")
connString.open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=D:\webroot\database.xls;Persist Security Info=False"
Set rs=Server.CreateObject("ADODB.Recordset")
Set rs2=Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = connString
rs2.ActiveConnection = connString

rs2.Source ="select * from [consumer$] where uid='"&request.cookies("guango")("id")&"'"
rs2.Open()

response.write("用户ID："+rs2("uid"))
response.write("  用户昵称："+rs2("name"))
response.write("  余额："+rs2("balance"))
response.write("<br>")
response.write("<br>")

rs.Source ="select * from [order$] where uid='"&request.cookies("guango")("id")&"'"
rs.Open()

response.write("<form method="&"post"&" action="&"change.asp"&">")
    response.write("<br><input name="&"oid"&" type="&"username >")
    response.write("<input name="&"ok"&" type="&"submit"&" value="&"立即支付"&">")
    response.write("</form>")

do while not rs.bof or not rs.eof

    if rs.eof or rs.bof then 
    response.write("<br>")
    response.write("没有其他订单")
    else 
    response.write("    订单号：")
    response.write(rs("oid"))
    response.write("   商品名：")
    response.write(rs("name"))
    response.write("   单价：")
    response.write(rs("price"))
    response.write("   数量：")
    response.write(rs("amount"))
    response.write("   总价格：")
    response.write(rs("prices"))
    response.write("   订单时间：")
    response.write(rs("time"))
    response.write("   送货地址：")
    response.write(rs("addr"))
    response.write("   状态：")
    response.write(rs("situaction"))
    if not rs("situaction")="订单支付成功" then
    response.cookies("guango")("oid")=rs("oid")
    end if

    
end if
    response.write("<br>")
    response.write("<br>")
    rs.movenext
loop
    
else 
    response.redirect(login.html)
end if
%>

</center>