<center>
<br><a href="default.html">������ҳ</a></br>
<br>
</br>
������ѯ�����
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

response.write("�û�ID��"+rs2("uid"))
response.write("  �û��ǳƣ�"+rs2("name"))
response.write("  ��"+rs2("balance"))
response.write("<br>")
response.write("<br>")

rs.Source ="select * from [order$] where uid='"&request.cookies("guango")("id")&"'"
rs.Open()

response.write("<form method="&"post"&" action="&"change.asp"&">")
    response.write("<br><input name="&"oid"&" type="&"username >")
    response.write("<input name="&"ok"&" type="&"submit"&" value="&"����֧��"&">")
    response.write("</form>")

do while not rs.bof or not rs.eof

    if rs.eof or rs.bof then 
    response.write("<br>")
    response.write("û����������")
    else 
    response.write("    �����ţ�")
    response.write(rs("oid"))
    response.write("   ��Ʒ����")
    response.write(rs("name"))
    response.write("   ���ۣ�")
    response.write(rs("price"))
    response.write("   ������")
    response.write(rs("amount"))
    response.write("   �ܼ۸�")
    response.write(rs("prices"))
    response.write("   ����ʱ�䣺")
    response.write(rs("time"))
    response.write("   �ͻ���ַ��")
    response.write(rs("addr"))
    response.write("   ״̬��")
    response.write(rs("situaction"))
    if not rs("situaction")="����֧���ɹ�" then
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