<%
if request.cookies("guango").haskeys then
else
      response.cookies("guango")("id")="0"
end if
Response.Cookies("guango")("name")="длсмцФ╟Э╦и"
response.redirect("buy.asp")%>