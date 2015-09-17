<%
Response.expires=0
if  NOT(session("LoggedIn")) then 
	dim url
	url = "login.asp?msg=noaccess"
	response.redirect (url)
end if 
%>