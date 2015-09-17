<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>ChooseLoginID</title>
</head>

<body>
First, choose the supervisor's login ID for the Location <BR>
<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%
Dim rsAccess
Dim strSQL
Dim conn 
Dim strLoginID
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

Dim strAction 
strAction = left(lcase(Request.Form("ACTION")),3)

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsAccess = Server.CreateObject("ADODB.Recordset")
strSQL= "SELECT * FROM tblAccess ORDER BY strLoginID"
rsAccess.Open strSQL, conn, 3, 3

%>

<div align="center"><FORM action="AddLocation.asp" method=POST name=frmChooseLoginID>
	<select NAME=cboLoginID>
	<option value="NewLoginID">New Login ID</option>
<%do while not rsAccess.EOF
if ( rsAccess("strLoginID") <> "admin" ) AND ( rsAccess("strLoginID") <> "security" ) AND ( rsAccess("strLoginID") <> "science" ) then
strLoginID = rsAccess("strLoginID")
%>
	<option value="<%= strLoginID %>"><%= strLoginID %></option>
<%	end if
	rsAccess.MoveNext
	loop 
	rsAccess.Close
	set rsAccess = nothing
	conn.close
	set conn = nothing
%>
    </select></td>
	
	&nbsp;&nbsp;<input type="submit" name="btnSubmit" value="Next">
</FORM></div>


</body>
</html>
