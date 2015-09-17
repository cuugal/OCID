<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<html>
<head>
	<title>ChooseLoginID</title>
</head>

<body>
Please choose the manager's login ID for the Location <BR>

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

<div align="center"><FORM action="UpdateLocation.asp" method=POST name=frmChooseLoginID>
	<select NAME=cboLoginID>
	<option value="NewLoginID">New Login ID</option>
<%do while not rsAccess.EOF 
strLoginID = rsAccess("strLoginID")
%>
	<option value="<%= strLoginID %>"><%= strLoginID %></option>
<%	rsAccess.MoveNext
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
