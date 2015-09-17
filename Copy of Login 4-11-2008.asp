<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%
Response.Expires=0
Response.Buffer = True
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

If Request.Form ("btnLogin") = "Login" then

Dim strLoginID, strPassword, msg
Dim strSQL
Dim rsAccess
Dim conn 

strLoginID=request.form("txtLoginID")
strPassword=request.form("txtPassword")
strSQL="select * from tblAccess where strLoginID='"
strSQL = strSQL & InjectionEncode(strLoginID) & "'"

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsAccess = Server.CreateObject("ADODB.Recordset")
rsAccess.Open strSQL, conn, 3, 3

If  rsAccess.eof then
   msg = "The login ID: " + strLoginID + " does not exist, try a different Login ID. Please contact administrator if you need a login ID and Password"

else If  rsAccess("strPassword")= strPassword then
		session("LoggedIn")= true
		session("strLoginID")= strLoginID
		Response.Redirect "ChemicalInventory.asp"
	else
		msg = "The password for " + strLoginID + " was not correct, please try again. Please contact administrator if you need a new Password"
	end if
end if

else
	msg = Request.QueryString("msg")
	if msg = "noaccess" then
		msg = "Please login, your user session has timed out (or you have not logged in yet)."
	end if
End if

%>
<HTML>
<HEAD>
	<TITLE>OCID Science - Login</TITLE>
</HEAD>

<BODY>
<FONT FACE="Arial" COLOR="black">
<FORM action="login.asp" name=frmlogin method=POST><TABLE CELLSPACING=10 BORDER=0 ALIGN="center" VALIGN = "MIDDLE">
<TR>
	<TD COLSPAN=2><font size="+1">Welcome to the</font> <font size="+2">UTS, Faculty of Science</font>
	<H2>Online Chemical Inventory Database</H2></TD>
</TR>
<TR>
<TD COLSPAN=2></TD>
</TR>
<TR><TD COLSPAN=2><FONT FACE="Arial" COLOR="red"><%=msg%></font>
<TR>
	<TD colspan=2><STRONG>Please Login</STRONG></TD>
</TR>
<TR>
	<TD>Login ID:</TD>
	<TD><INPUT NAME="txtLoginID" MAXLENGTH=50 ></TD>
</TR>
<TR>
	<TD>Password:</TD>
	<TD>
        <INPUT NAME="txtPassword" MAXLENGTH=50 type="password" >
      </TD>
</TR>
<TR>
	<TD>&nbsp;</TD>
	<TD><INPUT type="reset" value="Clear" name=btnReset>&nbsp;&nbsp;
	<INPUT type="submit" value="Login"  name=btnLogin></TD>
</TR>
</TABLE>
</FORM>

</FONT>

</BODY>
</HTML>
