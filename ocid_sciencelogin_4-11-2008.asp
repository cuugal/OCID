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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
        "http://www.w3.org/TR/html4/loose.dtd">
<html lang="en">
<head>
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
	<title>Online Chemical Inventory Database (OCID) - Science login</title>
	<link rel="stylesheet" type="text/css" href="http://www.ocid.uts.edu.au/ocid.css" />
</style>
</head>
<body>
<div id="header">
<img border="0" src="http://www.ocid.uts.edu.au/OCIDScience/uts-logo.gif" width="130" height="29" align="left">
	<div style="float: right;">
		<img border="0" src="http://www.ocid.uts.edu.au/images/ocid-logo2.gif" width="96" height="29"><br />
	</div>
	<div style="clear: both; float: right;"><sup style="font-size: 0.9em;">v3.1</sup></div>
</div>
<div style="background: #c7e3f9;">&nbsp;</div>
<div style="background: #fff;">&nbsp;</div>
<div id="gutter"></div>


<div id="col2">
<h2>Log in to OCID</h2>
	<ul id="navcontainer">
		<li><a href="http://www.ocid.uts.edu.au/OCIDScience/ocid_sciencelogin_4-11-2008.asp">Faculty of Science</a></li>
		<li><a href="http://www.ocid.uts.edu.au/OCIDDAB/ocid_dablogin_4-11-2008.asp">Faculty of Design, Architecture and Building</a></li>
	</ul></div>
<div id="col1"><h2>Online Chemical Inventory Database (OCID)</h2>

<hr />

<h3>Faculty of Science login</h3>

<form action="login.asp" name="frmlogin" method="post">
	<fieldset>
	<div class="required">
		<label for="first_name">Login ID:</label>
		<input type="text" name="txtLoginID" size="20" class="inputText" size="10" maxlength="100" value="" />
	</div>

	<div class="required">
		<label for="last_name">Password:</label>
		<input class="inputText" size="10" maxlength="100" value="" type="password" name="txtPassword" />
		<input type="submit" value="Login" name="btnLogin" />&nbsp;&nbsp;&nbsp;<input type="reset" value="Clear" name="btnReset" />
	</div>
	</fieldset>
</form>
</div>

<div id="col3">
<h2>Useful links</h2>
		<ul style="list-style-type: square;">
			<li><a href="http://www.ocid.uts.edu.au/ocidportal.htm">Online OCID feature tour</a></a></li>
			<li><a href="chemra.htm">OCID Risk Assessment form</a> (This form is an example only. It does not store any data.)</li>
			<li>Search Material Safety Data Sheets at <a href="http://www.chemwatch.uts.edu.au/">Chemwatch</a></li>
			<li>View : <a href="http://www.nicnas.gov.au/publications/CAR/">NICNAS Public Chemical Assessment Reports</a></li>
			<li>Search for a CAS number at <a href="http://www.chemfinder.com/">www.chemfinder.com/</a></li>
			<li>Search the <a href="http://hsis.ascc.gov.au/SearchHS.aspx">Australian Safety and Compensation Council (ASCC) Hazardous Substances Information System</a></li>
			<li><a href="http://www.ehs.uts.edu.au">UTS Environment, Health and Safety website</a></li>
		</ul>
</div>

<div id="footer">
&copy; Copyright UTS - CRICOS Provider No: 00099F&nbsp;&nbsp;&nbsp;&nbsp;E-mail page comments to the <a href="mailto:ehs.branch@uts.edu.au">EHS Branch</a><br />
<a href="http://www.uts.edu.au/disclaimer.html">Disclaimer</a> |
<a href="http://www.uts.edu.au/privacy.html">Privacy</a> |
<a href="http://www.uts.edu.au/copyright.html">Copyright</a> |
<a href="http://www.uts.edu.au/accessibility.html">Accessibility</a> |
<a href="http://www.gsu.uts.edu.au/policies/webpolicy.html">Web policy</a> |
<a href="http://www.uts.edu.au/">UTS homepage</a>
</div>
</body>
</html>