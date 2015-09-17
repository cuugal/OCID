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
<script type="text/javascript">
//This focuses the user client on the login ID form, for ease of login - CL 9/5/2011
function formfocus() {
   document.getElementById('txtLoginID').focus();
   }
 window.onload = formfocus;
 </script>
</head>
<body>
<div id="header">
<img border="0" src="http://www.ocid.uts.edu.au/OCIDScience/uts-logo.gif" width="130" height="29" align="left">
	<div style="float: right;">
		<a href="http://www.ocid.uts.edu.au" title="Online Chemical Inventory Database"><img border="0" src="http://www.ocid.uts.edu.au/images/ocid-logo2.gif" width="96" height="29"></a><br />
	</div>
	<div style="clear: both; float: right;"><sup style="font-size: 0.9em;">v3.1</sup></div>
</div>
<div style="background: #c7e3f9;">&nbsp;</div>
<div style="background: #fff;">&nbsp;</div>
<div id="gutter"></div>


<div id="col2">
<h2>Log in to OCID</h2>
	<ul id="navcontainer">
		<li><a href="http://www.ocid.uts.edu.au/OCIDScience/Login.asp" title="Log in to the Faculty of Science's section of OCID">Faculty of Science</a></li>
		<li><a href="http://www.ocid.uts.edu.au/OCIDDAB/Login.asp" title="Log in to the Faculty of Design, Architecture and Building's section of OCID">Faculty of Design, Architecture and Building</a></li>
		<li><a href="http://www.ocid.uts.edu.au/OCIDFEIT/Login.asp" title="Log in to the Faculty of Engineering and Information Technology">Faculty of Engineering and Information Technology</a></li>		
		<li><a href="http://www.ocid.uts.edu.au/OCIDFASS/Login.asp" title="Log in to the Faculty of Arts and Social Sciences section of OCID">Faculty of Arts and Social Sciences</a></li>
	</ul></div>
<div id="col1"><h2>Online Chemical Inventory Database (OCID)</h2>

<hr />

<h3>Faculty of Science login</h3>

<form action="login.asp" name="frmlogin" method="post">
	<fieldset>
	<div class="required">
		<label for="first_name">Login ID:</label>
		<input type="text" name="txtLoginID" id="txtLoginID" size="20" class="inputText" size="10" maxlength="100" value="" />
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
			<!-- DLJ 5Aug2010 <li><a href="http://www.ocid.uts.edu.au/ocidportal.htm">Online OCID feature tour</a></li> -->
			<!-- CL 7/6/2010  <li><a href="chemra.htm">OCID Risk Assessment form</a> (This form is an example only. It does not store any data.)</li> --> 
			<!-- removed by CL 17/6/2013 <li>Search Material Safety Data Sheets at <a href="http://www.chemwatch.uts.edu.au/">Chemwatch</a></li>
			<li>View : <a href="http://www.nicnas.gov.au/publications/CAR/">NICNAS Public Chemical Assessment Reports</a></li>
			<li>Search for a CAS number at <a href="http://www.chemexper.com/">www.chemexper.com</a></li>
			<li>Search the <a href="http://hsis.ascc.gov.au/SearchHS.aspx">Australian Safety and Compensation Council (ASCC) Hazardous Substances Information System</a></li>  -->
			<li><a href="http://www.safetyandwellbeing.uts.edu.au">UTS Safety &amp; Wellbeing: Chemical safety</a></li>
			<li><a href="TEMPLATE_OCID3-2.xls">Excel spreadsheet of OCID fields</a> (Used for setting up a location in OCID.)</li>
		</ul>
</div>

<div id="footer">
&copy; Copyright UTS - CRICOS Provider No: 00099F&nbsp;&nbsp;&nbsp;&nbsp;E-mail page comments to the <a href="mailto:safetyandwellbeing@uts.edu.au">Safety &amp; Wellbeing Branch</a><br />
<a href="http://www.uts.edu.au/disclaimer.html">Disclaimer</a> |
<a href="http://www.uts.edu.au/privacy.html">Privacy</a> |
<a href="http://www.uts.edu.au/copyright.html">Copyright</a> |
<a href="http://www.uts.edu.au/accessibility.html">Accessibility</a> |
<a href="http://www.gsu.uts.edu.au/policies/webpolicy.html">Web policy</a> |
<a href="http://www.uts.edu.au/">UTS homepage</a>
</div>
</body>
</html>