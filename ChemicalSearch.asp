
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<HTML>
<HEAD>
	<TITLE>Chemical Search</TITLE>
</HEAD>

<BODY>
<TABLE WIDTH="100%" BORDER=0 ALIGN="CENTER" VALIGN="MIDDLE">
<FORM ACTION="ChemicalSearchResults.html" METHOD="GET" ENCTYPE="application/x-www-form-urlencoded" TARGET="Results"><TR>
	<TD>Chemical Name:</TD>

	<TD>Location:</TD>

	<TD>CAS#:</TD>
</TR>
<TR>
	<TD><INPUT TYPE="TEXT" NAME="txtChemicalName" SIZE=20 MAXLENGTH=50></TD>

	<TD>

<% 
Dim rsLocation
Dim strSQL
Dim conn 

conn = Application("connChemical")
set rsLocation = Server.CreateObject("adodb.Recordset")
strSQL= "SELECT DISTINCT strBuildingLocation FROM tblLocation"
rsLocation.Open strSQL, conn
%>
<SELECT NAME="cboLocation">

<% do while not rsLocation.EOF %>
	<OPTION><%= rsLocation("strBuildingLocation") %>
<%	rsLocation.MoveNext
	loop 
	rsLocation.Close
	set rsLocation = nothing
%>
    </SELECT></TD>
	
	<TD><INPUT name=txtCAS1 style="HEIGHT: 22px" maxlength=5 size=5 > - <INPUT name=txtCAS2 style="HEIGHT: 22px" maxlength=2 size=2 > - <INPUT name=txtCAS3 style="HEIGHT: 22px" maxlength=1 size=1 ></TD>
	
	<TD><INPUT TYPE="SUBMIT" NAME="btnSearchChemical" VALUE="Search Chemical"></TD>

</TR>
</FORM>
</TABLE>


</BODY>
</HTML>
