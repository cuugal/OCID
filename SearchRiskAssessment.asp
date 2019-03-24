<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<!-- This document was created with HomeSite 2.5 -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<html>
<head>
	<title>Risk Assessment Search</title>

<!-- Added the following admin access warning after hiding link to Search Risk Assessments in the menu -->
	<%
Dim strLoginID
	strLoginID = lcase(session("strLoginID"))
	if strLoginID <> "admin" then
		Response.Write "Only the Administrator has access to these pages. Please contact the administrator if you think you should be able to access this."
		Response.End
	end if
%>

<script language="JavaScript">
<!--
function locations() {
  
	document.loadlocations.hdnchemicalname.value = document.search.txtChemicalName.value;
	document.loadlocations.hdnChkSort.value = document.search.chkSortByName.checked;
 	document.loadlocations.submit();
}
//-->
</script>
</head>

<body bgcolor="#CEF9C7" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<table WIDTH="100%" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
<tr>
<form method="post" action="SearchRiskAssessment.asp" name="loadlocations">      
<input type="hidden" name="hdnchemicalname"> 
<input type="hidden" name="hdnChkSort"> 
	<td></td>
	  <TD>
</TD>
</form> 
</tr>

<form ACTION="RiskAssessmentResults.asp" METHOD="POST" TARGET="Results" NAME="search">

<%
Function getChemicalName
	If request.form("hdnChemicalName") <> "" Then
		getChemicalName = request.form("hdnChemicalName")
	Else
		getChemicalName = "*"
	End If
End Function

Function getChkSort
	If request.form("hdnChkSort") = "true" Then
		getChkSort = " CHECKED "
	Else
		getChkSort = ""
	End If
End Function
%>

<input type="hidden" name="hdnbuildinglocation" value="<%= numBuildingID %>"> 
<input type="hidden" name="hdncampuslocation" value="<%= numCampusID %>"> 
<tr>
 
	<td valign="top" height="33">Chemical Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input TYPE="TEXT" NAME="txtChemicalName" SIZE="20" MAXLENGTH="50" value="<%= getChemicalName %>"></td>
	  <TD height="33" valign="top"></TD>
<td height="33" valign="top"><font size="-1">Sort by name</font> 
	<input type="CHECKBOX" name="chkSortByName" value="true"<%= getChkSort %>>
</td>
<td height="33" valign="top"><input TYPE="SUBMIT" NAME="btnSearchRiskAssessment" VALUE="Search Risk Assessment">
</td>
</tr>
</form>
</table>


</body>
</html>
