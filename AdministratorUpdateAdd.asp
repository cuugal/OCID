<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<html>
<head>
	<title>Administrator</title>
<% 	
Dim strLoginID
	strLoginID = lcase(session("strLoginID"))
	if strLoginID <> "admin" then
		Response.Write "Only the Administrator has access to these pages. Please contact the administrator if you think you should be able to access this."
		Response.End
	end if
	 %>
	
	<SCRIPT LANGUAGE=javascript>
<!--
function AddLocationID(obj){
	var LocationID
	var Index
	//LocationID = document.frmChooseLocation.cboLocation.value
	Index = document.frmChooseLocation.cboLocation.selectedIndex
	LocationID = document.frmChooseLocation.cboLocation.options[Index].value
	if ((LocationID == "") || (LocationID == null)){
		alert ("Please choose a Location")
		return false
		}
	obj.hdnNumLocationID.value = LocationID
	return true
}
	
//-->
</SCRIPT>

<script language="JavaScript">
<!--
function locations() {
  


 	document.loadlocations.submit();
}
//-->
</script>
</head>

<body bgcolor="#F9C7C7">
<table WIDTH="100%" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
<tr>
<form method="post" action="AdministratorUpdateAdd.asp" name="loadlocations"> 
	  </form>	
	 
<font color="#FFFFFF"> 
<td>	<form onSubmit="return AddLocationID(this)" ACTION="UpdateLocation.asp" METHOD="POST" TARGET="Results" NAME="frmUpdateLocation" LANGUAGE=javascript>

	<input TYPE="SUBMIT" NAME="ACTION" VALUE="Update this Location" style="float: right">
		<input TYPE="hidden" NAME="hdnNumLocationID">
		<input type="hidden" name="hdnNumCampusID" value=<%= numCampusID%>>
		<input type="hidden" name="hdnNumBuildingID" value=<%= numBuildingID%>>		
</form></font></td>	

	
	<td width="190">
	<form ACTION="ChooseLoginIDAdd.asp" METHOD="POST" TARGET="Results" NAME="frmAddLocation">
	
	<input TYPE="SUBMIT" NAME="ACTION" VALUE="Add a New Location" style="float: right"></form></td>
    
    <td width="171"><form ACTION="EditPreferences.asp" METHOD="POST" TARGET="Results" NAME="frmEditPre">
	   <input TYPE="SUBMIT" NAME="ACTION" VALUE="Edit Building Setup" style="float: left">
    </form> </td>
	
<td>
<FORM action="" method=POST name=frmChooseLocation> 

</FORM>
    
			</td>
	
	

</tr></table>

</body>
</html>