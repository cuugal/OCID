<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<html>
<head>
	<title>Update and Add information</title>

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
  

	document.loadlocations.hdnChkSort.value = document.frmUpdateChemicals.chkSortEditByName.checked;
 	document.loadlocations.submit();
}
//-->
</script>

<body bgcolor="#F9E5C7">
<table WIDTH="100%" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
<tr>
<form method="post" action="UpdateAdd.asp" name="loadlocations">
<input type="hidden" name="hdnChkSort">
	
</form>
	<form onsubmit="return AddLocationID(this)" ACTION="UpdateChemicals.asp" METHOD="POST" TARGET="Results" NAME="frmUpdateChemicals">
	<form ACTION="UpdateChemicals.asp" METHOD="POST" TARGET="Results" NAME="frmUpdateChemicals">
<%
Function getChkSort
	If request.form("hdnChkSort") = "true" Then
		getChkSort = " CHECKED "
	Else
		getChkSort = ""
	End If
End Function
%>

	<td align="right">
		<font size="-1">Sort by name</font><INPUT TYPE="Checkbox" Name="chkSortEditByName"<%= getChkSort %>>
		<input TYPE="SUBMIT" NAME="btnUpdateChemical" VALUE="Edit List of Chemicals at this Location ">
		<input TYPE="hidden" NAME="hdnNumLocationID">
		</td>
	
	<input type="hidden" name="hdnbuildinglocation" value="<%= numBuildingID %>">
	<input type="hidden" name="hdncampuslocation" value="<%= numCampusID %>">
	</form>	
    
   	<form onsubmit="return AddLocationID(this)" ACTION="AddChemical.asp" METHOD="POST" TARGET="Results" NAME="frmAddChemicals">
	<td align="right">&nbsp;</td>
	<td align="right"><input TYPE="SUBMIT" NAME="btnAddChemical" VALUE="  Add a New Chemical to this Location">
	<input TYPE="hidden" NAME="hdnNumLocationID">
	</td>
	</form>
</tr></table>
</body>
</html>