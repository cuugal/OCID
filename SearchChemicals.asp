<%@ Language=VBScript%>
<!--#INCLUDE FILE="DbConfig.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<HTML>
<HEAD>
	<TITLE>Chemical Search</TITLE>

<script language="JavaScript">
<!--
function locations() {
	document.loadlocations.hdnchemicalname.value = document.search.txtChemicalName.value;
	document.loadlocations.hdnCAS1.value = document.search.txtCAS1.value;
	document.loadlocations.hdnCAS2.value = document.search.txtCAS2.value;
	document.loadlocations.hdnCAS3.value = document.search.txtCAS3.value;
	document.loadlocations.hdnChkSort.value = document.search.chkSortByName.checked;
//	alert(document.loadlocations.hdnChkSort.value);
 	document.loadlocations.submit();
}

/*function campuses() {

	document.loadlocations.hdnchemicalname.value = document.search.txtChemicalName.value;
 	document.loadlocations.submit();
}*/
//-->
</script>

	<base target="_self">

</HEAD>

<BODY bgcolor="#C7E3F9">
<table width="100%" border="0" align="CENTER" valign="MIDDLE">
  <tr valign="middle"> 
    
<form method="post" action="SearchChemicals.asp" name="loadlocations">      
<input type="hidden" name="hdnchemicalname">
<input type="hidden" name="hdnCAS1">
<input type="hidden" name="hdnCAS2">
<input type="hidden" name="hdnCAS3">
<input type="hidden" name="hdnChkSort">


<% Dim numCampusID
   Dim numBuildingID
   Dim numLocationID
   
   numCampusID = session("numCampusID")
   numBuildingID = session("numBuildingID")
   numLocationID = session("numLocationID")
   
%> 


</form>
    
  
<form action="SearchChemicalsResults.asp" method=post enctype="application/x-www-form-urlencoded" target="Results" name="search">
</tr>
<%
Function getChemicalName
	If request.form("hdnChemicalName") <> "" Then
		getChemicalName = request.form("hdnChemicalName")
	Else
		getChemicalName = "*"
	End If
End Function

Function getCAS1
	If request.form("hdnCAS1") <> "" Then
		getCAS1 = request.form("hdnCAS1")
	Else
		getCAS1 = ""
	End If
End Function

Function getCAS2
	If request.form("hdnCAS2") <> "" Then
		getCAS2 = request.form("hdnCAS2")
	Else
		getCAS2 = ""
	End If
End Function

Function getCAS3
	If request.form("hdnCAS3") <> "" Then
		getCAS3 = request.form("hdnCAS3")
	Else
		getCAS3 = ""
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
<input type="hidden" name="hdnRoomlocation" value="<%= numLocationID %>">


    <tr> 
      <td width="279" valign="top">
        Chemical Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="TEXT" name="txtChemicalName" size=17 maxlength=50 value="<%= getChemicalName %>">
      </td>

     
      <td width="508"> 
        &nbsp;&nbsp; 
        CAS#:
		<!-- changed from maxlength=5 size=5 on 28June2013 -->
		<input name="txtCAS1" style="HEIGHT: 22px" maxlength="7" size="7" value="<%= getCAS1 %>">
        - 
        <input name="txtCAS2" style="HEIGHT: 22px" maxlength=2 size=2 value="<%= getCAS2 %>">
        - 
        <input name="txtCAS3" style="HEIGHT: 22px" maxlength=1 size=1 value="<%= getCAS3 %>">
      </td>
      <td width="100"><font size="-1">Sort by name</font> 
		<input type="CHECKBOX" name="chkSortByName"<%= getChkSort %>>
      </td>
      <td width="50">
        <input type="submit" name="btnSearchChemical" value="Search">
      </td>
    </tr>
  </form>
</table>
</BODY>
</HTML>