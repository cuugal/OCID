<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<html>
<head>
	<title>Update and Add information</title>

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
  

	document.loadlocations.hdnChkSort.value = document.frmUpdateChemicals.chkSortEditByName.checked;
 	document.loadlocations.submit();
}
//-->
    </script>

<body>
<table WIDTH="100%" BORDER="0" ALIGN="CENTER" VALIGN="MIDDLE">
<tr>
<form method="post" action="UpdateAdd.asp" name="loadlocations">
<input type="hidden" name="hdnChkSort">

	  <td>
Campus: <% 

Dim numCampusID
Dim campusLocation
Dim strCampusSQL
Dim connCampus
'Dim constr2
'constr2 = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

set connCampus = Server.CreateObject("ADODB.Connection")
connCampus.open constr

strCampusSQL= "SELECT numCampusID, strCampusName FROM tblCampus"

set campusLocation = Server.CreateObject("ADODB.Recordset")
'strSQL= strSQL + strBuildingLocationID + "' ORDER BY strStoreLocation, strStoreType"
campusLocation.Open strCampusSQL, connCampus, 3, 3

numCampusID = cint(request.form("cboCampus"))
if numCampusID = "" then
	numCampusID = 0
end if
'response.Write(cstr(numCampusID))
%> 
        <select name="cboCampus" onChange="javascript:locations()">
          <option value="0" 
		  <% if numCampusID = 0 then
		  response.Write "selected"
		  end if %>
		  >All</option>
          <% do while not campusLocation.EOF %>
          <option value="<%=campusLocation("numCampusID")%>"
		  <% if campusLocation("numCampusID") = numCampusID then
		  response.Write "selected"
		  end if %>		  
		  ><%= campusLocation("strCampusName") %></option>
          <% campusLocation.MoveNext
	loop

	campusLocation.Close
	set campusLocation = nothing
	connCampus.close
	set connCampus = nothing
%> 
        </select>
Building: 
<% 
Dim numBuildingID
Dim buildingLocation
Dim strBuildingSQL
Dim connBuilding
'Dim constr3

'numCampusID = request.form("cboCampus")
'if numCampusID = "" then
'	numCampusID = 0
'end if
'response.Write(cstr(numCampusID))

'constr3 = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

set connBuilding = Server.CreateObject("ADODB.Connection")
connBuilding.open constr
set buildingLocation = Server.CreateObject("ADODB.Recordset")
strBuildingSQL = "SELECT numBuildingID, strBuildingName FROM tblBuilding WHERE numCampusID="
strBuildingSQL = strBuildingSQL + cstr(numCampusID) + " ORDER BY numBuildingID"
buildingLocation.Open strBuildingSQL, connBuilding, 3, 3
'response.write(strBuildingSQL)

numBuildingID = cint(request.form("cboBuildingLocation"))
if numBuildingID = "" then
	numBuildingID = 0
end if
'response.Write(cstr(numCampusID))

%> 
<select name="cboBuildingLocation" onChange="javascript:locations()">
          <option value="0"
		  <% if numBuildingID = 0 then
		  response.Write "selected"
		  end if %>
		  >All</option>
		<%	do while not buildingLocation.EOF %> 
         <option value="<%= buildingLocation("numBuildingID")%>"
		  <% if numBuildingID = buildingLocation("numBuildingID") then
		  response.Write "selected"
		  end if %>		 
		 ><%= buildingLocation("strBuildingName") %></option>
          <%	buildingLocation.MoveNext
	loop
	buildingLocation.Close
	set buildingLocation = nothing
	connBuilding.close
	set connBuilding = nothing
%>
</select>
</td>	
</form>
	<form onsubmit="return AddLocationID(this)" ACTION="UpdateChemicals.asp" METHOD="POST" TARGET="Results" NAME="frmUpdateChemicals">

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
		<font size="-1">Sort by name</font><INPUT TYPE="Checkbox" Name="chkSortEditByName"<%= getChkSort %> value="ON">
		<input TYPE="SUBMIT" NAME="btnUpdateChemical" VALUE="Edit List of Chemicals at this Location ">
		<input TYPE="hidden" NAME="hdnNumLocationID">
		</td>
	<input type="hidden" name="hdnbuildinglocation" value="<%= numBuildingID %>">
	<input type="hidden" name="hdncampuslocation" value="<%= numCampusID %>">
	</form>	
</tr>
<tr>
	<FORM action="" method=POST name=frmChooseLocation>

<td> Room: <% 

'Dim numBuildingID
Dim rsLocation
Dim strSQL
Dim conn 
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

'numBuildingID = request.form("cboBuildingLocation")
'if numBuildingID = "" then
'	numBuildingID = 0
'end if
'response.Write(cstr(numBuildingID))

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsLocation = Server.CreateObject("ADODB.Recordset")
strSQL= "SELECT tblLocation.numLocationID, tblLocation.numStoreLocationID, tblStoreLocation.strStoreLocation, tblLocation.strStoreNotes FROM tblLocation, tblStoreLocation WHERE tblLocation.numBuildingID = "
strSQL= strSQL + cstr(numBuildingID) + " AND tblLocation.numStoreLocationID = tblStoreLocation.numStoreLocationID ORDER BY tblStoreLocation.strStoreLocation, tblLocation.numStoreTypeID"
rsLocation.Open strSQL, conn, 3, 3
%> 
        <select name="cboLocation">
          <option value="0">All Rooms</option>
          <% do while not rsLocation.EOF %>
         <option value="<%=rsLocation("numLocationID")%>"><%= rsLocation("strStoreLocation") + ", " + rsLocation("strStoreNotes") %></option>
          <%	rsLocation.MoveNext
	loop 
	rsLocation.Close
	set rsLocation = nothing
	conn.close
	set conn = nothing
	

%> 
        </select></td>
</FORM>
    
   	<form onsubmit="return AddLocationID(this)" ACTION="AddChemical.asp" METHOD="POST" TARGET="Results" NAME="frmAddChemicals">
	<td align="right"><input TYPE="SUBMIT" NAME="btnAddChemical" VALUE="  Add a New Chemical to this Location">
	<input TYPE="hidden" NAME="hdnNumLocationID">
	</td>
	</form>
</tr></table>
</body>
</html>