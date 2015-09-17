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

</HEAD>

<BODY>
<table width="100%" border=0 align="CENTER" valign="MIDDLE">
  <tr valign="middle"> 
    <td width="150">Chemical Name:</td>
<form method="post" action="SearchChemicals.asp" name="loadlocations">      
<input type="hidden" name="hdnchemicalname">
<input type="hidden" name="hdnCAS1">
<input type="hidden" name="hdnCAS2">
<input type="hidden" name="hdnCAS3">
<input type="hidden" name="hdnChkSort">

<td height="30">
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
    <td width="250">CAS#:</td>
    </tr>
<form action="SearchChemicalsResults.asp" method=post enctype =  "application/x-www-form-urlencoded" target="Results" name="search">
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

    <tr> 
      <td width="100">
        <input type="TEXT" name="txtChemicalName" size=17 maxlength=50 value="<%= getChemicalName %>">
      </td>

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
strSQL= "SELECT tblLocation.numLocationID, tblLocation.numStoreLocationID, tblStoreLocation.strStoreLocation, tblLocation.strStoreNotes FROM tblLocation, tblStoreLocation WHERE tblStoreLocation.numBuildingID = "
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
        </select>
      </td>
      <td width="200"> 
        <input name="txtCAS1" style="HEIGHT: 22px" maxlength=5 size=5 value="<%= getCAS1 %>">
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