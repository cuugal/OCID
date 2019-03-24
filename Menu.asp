<!-- This document was created with HomeSite 2.5 -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<%@ Language=VBScript%>
<!--#INCLUDE FILE="DbConfig.asp"-->
<html>
<script language="JavaScript">
<!--
function locations() {
	
 	document.loadlocations.submit();
}

function clearMenu() 
{
document.loadlocations.cboCampus.value =0
document.loadlocations.cboBuildingLocation.value =0
document.loadlocations.cboLocation.value =0

}
/*function campuses() {

	document.loadlocations.hdnchemicalname.value = document.search.txtChemicalName.value;
 	document.loadlocations.submit();
}*/
//-->
</script>
<head>


	<title>Menu</title>
<script LANGUAGE="javascript">
<!--

function ChangeResults(page){
	parent.frames["Results"].location.href = page
	return true
	}
//-->
</script>

<SCRIPT LANGUAGE="JavaScript">
<!--
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1); 
function printPage(frame, arg) {
  if (frame == window) {
    printThis();
  } else {
    link = arg; // a global variable
     printFrame(frame);
  }
  return false;
}

function printThis() {
  if (pr) { // NS4, IE5
    window.print();
  } else if (da && !mac) { // IE4 (Windows)
    vbPrintPage();
  } else { // other browsers
    alert("Sorry, your browser doesn't support this feature.");
  }
}

function printFrame(frame) {
  if (pr && da) { // IE5
    frame.focus();
    window.print();
    link.focus();
  } else if (pr) { // NS4
    frame.print();
  } else if (da && !mac) { // IE4 (Windows)
    frame.focus();
    setTimeout("vbPrintPage(); link.focus();", 100);
  } else { // other browsers
    alert("Sorry, your browser doesn't support this feature.");
  }
}
if (da && !pr && !mac) with (document) {
  writeln('<OBJECT ID="WB" WIDTH="0" HEIGHT="0" CLASSID="clsid:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>');
  writeln('<' + 'SCRIPT LANGUAGE="VBScript">');  
  writeln('Sub window_onunload');
  writeln('  On Error Resume Next');  
  writeln('  Set WB = nothing');
  writeln('End Sub'); 
  writeln('Sub vbPrintPage');
  writeln('  OLECMDID_PRINT = 6');
  writeln('  OLECMDEXECOPT_DONTPROMPTUSER = 2');
  writeln('  OLECMDEXECOPT_PROMPTUSER = 1');  
  writeln('  On Error Resume Next');
  writeln('  WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER');
  writeln('End Sub');  
  writeln('<' + '/SCRIPT>');
}
// --></SCRIPT>

	
</head>

<style type="text/css">
H3, H4, p {margin-top: 0px}
</style>

<body BGCOLOR="#FFFFFF" LINK="#FFFFFF" VLINK="#FFFFFF" ALINK="#FFFF00">
<form  method="POST" name = loadlocations action ="Menu.asp" >
<table WIDTH="100%" BORDER="0" VALIGN="TOP" cellpadding="0" cellspacing ="0" bordercolorlight="#000000" bordercolordark="#000000">
<tr>
	<td ALIGN="left" VALIGN="TOP" bgcolor="#FFFFFF" width="238">
	<img border="0" src="uts-logo.gif" width="130" height="29" align="left">
	<img border="0" src="ocid-logo.gif" width="96" height="29" align="left">
	</td>
	<td align="left" width="12">
	<font size="2" >v4.1</font>
	</td>
	
	<td ALIGN="CENTER" VALIGN="TOP" bgcolor="#FFFFFF"><font FACE="Arial" COLOR=#3333ff>
		<a HREF="SearchChemicals.asp" onclick="ChangeResults('NewSearch.html')" TARGET="Search" NAME="hplSearchChemicals"><h3>
	<font color="#000080">Search Chemicals</font></h3></a></td>
	
	<!--td ALIGN="CENTER" VALIGN="TOP" bgcolor="#FFFFFF"><font FACE="Arial" COLOR="Orange">
	<a HREF="SearchRiskAssessment.asp" onclick="ChangeResults('NewSearchRiskAssessment.html')"TARGET="Search" NAME="hplSearchRiskAssessment"><h3>
	<font color="#008000">Risk Assessment</font></h3></a></td-->
	
	<td ALIGN="CENTER" VALIGN="TOP" bgcolor="#FFFFFF"><font FACE="Arial" COLOR="Green">
	<a HREF="UpdateAdd.asp" onclick="ChangeResults('NewUpdateAdd.html')" TARGET="Search" NAME="hplUpdateAdd"><h3>
	<font color="#FF6600">Update &amp; Add</font></h3></a></td>
	

	<td ALIGN="CENTER" VALIGN="TOP" bgcolor="#FFFFFF"><font FACE="Arial" COLOR=#ff3300>
	<a HREF="AdministratorUpdateAdd.asp" onclick="ChangeResults('NewAdministratorUpdateAdd.html')"TARGET="Search" NAME="hplAdmin"><h3>
	<font color="#FF0000">Admin</font></h3></a></td>

	<td ALIGN="CENTER" VALIGN="TOP" bgcolor="#FFFFFF"><A HREF="#" onClick="return printPage(parent.Results, this)"><h4>
	<font FACE="Arial" color="#000000">PRINT</font></h4></A></td>
</tr>
</table>
<table border="0" width="100%" id="table1" bordercolor="#FFFFFF" style="border-collapse: collapse" bgcolor="#FFFFFF">
	<tr>
		<td width="14%" align="right" valign="top">Building:&nbsp&nbsp</td> 
		<td width="14" align="left" valign="top">
		<%
		
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
		%><font color="#FFFFFF"><select name="cboCampus" onChange="javascript:locations()">
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
        </select></font></td>
		
		<td width="14%" align="right" valign="top">Floor:&nbsp&nbsp</td>
		<td width="14" align="left" valign="top">
<%
Dim numBuildingID
Dim buildingLocation
Dim strBuildingSQL
Dim connBuilding


set connBuilding = Server.CreateObject("ADODB.Connection")
connBuilding.open constr
set buildingLocation = Server.CreateObject("ADODB.Recordset")
strBuildingSQL = "SELECT numBuildingID, strBuildingName FROM tblBuilding WHERE numCampusID="
strBuildingSQL = strBuildingSQL + cstr(numCampusID) + " ORDER BY strBuildingName"
buildingLocation.Open strBuildingSQL, connBuilding, 3, 3
'response.write(strBuildingSQL)

numBuildingID = cint(request.form("cboBuildingLocation"))
if numBuildingID = "" then
	numBuildingID = 0
end if
%><font color="#FFFFFF"><select name="cboBuildingLocation" onChange="javascript:locations()">
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
</select></font> </td>
		
<%
Dim rsLocation
Dim strSQL
Dim conn 

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsLocation = Server.CreateObject("ADODB.Recordset")
strSQL= "SELECT tblLocation.numLocationID, tblLocation.numStoreLocationID, tblStoreLocation.strStoreLocation, tblLocation.strStoreNotes FROM tblLocation, tblStoreLocation WHERE tblStoreLocation.numBuildingID = "
strSQL= strSQL + cstr(numBuildingID) + " AND tblLocation.numStoreLocationID = tblStoreLocation.numStoreLocationID ORDER BY tblStoreLocation.strStoreLocation, tblLocation.numStoreTypeID"
rsLocation.Open strSQL, conn, 3, 3

numLocationID = cint(request.form("cboLocation"))
if numLocationID = "" then
	numLocationID = 0
end if
%>		
<td width="14%" align="right" valign="top">Room:&nbsp&nbsp</td>
		<td width="14%" align="left" valign="top"><font color="#FFFFFF">
		 <select name="cboLocation" onChange="javascript:locations()">
          <option value="0" <% if numLocationID = 0 then
		  response.Write "selected"
		  end if %>>All Rooms</option>
          <% do while not rsLocation.EOF %>
         <option value="<%=rsLocation("numLocationID")%>" <% if numLocationID = rsLocation("numLocationID") then
		  response.Write "selected"
		  end if %>	>
		  
		  <%= rsLocation("strStoreLocation") + ", " + rsLocation("strStoreNotes") + ", (" + CStr(rsLocation("numLocationID")) + ")" %></option>
                   
        
          <%	rsLocation.MoveNext
	loop 
	rsLocation.Close
	set rsLocation = nothing
	conn.close
	set conn = nothing
	

%> 
        </select></font>
        <%
        '----------------------code to insert the combo values into session variables-----------------------------------
         session("numCampusId")= numcampusId
         session("numBuildingId")= numbuildingId
         session("numLocationId")= numLocationId
        '---------------------------------------------------------------------------------------------------------------
        
        %>
        </form>
<form action="SearchLocationResults.asp" method=post enctype="application/x-www-form-urlencoded" name="searchLocation" ID="Form1">
		<%
		' response.write(session("numCampusId"))
        ' response.write(session("numBuildingId"))
        ' response.write(session("numLocationId"))
  

		%>
		<td width="14%" align="left" valign="top">
         
			<input type="button" value="Clear" name="btnClear" onclick =clearMenu();>
		</td>
</form>		
	</tr>
</table>

</body>
</html>