<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<% dim strDate
strDate = DanDate(Date, "%d/%m/%Y" )

Sub RefreshMenu()
	
	Dim Output
	Output = "<SCRIPT LANGUAGE=javascript><!-- " + chr(13)
	Output = Output + "parent.frames['Search'].location.reload() " + chr(13)
	Output = Output + "//--></SCRIPT>"
	
	Response.Write(OutPut)
	
End Sub

Dim strLoginID
strLoginID = Request.Form("cboLoginID")

Call checkAction

Dim strStoreNote
Dim strStoreManager
Dim strStoreLocation
Dim strLicensedDepot
Dim numCampus
Dim numBuilding
Dim numStoreTypeID
Dim dtmLastUpdated
Dim rsAddStore
Dim rsStore
Dim strSQLAdd2
Dim strSQLAdd

'Dim strPassword
'Dim strPassword2

'Dim strLoginIDChg

'Dim strLoginID
Dim strPassword
Dim strAddNewLocation
Dim strSQL
Dim conn 

if request("txtLoginID") = "" then
'	strLoginID = Request.Form("cboLoginID")
	if strLoginID <> "" then
	if strLoginID <> "NewLoginID" then
		set conn = Server.CreateObject("ADODB.Connection")
		conn.open constr
		set rsAccess = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM tblACCESS, tblLocation WHERE tblLocation.strLoginID = tblAccess.strLoginID AND tblAccess.strLoginID = '"
		strSQL = strSQL + strLoginID + "'"
		rsAccess.Open strSQL, conn, 3, 3
		if rsAccess.EOF then
			rsAccess.close
			set rsAccess = nothing
			set rsAccess = Server.CreateObject("ADODB.Recordset")			
			strSQL = "SELECT * FROM tblACCESS WHERE tblAccess.strLoginID = '" + strLoginID + "'"
			rsAccess.Open strSQL, conn, 3, 3
			strPassword = rsAccess("strPassword")
			strStoreManager = rsAccess("strFirstName") + " " + rsAccess("strLastName")
		else
			strPassword = rsAccess("strPassword")
			strStoreManager = rsAccess("strStoreManager")
		end if
	else
		strLoginID = ""
'		strPassword=""
'		strStoreManager=""
	end if
	end if
end if
	
'Dim strBuildingLocationID


strAddNewLocation = ucase(Request.Form("btnAddLocation"))
if (strAddNewLocation = "ADD NEW LOCATION") THEN

Dim strPassword2
strPassword = LCASE(Request.Form("txtPassword"))
strPassword2 = LCASE(Request.Form("txtPassword2"))
if (strPassword <> strPassword2) then
	Response.Write "Passwords do not match, press the back key on the browser and enter again"
	Response.End
end if

Dim rsLocation

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsLocation = Server.CreateObject("ADODB.Recordset")

	numCampus = request.form("cboCampus")
	numBuilding = request.form("cboBuildingLocation")
'	strStoreType = Request.Form("cboStoreType")
	strStoreNote = Request.Form("txtStoreNotes")
	if strStoreNote = "" then
		strStoreNote = " "
	end if
	strStoreManager = Request.Form("txtStoreManager")
	strStoreLocation = Request.form("txtNewStoreLocation")
	strLoginID = Request.Form("txtLoginID")
	dtmLastUpdated = strDate
	numStoreTypeID = request.form("cboStoreType")
	strLicensedDepot = request.form("chkLicensedDepot")
	if strLicensedDepot <> "true" then
		strLicensedDepot = "false"
	end if 


	For each item in Request.Form

	if item <> "txtStoreNotes" AND item <> "txtMaxStorage" AND item <> "txtDepotClass" then
		if request.form(item) = "" then
			Response.Write "All fields must contain values. Please go back and fill in "
			Response.Write item
			Response.End
		end if
	end if
	Next

	if request.form("cboCampus") = "0" then
			Response.Write "All fields must contain values. Please go back and select a "
			Response.Write "building"
			Response.End	
	end if
	if request.form("cboBuildingLocation") = "0" then
			Response.Write "All fields must contain values. Please go back and select a "
			Response.Write "floor"
			Response.End	
	end if
	if request.form("cboStoreType") = "0" then
			Response.Write "All fields must contain values. Please go back and select a "
			Response.Write "store type"
			Response.End	
	end if


strSQL = "SELECT tblAccess.strLoginID FROM tblAccess WHERE tblAccess.strLoginID = '" + InjectionEncode(strLoginID) + "'"
set rsAccess = Server.CreateObject("ADODB.Recordset")
rsAccess.Open strSQL, conn, dynaset, 3	

if rsAccess.EOF then
	strSQL = "INSERT INTO tblAccess (strLoginID, strPassword) VALUES "
	strSQL = strSQL + "('" + InjectionEncode(strLoginID) + "', '" + InjectionEncode(strPassword) + "')"
else
	strSQL = "UPDATE tblAccess SET strPassword = '" + InjectionEncode(strPassword) + "' WHERE (strLoginID = '" + InjectionEncode(strLoginID) + "')"
end if
rsAccess.close
rsAccess.Open strSQL, conn, 3, 3
set rsAccess = Nothing

'strSQL = "INSERT INTO tblStoreLocation "
'strSQL = strSQL + "(numBuildingID, strStoreLocation) "
'strSQL = strSQL + "VALUES (" + cstr(numBuilding) + ""
'strSQL = strSQL + ", '" + strStoreLocation + "')"
'set rsStoreLocation = Server.CreateObject("ADODB.Recordset")
'rsStoreLocation.Open strSQL, conn, dynaset, 3	
'set rsStoreLocation = nothing

'strSQL = "SELECT numStoreLocationID FROM tblStoreLocation"
'set rsStoreID = Server.CreateObject("ADODB.Recordset")
'rsStoreID.Open strSQL, conn, dynaset, 3

'Dim numStoreID

'do while not rsStoreID.EOF
'numStoreID = rsStoreID("numStoreLocationID")
'rsStoreID.movenext
'loop
'set rsStoreLocation = nothing
		strSQLAdd = "INSERT INTO tblStoreLocation "
		strSQLAdd = strSQLAdd + "(strStoreLocation, numBuildingID) "
		strSQLAdd = strSQLAdd + "VALUES ('" + InjectionEncode(strStoreLocation) + "', " + cstr(numBuilding) + ")"	
		set rsAddStore = Server.CreateObject("ADODB.Recordset")
		rsAddStore.Open strSQLAdd, conn, dynaset, 3

		strSQLAdd2 = "SELECT * FROM tblStoreLocation"		
		set rsStore = Server.CreateObject("ADODB.Recordset")
		rsStore.Open strSQLAdd2, conn, dynaset, 3	
	
     do while not rsStore.EOF
		strStoreLocation = rsStore("numStoreLocationID")
		rsStore.MoveNext
      loop

Dim numLID 
numLID = "null"
strSQL = "INSERT INTO tblLocation "
strSQL = strSQL + "(numCampusID, numBuildingID, numStoreLocationID, numStoreTypeID, "
strSQL = strSQL + "strStoreManager, strStoreNotes, strLoginID, dtmLastUpdated) "
'strSQL = strSQL + "VALUES ('" + strBuildingLocation + "'"
'strSQL = strSQL + ", '" + strBuildingLocationID + "'"
strSQL = strSQL + "VALUES (" + cstr(numCampus) + ""
strSQL = strSQL + ", " + cstr(numBuilding) + ""
strSQL = strSQL + ", " + cstr(strStoreLocation) + ""
strSQL = strSQL + ", " + cstr(numStoreTypeID) + ""
'strSQL = strSQL + ", " + strLicensedDepot + ""
strSQL = strSQL + ", '" + InjectionEncode(strStoreManager) + "'"
strSQL = strSQL + ", '" + InjectionEncode(strStoreNote) + "'"
strSQL = strSQL + ", '" + InjectionEncode(strLoginID) + "'"
strSQL = strSQL + ", '" + dtmLastUpdated + "')"

'rsLocation.Open strSQL, conn, dynaset, 3
rsLocation.Open strSQL, conn, dynaset, 3
set rsLocation = nothing
conn.close
set conn = nothing
RefreshMenu()
Response.Write "The location has been added"
Response.End

ELSE

END IF
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Add Location</TITLE>
<script language="JavaScript">
function locations() {
//	document.frmUpdateLocation.hdnchemicalname.value = document.search.txtChemicalName.value;
//	document.frmUpdateLocation.cboStoreType.value = document.frmUpdateLocation.cboStoreType.value
//	document.frmUpdateLocation.txtStoreType.value = document.frmUpdateLocation.txtStoreTest.value;
	document.frmAddLocation.action.value = "addLocation";
 	document.frmAddLocation.submit();
}

function clearUp() {
	document.frmAddLocation.cboCampus.value = 0;
}
</script>

</HEAD>
<BODY>
<% 
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

'strLoginID = Request.Form("cboLoginID")
'if strLoginID <> "NewLoginID" then
'	set conn = Server.CreateObject("ADODB.Connection")
'	conn.open constr
'	set rsAccess = Server.CreateObject("ADODB.Recordset")
'	strSQL = "SELECT * FROM tblACCESS WHERE strLoginID = '"
'	strSQL = strSQL + strLoginID + "'"
'	rsAccess.Open strSQL, conn, 3, 3
'	strPassword = rsAccess("strPassword")
'else
'	strLoginID=""	

'end if

 %>
<DIV align=center>
<BR><FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
<BR>
<FORM action="AddLocation.asp" method=POST name=frmAddLocation>
<input type="hidden" name=action value="abc">

<TABLE align=center border=0 cellPadding=1 cellSpacing=10>
    <TR>
		<TD><STRONG><FONT color=red 
            face="">Add a Location</FONT></STRONG><BR><BR></TD></TR>
<TR>
        <TD>Location:</TD>
        <TD>
			<% 
'strBuildingLocationID = rsLocation("strBuildingLocationID")
%> 
Building: <% 

Dim numCampusID
Dim campusLocation
Dim strCampusSQL
Dim connCampus
Dim inCampusID
'Dim constr2
'constr2 = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")


set connCampus = Server.CreateObject("ADODB.Connection")
connCampus.open constr

strCampusSQL= "SELECT numCampusID, strCampusName FROM tblCampus"
'strSQL= strSQL + strBuildingLocationID + "' ORDER BY strStoreLocation, strStoreType"

set campusLocation = Server.CreateObject("ADODB.Recordset")
campusLocation.Open strCampusSQL, connCampus, 3, 3

'inCampusID = cint(request.form("hdnNumCampusID"))
numCampusID = cint(request.form("cboCampus"))
'if numCampusID = "0" then
'	numCampusID = inCampusID
'end if
'response.Write(cstr(numCampusID) + "T")
%> 
        <select name="cboCampus" onChange="javascript:locations()">
		  <option value="0">Please select</option>
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
Floor: 
<% 
Dim numBuildingID
Dim buildingLocation
Dim strBuildingSQL
Dim connBuilding
Dim inBuildingID
'Dim constr3

'numCampusID = request.form("cboCampus")
'if numCampusID = "" then
'	numCampusID = 0
'end if
'response.Write(cstr(numCampusID))

'constr3 = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

'response.Write(cstr(numCampusID))

'inBuildingID = cint(request.form("hdnNumBuildingID"))

numBuildingID = cint(request.form("cboBuildingLocation"))
'if numBuildingID = "0" then
'	numBuildingID = inBuildingID
'end if

set connBuilding = Server.CreateObject("ADODB.Connection")
connBuilding.open constr
set buildingLocation = Server.CreateObject("ADODB.Recordset")
strBuildingSQL = "SELECT numBuildingID, strBuildingName FROM tblBuilding WHERE numCampusID = "
strBuildingSQL = strBuildingSQL + cstr(numCampusID) + " ORDER BY numBuildingID"
buildingLocation.Open strBuildingSQL, connBuilding, 3, 3
'response.write(strBuildingSQL)
%> 
<select name="cboBuildingLocation">
		<option value="0">Please select</option>
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

</TD></TR>
<TR>
        <TD>Store Location (Room):</TD>
        <TD> 
<input type="text" name="txtNewStoreLocation" size="15" value="<%= strStoreLocation %>">
<%
'Dim numLocationIDcbo
'Dim rsRoomLocation
'Dim strSQLRoom
'Dim connRoom
'Dim inLocationID

'set connRoom = Server.CreateObject("ADODB.Connection")
'connRoom.open constr
'strSQLRoom = "SELECT tblLocation.numLocationID, tblLocation.numStoreLocationID, tblStoreLocation.strStoreLocation, tblLocation.strStoreNotes FROM tblLocation, tblStoreLocation WHERE tblLocation.numBuildingID = "
'strSQLRoom = strSQLRoom + cstr(numBuildingID) + " AND tblLocation.numStoreLocationID = tblStoreLocation.numStoreLocationID ORDER BY tblStoreLocation.strStoreLocation"
'strSQLRoom = "SELECT numStoreLocationID, strStoreLocation FROM tblStoreLocation WHERE numBuildingID = "
'strSQLRoom = strSQLRoom + cstr(numBuildingID) + " ORDER BY strStoreLocation"
'set rsRoomLocation = Server.CreateObject("ADODB.Recordset")
'rsRoomLocation.Open strSQLRoom, connRoom, 3, 3
%>
<!--		  <input type="text" name=txtStoreLocation value="<%'= strStoreLocation %>"> -->
<font face="Arial,serif" size="2">eg. Room 1111, Potting Shed, Room 5A etc</font></TD></TR>
    
    <TR>
        <TD>Store Type:</TD>
        <TD>
          <% 

Dim rsStoreType
Dim inStoreTypeID

inStoreTypeID = request.form("cboStoreType")
'if inStoreTypeID = "" then
'	inStoreTypeID = rsLocation("numStoreTypeID")
'	response.write cstr("replaced")
'end if

'response.write cstr(inStoreTypeID)

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

strSQL= "SELECT * FROM tblStoreType"

set rsStoreType = Server.CreateObject("ADODB.Recordset")
rsStoreType.Open strSQL, conn, 3, 3
%>
          <select name="cboStoreType" onChange="javascript:locations()">
		  <option value="0">Please select</option>
          <% do while not rsStoreType.EOF %>
		  <option value="<%= rsStoreType("numStoreTypeID") %>"
  		  <% if cint(inStoreTypeID) = rsStoreType("numStoreTypeID") then
		  response.Write "selected"
		  end if %>		
		  ><%= rsStoreType("strStoreType") %></option>
          <%	rsStoreType.MoveNext
	loop 
	rsStoreType.Close
	set rsStoreType = nothing
'	conn.close
'	set conn = nothing
%> 
	</select></td></tr>


    <TR>
	<TD>Store Notes:</TD>
        <TD>
<% 
'strStoreNote = "shttyu"
'response.write strStoreNote
%>
            <INPUT name=txtStoreNotes value="<%= strStoreNote %>"  size="50" maxlength="50">
		<font face="Arial,serif" size="2">Short description eg. LAB, GENERAL, DG STORE, GLASSHOUSE, etc </font>
		</TD></TR>

    <TR>
        <TD>Supervisor:</TD>
        <TD>
            <INPUT name=txtStoreManager style="HEIGHT: 22px; WIDTH: 265px" value="<%= strStoreManager %>"></TD></TR>
    <TR>
</TABLE>
        <font face="Arial,serif" size="2">Enter the Supervisor's Login ID and Password to allow them to update and add chemicals to this location</font>
  <TABLE>
    <TR>
	<TD>Date:</TD>
        <TD>
			<b><font size="2"><%= strDate %></font></b>
			</TD></TR>
      
	<TR>
        <TD>Login ID:</TD>
        <TD>
            <INPUT name=txtLoginID value="<%= strLoginID %>" style="HEIGHT: 22px; WIDTH: 265px"></TD></TR>
    
	<TR>
        <TD>Password:</TD>
        <TD>
            <INPUT type=password value="<%= strPassword %>" name=txtPassword style="HEIGHT: 22px; WIDTH: 265px"></TD></TR>
    <TR>
        <TD>Confirm Password:</TD>
        <TD>
            <INPUT type=password value="<%= strPassword %>" name=txtPassword2 style="HEIGHT: 22px; WIDTH: 265px"></TD></TR>
    
	<TR>
        <TD>&nbsp;</TD>
        <TD><INPUT type="reset" value="Clear Form" name=btnClear>&nbsp;&nbsp;
			<INPUT type="submit" value="Add New Location" name=btnAddLocation></TD></TR></TABLE>
			
</FORM>
</FONT></DIV>

</BODY>
</HTML>
<%
Sub checkAction()
	if request.form("action") = "addLocation" then
		strStoreNote = request("txtStoreNotes")
		strStoreManager = request("txtStoreManager")
		strLicensedDepot = request("chkLicensedDepot")
		strLoginID = request("txtLoginID")
		strPassword = request("txtPassword")
		strPassword2 = request("txtPassword2")
		strStoreLocation = request("txtNewStoreLocation")
	end if
end Sub
%>