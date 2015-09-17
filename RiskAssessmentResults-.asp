<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<%
Dim conn
Dim rsRA, rsLocation, rsBuildingCampus
Dim strSQL, strSQL2, strSQL3
Dim strChemicalName
Dim strLocation
Dim strBuildingLocationID, numCampusID
Dim strSortByName
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

strChemicalName = Request.Form("txtChemicalName")
'numLocationID = Request.Form("cboLocation")
'strBuildingLocationID = Request.Form("hdnbuildinglocation")
'numCampusID = Request.Form("hdncampuslocation")


  numLocationID = cstr(session("numLocationID"))
  strBuildingLocationID = cstr(session("numBuildingID"))
  numCampusID = cstr(session("numCampusID")) 
  
strSortByName = Request.form("chkSortByName")
strChemicalName = Replace(strChemicalName,"*","%") 

strSQL = "SELECT * FROM qryChemicalRiskAssessment WHERE "
if numCampusId <> "0" then
	if strBuildingLocationID <> "0" then 
'		rem a Builing location has been chosen
		if numLocationID = "0" then
'			rem a room has NOT been chosen, ie search all rooms in a building
			strSQL = strSQL + "(qryChemicalRiskAssessment.numBuildingID = " + strBuildingLocationID + ") AND "
'			strSQL = strSQL + "IsNumeric(qryChemicalRiskAssessment.numRiskAssessmentID) AND "
		else
'			rem a room has been chosen
			strSQL = strSQL + "(qryChemicalRiskAssessment.numLocationID = " + numLocationID + ") AND "
'			strSQL = strSQL + "IsNumeric(qryChemicalRiskAssessment.numRiskAssessmentID) AND "
		end if
	else
'		rem all locations are searched - return only chemicals that have a risk assessment
		strSQL = strSQL + "(qryChemicalRiskAssessment.numCampusID = " + numCampusID + ") AND "	
'		strSQL = strSQL + "IsNumeric(qryChemicalRiskAssessment.numRiskAssessmentID) AND "
	end if 
end if

strSQL = strSQL + "(qryChemicalRiskAssessment.strChemicalName LIKE '" + InjectionEncode(strChemicalName) + "') ORDER BY "

if strSortByName <> "true" then
	if numCampusID = "0" then
 		strSQL = strSQL + "qryChemicalRiskAssessment.strCampusName, "
	end if
	
	if strBuildingLocationID = "0" then
 		strSQL = strSQL + "qryChemicalRiskAssessment.strBuildingName, "
  	end if

 	if numlocationID = "0" then
 		strSQL = strSQL + "qryChemicalRiskAssessment.strStoreLocation, qryChemicalRiskAssessment.strStoreType, qryChemicalRiskAssessment.strStoreNotes, "
  	end if

	if numlocationID <> "0" then
 		strSQL = strSQL + "qryChemicalRiskAssessment.strSpecificLocation, "
  	end if
  
end if

' if ((numlocationID = "0") and strSortByName <> "true") then
' 	strSQL = strSQL + " ORDER BY tblStoreLocation.strStoreLocation, qryChemicalRiskAssessment.strChemicalName"
'else
	strSQL = strSQL + "qryChemicalRiskAssessment.strChemicalName"
'end if
'if numlocationID = "0" then
'	strSQL = strSQL + " ORDER BY tblLocation.strStoreLocation"
'else
'	strSQL = strSQL + " ORDER BY qryChemicalRiskAssessment.strChemicalName"
'end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

set rsRA = Server.CreateObject("ADODB.Recordset")
rsRA.Open strSQL, conn

if rsRA.EOF then
	Response.Write "<BR><DIV align='center'><FONT color='red' face='Arial'>There are no results for that search, please make your search less specific (try using *, for wildcard search).</FONT></DIV>"
	Response.End
end if

strSQL2 = "SELECT * FROM tblLocation, tblStoreType, tblStoreLocation WHERE tblStoreType.numStoreTypeID = tblLocation.numStoreTypeID AND "
strSQL2 = strSQL2 + "tblLocation.numStoreLocationID = tblStoreLocation.numStoreLocationID"


if strBuildingLocationID <> "0" then 
		rem a Builing location has been chosen
	if numLocationID = "0" then
			rem a room has NOT been chosen, ie search all rooms at in a building
		strSQL2 = strSQL2 + " AND (tblLocation.numBuildingID = " + cstr(rsRA("numBuildingID")) + ")"
	else
			rem a room has been chosen
		strSQL2 = strSQL2 + " AND (tblLocation.numLocationID = " + cstr(rsRA("numLocationID")) + ")"
	end if

end if

set rsLocation = Server.CreateObject("ADODB.Recordset")
rsLocation.Open strSQL2, conn

strSQL3 = "SELECT tblBuilding.strBuildingName, tblCampus.strCampusName FROM tblCampus, tblBuilding "
strSQL3 = strSQL3 + "WHERE tblCampus.numCampusID = " + cstr(rsLocation("numCampusID"))  + " AND "
strSQL3 = strSQL3 + "tblBuilding.numBuildingID = " + cstr(strBuildingLocationID)  + " AND tblCampus.numCampusID = tblBuilding.numCampusID"

set rsBuildingCampus = Server.CreateObject("ADODB.Recordset")
rsBuildingCampus.Open strSQL3, conn

Dim numRiskAssessment

%>
<HTML>
<HEAD>
	<TITLE>Risk Assessment Results</TITLE>
</HEAD>

<BODY>
<FORM>
<% if numLocationID <> "0" then %>
<p align="left">Store Manager: <%= rsRA("strStoreManager") %> <Br>
Location: <%= rsBuildingCampus("strCampusName") + ", " + rsBuildingCampus("strBuildingName") + ", " + rsLocation("tblStoreLocation.strStoreLocation")+ ",  " + rsLocation("strStoreType") + ", " + rsLocation("strStoreNotes") %><BR>
Last Updated: <%= rsRA("dtmLastUpdated") %></p>
<% end if %>

<TABLE WIDTH="100%" BORDER=1 ALIGN="center" VALIGN = "TOP">
<TR ALIGN="left" VALIGN="top" BGCOLOR="yellow">
	<TD>Name</TD>

<% if numCampusID = "0" then %>
	<td>Campus</td>
<% end if 
	if numCampusID = "0" OR strBuildingLocationID = "0" then
%>
	<td>Building</td>
<% 	end if
	if numlocationID <> "0" then %>
	<TD>Specific Location</TD>
<% end if %>

<% if numCampusID = "0" OR numlocationID = "0" then %>
		<TD>Location</TD>
<% 	end if 
'   end if
%>

<% 'if strBuildingLocationID <> "0" then 
'	if numlocationID = "0" then %>
<!--	<TD>Location</TD> -->
<%' 	end if
  ' end if %>

	<TD>Haz(Y/N)</TD>
	<TD>Details</TD>
	<TD>ADD?</TD>
	<TD>Work Activity</TD>
	<TD>Assessors Name</TD>
	<TD>Date</TD>

</TR>
    <% do while not rsRA.EOF %>
<TR>
	<TD><%= rsRA("strChemicalName") %></TD>

<% if numCampusID = "0" then %>
	<td><%= rsRA("strCampusName") %></td>
<% end if 
	if numCampusID = "0" OR strBuildingLocationID = "0" then
%>
	<td><%= rsRA("strBuildingName") %></td>
<% 	end if
	if numlocationID <> "0" then %>	
	<TD>
        <%= rsRA("strSpecificLocation") %>
	</TD>
<% end if %>

<% if numCampusID = "0" OR numlocationID = "0" then %>	
	<TD>
        <%= rsRA("strStoreLocation") + ", " + rsRA("strStoreType") + ", " + rsRA("strStoreNotes") %>
	</TD>
<% 	end if 
'   end if
%>

<% 'if strBuildingLocationID <> "0" then 
	'if numlocationID = "0" then %>
<!--	<TD>
        <%'= rsRA("strStoreLocation") %>
	</TD> -->
<% '	end if 
   'end if
%>
			
<TD><%= rsRA("strHazardous")%></TD>
			
	<TD><% numRiskAssessmentID = rsRA("numRiskAssessmentID") 
		if IsNumeric(numRiskAssessmentID) then 
		strStoreManager = server.urlencode(rsRA("strStoreManager"))
		strChemicalName = server.urlencode(rsRA("strChemicalName"))
		%>
				<A HREF="UpdateRiskAssessment.asp?numRiskAssessmentID=<%= numRiskAssessmentID%>&numChemicalID=<%= rsRA("numChemicalContainerID") %>&strStoreManager=<%= strStoreManager %>&strChemicalName=<%= strChemicalName %>&numLocationID=<%= numLocationID %>">View</A>
		<% else %>
				None
		<% End If %>
	</TD>
	<% 
	
	numChemicalContainerID = rsRA("numChemicalContainerID")
	strStoreManager = server.urlencode(rsRA("strStoreManager"))
	strChemicalName = server.urlencode(rsRA("strChemicalName"))
	
	    %>
	<TD><A HREF="AddRiskAssessment.asp?numChemicalID=<%= numChemicalContainerID %>&strStoreManager=<%= strStoreManager %>&strChemicalName=<%= strChemicalName %>&numLocationID=<%= numLocationID %>">ADD</A></TD>
	
	<TD>
            <%= rsRA("strWorkActivity")  %></TD>
	<TD>
            <%= rsRA("strAssessorsName") %></TD>
	<TD>
            <%= rsRA("dtmDateOfAssessment") %></TD>
            
	</TR>
    <% 
rsRA.MoveNext

	loop 

	rsRA.Close
	set rsRA = nothing
	conn.close
	set conn = nothing	
%>
    
</TABLE>
</form>


</BODY>
</HTML>
