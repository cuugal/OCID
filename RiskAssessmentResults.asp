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
  
  if isNull(numLocationId) or numLocationId = ""  then numLocationId = 0 end if
  if isNull(strBuildingLocationID) or strBuildingLocationID = ""  then strBuildingLocationID = 0 end if
  if isNull(numCampusID) or numCampusID = ""  then numCampusID = 0 end if
  
strSortByName = Request.form("chkSortByName")
strChemicalName = Replace(strChemicalName,"*","%") 

strSQL = "SELECT * FROM qryChemicalRiskAssessment WHERE "
if numCampusId <> "0" Then
'if a building has been chosen - NOTE building is actually stored in numCampusId
	if strBuildingLocationID <> "0" then 
'		rem a floor has been chosen
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
		'DLJ commented out 21March2014 and added response to prevent error  strSQL = strSQL + "(qryChemicalRiskAssessment.numCampusID = " + numCampusID + ") AND "	
	Response.Write "<BR><DIV align='center'><FONT color='red' face='Arial'>There are too many result for that search, please select a Floor.</FONT></DIV>"
	Response.End
'		strSQL = strSQL + "IsNumeric(qryChemicalRiskAssessment.numRiskAssessmentID) AND "
	end if 
end If

If numCampusId = "0" Then
		'DLJ 21March2014  added response to prevent error	
	Response.Write "<BR><DIV align='center'><FONT color='red' face='Arial'>There are too many result for that search, please select a Building.</FONT></DIV>"
	Response.End
End if

strSQL = strSQL + "(qryChemicalRiskAssessment.strChemicalName LIKE '" + InjectionEncode(strChemicalName) + "') ORDER BY "



if strSortByName <> "true" Then
'DLJ edit 21March2014 to sort by assessors name - better if put nulls at end and alphabetical
	strSQL = strSQL + "qryChemicalRiskAssessment.strAssessorsName DESC, "
	'if numCampusID = "0" then
 	'	strSQL = strSQL + "qryChemicalRiskAssessment.strCampusName, "
	'end if
	
	'if strBuildingLocationID = "0" then
 	'	strSQL = strSQL + "qryChemicalRiskAssessment.strBuildingName, "
  	'end if

 	'if numlocationID = "0" then
 	'	strSQL = strSQL + "qryChemicalRiskAssessment.strStoreLocation, qryChemicalRiskAssessment.strStoreType, qryChemicalRiskAssessment.strStoreNotes, "
  	'end if

	'if numlocationID <> "0" then
 	'	strSQL = strSQL + "qryChemicalRiskAssessment.strSpecificLocation, "
  	'end if
  
end If



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
'response.write strSQL

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
	if  numLocationID = "0" then
			rem a room has NOT been chosen, ie search all rooms at in a building
		strSQL2 = strSQL2 + " AND (tblLocation.numBuildingID = " + cstr(rsRA("numBuildingID")) + ")"
	else
			rem a room has been chosen
		''strSQL2 = strSQL2 + " AND (tblLocation.numLocationID = " + cstr(rsRA("numLocationID")) + ")"
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
dim colorText
%>
<HTML>
<HEAD>
	<TITLE>Risk Assessment Results</TITLE>
</HEAD>

<BODY>
<FORM>
<% if numLocationID <> "0" then %>
<p align="left">Supervisor: <%= rsRA("strStoreManager") %> <Br>
Location: <%= rsBuildingCampus("strCampusName") + ", " + rsBuildingCampus("strBuildingName") + ", " + rsLocation("tblStoreLocation.strStoreLocation") %><BR>
Last Updated: <%= rsRA("dtmLastUpdated") %><br/>
Store Type: <%=rsLocation("strStoreType") + ", " + rsLocation("strStoreNotes")%>
</p>
<% end if %>

<TABLE WIDTH="100%" BORDER=1 ALIGN="center" VALIGN = "TOP">
<TR ALIGN="left" VALIGN="top" BGCOLOR="yellow">
	<TD>Name</TD>
	<TD>Assessors Name</TD>
	<TD>Assessment Date</TD>
	<TD>Work Activity</TD>
	 <TD>Risks Controlled(Y/N)</TD>
	<TD>Details</TD>
	<TD>ADD?</TD>
	
   
	
	
	<td>Location of Use</td>

</TR>
    <% do while not rsRA.EOF 
	colorText = "" 
	%>
<TR>
	<TD><%= rsRA("strChemicalName") %></TD>
	<TD> <%= rsRA("strAssessorsName") %></TD>
	<% 
	
	if(rsRA("dtmDateOfAssessment") <> "" and not isNull(rsRA("dtmDateOfAssessment"))) then
		dim assessmentDate
		assessmentDate = CDate(rsRA("dtmDateOfAssessment"))
		if DateAdd("yyyy",2, CDate(assessmentDate)) < DateValue(Date()) then
			colorText = "bgcolor='Red'"
		end if
	end if %>
	<TD <%=colorText%> >
		<% =rsRA("dtmDateOfAssessment") 

		%>
	</TD>
	<TD> <%= rsRA("strWorkActivity")  %></TD>
	<TD>
		<% if (ucase(rsRA("strRiskControlled"))=ucase("TRUE")) then%>
		Yes
		<%elseif (ucase(rsRA("strRiskControlled"))=ucase("FALSE")) then%>
		No
		<%else%>
		<!-- used to say NULL-->
		<%end if%>								
    </TD>
			
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
	
	
    
	
	
	<% if numlocationID <> "0" then %>	
	<!-- If a room is selected -->
	<TD>
        <%= rsRA("strLocationOfUse") %>
	</TD>
	<% end if %>

	<% if numCampusID = "0" OR numBuildingID = "0" Or numlocationID = "0" then %>	
	<!-- if no building or no room is selected -->
	<TD>
        <%= rsRA("strLocationOfUse") %>
       <!-- DLJ commented out on 21March2014 <%= rsRA("strStoreLocation") + ", " + rsRA("strStoreType") + ", " + rsRA("strStoreNotes") %> -->
	</TD>
	
<% 	end if %>
            
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
