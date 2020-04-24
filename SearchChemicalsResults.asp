<%@ Language = VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%
rem OPTION EXPLICIT
rem Response.Buffer = true
Dim rsChemicals, rsLocation
Dim strSQL, strSQL2
Dim conn
Dim numLocationID
Dim strBuildingLocationID
Dim numCampusID
Dim storeTypeID
Dim strChemicalName
Dim strCAS
Dim strLocation
Dim strSortByName
Dim numBarcode

'Function InjectionEncode(str)
'	InjectionEncode=Replace(str,"'","'''")
'End Function

strSortByName = Request.Form("chkSortByName")
strChemicalName = Request.Form("txtChemicalName")
strChemicalName = Replace(strChemicalName,"*","%") 

'numLocationID = Request.Form("hdnRoomLocation")
'strBuildingLocationID = Request.Form("hdnbuildinglocation")
'numCampusID = Request.Form("hdncampuslocation")

numLocationID = cstr(session("numLocationID"))
strBuildingLocationID = cstr(session("numBuildingID"))
numCampusID = cstr(session("numCampusID")) 

'response.write(numLocationId)
'response.write(strBuildingLocationId)
'response.write(numCampusId)

rem dlj remarked out
rem response.write "LocationID: " + numLocationID
rem response.write " - Building: " + strBuildingLocationID


'Put together CAS number
strCAS = Request.Form("txtCAS1") + "-" + Request.Form("txtCAS2") + "-" + Request.Form("txtCAS3")

'******************************************old Query*************************************************************************************
'strSQL = "SELECT tblChemicalContainer.numChemicalContainerID, tblChemicalContainer.strChemicalName, tblLocation.strBuildingLocation, tblLocation.boolLicensedDepot, tblStoreType.strStoreType, "
'strSQL = strSQL + "tblStoreLocation.strStoreLocation, tblLocation.strStoreNotes, tblLocation.numStoreTypeID, tblLocation.numBuildingID, tblLocation.numCampusID, tblLocation.strStoreManager,tblLocation.dtmLastUpdated, tblChemicalContainer.strSpecificLocation, "
'strSQL = strSQL + "tblChemicalContainer.numQuantity, tblChemicalContainer.strContainerSize, tblChemicalContainer.strHazardous, tblChemicalContainer.strGrade, tblChemicalContainer.numLocationID, "
'strSQL = strSQL + "tblChemicalContainer.strUnNumber, tblChemicalContainer.strPG, "
'strSQL = strSQL + "tblCampus.strCampusName, tblBuilding.strBuildingName "
'strSQL = strSQL + "FROM tblChemicalContainer, tblLocation, tblStoreType, tblStoreLocation, tblCampus, tblBuilding "
'strSQL = strSQL + "WHERE tblChemicalContainer.numLocationID = tblLocation.numLocationID AND "
''''strSQL = strSQL + "((tblLocation.numLocationID)=[id]) AND "
'strSQL = strSQL + "tblCampus.numCampusID = tblLocation.numCampusID AND "
'strSQL = strSQL + "tblBuilding.numBuildingID = tblLocation.numBuildingID AND "
'strSQL = strSQL + "tblStoreType.numStoreTypeID = tblLocation.numStoreTypeID AND "
'strSQL = strSQL + "tblLocation.numStoreLocationID = tblStoreLocation.numStoreLocationID AND "
'******************************************old Query*************************************************************************************

'******************************************New Query*************************************************************************************

 strSQL ="SELECT tblChemicalContainer.numChemicalContainerID, tblChemicalContainer.strChemicalName, tblLocation.strBuildingLocation, tblLocation.strLoginId,"_
&" tblChemicalContainer.numSize, tblChemicalContainer.strContainerUnits ,"_
&" tblChemicalContainer.strOtherInfo, tblChemicalContainer.strCas,tblLocation.boolLicensedDepot, tblStoreType.strStoreType, tblStoreLocation.strStoreLocation,"_
&" tblLocation.strStoreNotes, tblLocation.numStoreTypeID, tblLocation.numBuildingID, tblLocation.numCampusID, tblLocation.strStoreManager, tblLocation.dtmLastUpdated,"_
&" tblChemicalContainer.strSpecificLocation, tblChemicalContainer.numQuantity, tblChemicalContainer.strContainerSize, tblChemicalContainer.strHazardous, "_
&" tblChemicalContainer.strGrade, tblChemicalContainer.strDangerousGoodsClass, tblChemicalContainer.numLocationID, tblChemicalContainer.strUnNumber, tblChemicalContainer.numBarcode, tblChemicalContainer.strPG,"_
&" tblCampus.strCampusName, tblBuilding.strBuildingName, tblChemicalContainer.strSSDG,   (select count(*) from tblRiskAssessment where tblRiskAssessment.numChemicalContainerId = tblChemicalContainer.numChemicalContainerId) as numRIskAssessmentId"_
 &" FROM ((((tblChemicalContainer RIGHT JOIN tblLocation ON tblLocation.numLocationId = tblChemicalContainer.numLocationId) INNER JOIN tblCampus ON tblCampus.numCampusId = tblLocation.numCampusId) "_
 &"INNER JOIN tblBuilding ON tblBuilding.numBuildingId = tblLocation.numBuildingId) INNER JOIN tblStoreType ON tblStoreType.numStoreTypeId = tblLocation.numStoreTypeId) "_
 &"INNER JOIN tblStoreLocation ON tblStoreLocation.numStoreLocationId = tblLocation.numStoreLocationId Where "_


'******************************************New Query*************************************************************************************
if numCampusId <> "0" then
	if strBuildingLocationID <> "0" then 
		rem a Builing location has been chosen
		if numLocationID = "0" then
			rem a room has NOT been chosen, ie search all rooms at in a building
			strSQL = strSQL + "(tblLocation.numBuildingID = " + strBuildingLocationID + ") AND "
		else
			rem a room has been chosen
		strSQL = strSQL + "(tblChemicalContainer.numLocationID = " + numLocationID + ") AND "
		end if
	else
		strSQL = strSQL + "(tblLocation.numCampusID = " + numCampusID + ") AND "	
	end if 
end if

strSQL = strSQL + "(tblChemicalContainer.strChemicalName LIKE '" + InjectionEncode(strChemicalName) + "') AND "
'strSQL = strSQL + "(tblChemicalContainer.strChemicalName LIKE '" + Replace(InjectionEncode(strChemicalName),"*","") + "%') AND "

if (strCAS = "--") then 
	strSQL = strSQL + "(tblChemicalContainer.strCAS LIKE '%' OR tblChemicalContainer.strCAS IS NULL)"
else
'	strSQL = strSQL + "(tblChemicalContainer.strCAS LIKE '" + strCAS + "')"
	strSQL = strSQL + "(tblChemicalContainer.strCAS LIKE '" + strCAS + "')"
end if



strSQL = strSQL + " ORDER BY "

'rem  IF A LOCATION IS SELECTED AND SORT BY NAME NOT SELECTED THEN ORDER BY LOCATION

if strSortByName <> "on" then
	if numCampusID = "0" then
 		strSQL = strSQL + "tblCampus.strCampusName, "
	end if
	
	if strBuildingLocationID = "0" then
 		strSQL = strSQL + "tblBuilding.strBuildingName, "
  	end if

 	if numlocationID = "0" then
 		strSQL = strSQL + "tblStoreLocation.strStoreLocation, tblStoreType.strStoreType, tblLocation.strStoreNotes, "
  	end if

	if numlocationID <> "0" then
 		strSQL = strSQL + "tblChemicalContainer.strSpecificLocation, "
  	end if
  
end if
' if ((strBuildingLocationID <> "0") and strSortByName <> "on") then
' 	strSQL = strSQL + "tblChemicalContainer.strSpecificLocation, "
'  end if
  
' if ((numlocationID = "0") and strSortByName <> "on") then
' 	strSQL = strSQL + "tblStoreLocation.strStoreLocation, "
'  end if

strSQL = strSQL + "tblChemicalContainer.strChemicalName"

'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

set rsChemicals = Server.CreateObject("ADODB.Recordset")
'response.write(strSQL)
rsChemicals.Open strSQL, conn


'set rsChemcials = nothing
'rem cant get this bit to work 
'rem an * is not recognised as a value
'rem if (strChemicalName = "*" and numLocationID = "ALL") then
'rem	Response.Write "<BR><DIV align='center'><FONT color='red' face='Arial'>Cannot search ALL chemicals at ALL locations, please make your search more specific.</FONT></DIV>"
'rem	Response.End
'rem end if

if rsChemicals.EOF then
	Response.Write "<BR><DIV align='center'><FONT color='red' face='Arial'>There are no results for that search, please make your search less specific (try using *, for wildcard search).</FONT></DIV>"
	Response.End
end if

strSQL2 = "SELECT tblBuilding.strBuildingName, tblCampus.strCampusName FROM tblCampus, tblBuilding "
strSQL2 = strSQL2 + "WHERE tblCampus.numCampusID = " + cstr(rsChemicals("numCampusID"))  + " AND "
strSQL2 = strSQL2 + "tblBuilding.numBuildingID = " + cstr(rsChemicals("numBuildingID"))  + " AND tblCampus.numCampusID = tblBuilding.numCampusID"

set rsLocation = Server.CreateObject("ADODB.Recordset")
rsLocation.Open strSQL2, conn
'set rsLocation = nothing
'conn.close
'set conn = nothing

%>
<HTML>
<HEAD>
	<TITLE>Chemical Search Results</TITLE>
</HEAD>

<BODY>
<% 
	'strLocation = cstr(rsChemicals("numCampusID"))+ ", " + cstr(rsChemicals("numBuildingID"))
'	response.write(strSortByName) + " 2"
%>
<% if numLocationID <> "0" then %>
<table border="0" width="100%" id="table1">
	<tr>
		<td width="495" valign="top">
		<p align="left">Supervisor: <%= rsChemicals("strStoreManager") %> <br />
Location: <%= rsLocation("strCampusName") + ", " + rsLocation("strBuildingName") + ", " +rsChemicals("strStoreLocation") %><br />
Last Updated: <%= rsChemicals("dtmLastUpdated") %> <br />
Store Type: <%= rsChemicals("strStoreType") + ", " + rsChemicals("strStoreNotes") %> <br />
Location ID: <%= rsChemicals("numLocationID") %> </p>
<% end if %>


		<p>&nbsp;</td>
        <%

if (numLocationID<>"0" and strBuildingLocationID<>"0" and numCampusID<>"0") then
set rsStoreT=server.CreateObject("ADODB.RecordSet")
str="select distinct(numStoreTypeID) AS storeTypeID from qryDangerousGood where numLocationID=" + numLocationID + " And numBuildingID=" + strBuildingLocationID + " and numCampusID=" + numCampusID


%>
	
		<td align="left" valign="top">Dangerous goods summary for this location: <br/>
        <%
'============================Code by Jeff, Show danger class item from summary =============================================
	set rsSummary=server.CreateObject("ADODB.RecordSet")
	strSummary="select Distinct(strDangerousGoodClass), PG, sum(Total) as TotalAmount from qryDangerousGood where numLocationID=" + numLocationID + " And numBuildingID=" + strBuildingLocationID + " and numCampusID=" + numCampusID + " and (PG <> '' and ucase(PG)<>ucase('none')) group by strDangerousGoodClass, PG order by strDangerousGoodClass"
'response.write(strSummary)
	rsSummary.open strSummary, conn



			%>
        <TABLE style="WIDTH: 500px" width=536 border=0>
     <!--TR>
    <TD width="13%" height="17" align=middle bgColor=#ffff00>&nbsp;</TD>
    <TD align=middle width="9%" bgColor=#ffff00>&nbsp;</TD>
    <TD align=middle width="12%" bgColor=#ffff00>&nbsp;</TD>
    <TD align=middle bgColor=#ffff00 colspan="4">Quantities ( L / Kg )</TD>
    </TR-->

     <TR>
    <TD width="13%" height="12" align=middle bgColor=#FFFFFF>&nbsp;</TD>
    <TD align=middle width="21%" bgColor=#e9e9e9 colspan="4">Quantities ( L / Kg )</TD>
    </TR>


	<TR>
    <TD width="13%" height="17" align=middle bgColor=#FFFF00>DG Class</TD>
    <TD align=middle width="9%" bgColor=#ffff00>PG I</TD>
    <TD align=middle width="12%" bgColor=#ffff00>PG II</TD>
    <TD align=middle width="16%" bgColor=#ffff00>PG III</TD>
    <TD align=middle width="50%" bgColor=#ffff00>Total</TD>
    </TR>
    <%
    dim dangerClass(30)
	dim PGI(30)
	dim PGII(30)
	dim PGIII(30)
	dim first
	dim i
	dim j
	i=0
	first = true

	
	
	
    
	%>
  <% do until rsSummary.eof 
  if first=true and isnumeric(rsSummary.fields.item("strDangerousGoodClass"))  then
          dangerClass(i)=cint(rsSummary.fields.item("strDangerousGoodClass"))
           if rsSummary.fields.item("PG")="I" then
	          PGI(i)=PGI(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i)=PGII(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i)=PGIII(i)+rsSummary.fields.item("TotalAmount")
	       end if
		
		first=false
		 'response.write(rsSummary.fields.item("strDangerousGoodClass") + " is 0" + "<br/>")
	elseif isnumeric(rsSummary.fields.item("strDangerousGoodClass")) then
	    if dangerClass(i)=cint(rsSummary.fields.item("strDangerousGoodClass")) then
	       dangerClass(i)=cint(rsSummary.fields.item("strDangerousGoodClass"))
           if rsSummary.fields.item("PG")="I" then
	          PGI(i)=PGI(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i)=PGII(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i)=PGIII(i)+rsSummary.fields.item("TotalAmount")
	       end if
		  ' response.write(rsSummary.fields.item("strDangerousGoodClass") + " is 1" + "<br/>")
	    else
		     dangerClass(i+1)=cint(rsSummary.fields.item("strDangerousGoodClass"))
           if rsSummary.fields.item("PG")="I" then
	          PGI(i+1)=rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i+1)=rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i+1)=rsSummary.fields.item("TotalAmount")
	       end if
			i=i+1
					  ' response.write(rsSummary.fields.item("strDangerousGoodClass") + " is 2" + "<br/>")
		end if 

	 elseif ucase(rsSummary.fields.item("strDangerousGoodClass"))=ucase("none")  then
	      if first=true then
		      dangerClass(i)="None"
             if rsSummary.fields.item("PG")="I" then
	          PGI(i)=PGI(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i)=PGII(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i)=PGIII(i)+rsSummary.fields.item("TotalAmount")
	       end if
		      first=false
		   elseif ucase(dangerClass(i))=ucase("none") then
	 	      dangerClass(i)="None"
            if rsSummary.fields.item("PG")="I" then
	          PGI(i)=PGI(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i)=PGII(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i)=PGIII(i)+rsSummary.fields.item("TotalAmount")
	       end if
		   else
		      dangerClass(i+1)="None"
              if rsSummary.fields.item("PG")="I" then
	             PGI(i+1)=rsSummary.fields.item("TotalAmount")
	          elseif rsSummary.fields.item("PG")="II" then
	             PGII(i+1)=rsSummary.fields.item("TotalAmount")
	          elseif rsSummary.fields.item("PG")="III" then
	             PGIII(i+1)=rsSummary.fields.item("TotalAmount")
	          end if
			  i=i+1
			end if
		   		  ' response.write(rsSummary.fields.item("strDangerousGoodClass") + " is 3" + "<br/>")
	 elseif rsSummary.fields.item("strDangerousGoodClass")="" then
	        if first=true then
	 	       dangerClass(i)="Empty"
              if rsSummary.fields.item("PG")="I" then
	          PGI(i)=PGI(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i)=PGII(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i)=PGIII(i)+rsSummary.fields.item("TotalAmount")
	       end if
			   first=false
			elseif dangerClass(i)="" or ucase(dangerClass(i))=ucase("empty") then
	 	       dangerClass(i)="Empty"
             if rsSummary.fields.item("PG")="I" then
	          PGI(i)=PGI(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="II" then
	          PGII(i)=PGII(i)+rsSummary.fields.item("TotalAmount")
	       elseif rsSummary.fields.item("PG")="III" then
	          PGIII(i)=PGIII(i)+rsSummary.fields.item("TotalAmount")
	       end if
			 else
			   dangerClass(i+1)="Empty"
               if rsSummary.fields.item("PG")="I" then
	              PGI(i+1)=rsSummary.fields.item("TotalAmount")
	           elseif rsSummary.fields.item("PG")="II" then
	              PGII(i+1)=rsSummary.fields.item("TotalAmount")
	           elseif rsSummary.fields.item("PG")="III" then
	              PGIII(i+1)=rsSummary.fields.item("TotalAmount")
	           end if
			   i=i+1
			 end if
			 
	      		  ' response.write(rsSummary.fields.item("strDangerousGoodClass") + " is 4" + "<br/>")
	end if
	'response.write("i is " & i & "<br/>")
	rsSummary.movenext
	loop
	

	
	
	%>
  <%for j=0 to i %>

   <tr>

      <TD width="20%" height="17" align=middle bgColor=#e9e9e9>Class <%=dangerClass(j)%></TD>
    <TD align=middle width="9%" bgColor=#e9e9e9>
	<% if PGI(j)="" or isnull(PGI(j)) then
	 PGI(j)=0 
	    response.write(PGI(j)) 
	 else 
	    response.write(PGI(j)) 
	 end if%>
     </TD>          
      <TD align=middle width="12%" bgColor=#e9e9e9>
	  <% if PGII(j)="" or isnull(PGII(j)) then 
	  PGII(j)=0 
	  response.write(PGII(j))
	  else 
	  response.write(PGII(j)) 
	  end if%>
      </TD>
      
    <TD align=middle width="16%" bgColor=#e9e9e9>
	<% if PGIII(j)="" or isnull(PGIII(j))then  
	      PGIII(j)=0
		  response.write(PGIII(j)) 
	   else 
	      response.write(PGIII(j)) 
	   end if%></TD>
    <TD align=middle width="50%" bgColor=#e9e9e9><%=PGI(j) + PGII(j) + PGIII(j)%></TD>

    </tr>
<%next %>
      </TABLE>

</td>

<%end if%>




	</tr>
</table>

<%'====================================End Code=============================================================================%>


<TABLE WIDTH="100%" BORDER=1 ALIGN="center" VALIGN = "TOP">

<TR ALIGN="left" VALIGN="top" BGCOLOR="yellow">
<!-- DLJ 4Aug15 added chemical ID number -->
<!-- DLJ 5March2019 added Barcode ID-->
	<TD>ID</TD>
	<td>Barcode ID</td>
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

<% ' if numLocationID <> "0" then
	'if rsChemicals("numStoreTypeID") = "1" then %>
<% '	end if 
   'end if
   
   Dim strLoginId
   strLoginId = Session("strLoginId")
%>
<!-- DLJ 23Jan8 moved UN Number and pg out of above if statement so they are shown in ALL search results -->
<!-- DLJ 23Jan8 changed order of fields and added DG Class-->

	<TD>Quantity</TD>	
	<!--<TD>Size</TD>-->
	<td>Size</td>
	<td>Units</td>
	<TD>CAS #</TD>
	<TD>Grade</TD>

	<TD>Other Info</TD>
	<TD>Hazardous?</TD>
	<!--TD>R A Done</TD-->
    <TD>DG Class</TD>
    <TD>PG (I, II, III)</TD>
	<!--<td>Strlogin</td>-->
	
	<!--
		<%'AA Feb 2014 - hide SSDG from general users 
		if strLoginId <> "science" then %>
		<TD>SS</TD>
		<% end if %>
	-->

</TR>
    <% 
		
       do while not rsChemicals.EOF 
	   
	   '*** AA Jan 2014 ***'
	   'Admin will display all results regardless of strSSD set or not
		if(strLoginId <> "admin" and strLoginId <> rsChemicals("strLoginId") and rsChemicals("strSSDG")= "Yes") then
		'Do nothing - don't show the lines that aren't our SSDG
		else
	   %>
<TR>
	
	<TD><font size="-1"><%= rsChemicals("numChemicalContainerID") %></font></TD>
	<TD><font size="-1"><%= rsChemicals("numBarcode") %></font></TD>


	<TD>
	        <A HREF="ChemicalDetails.asp?numChemicalID=<%= rsChemicals("numChemicalContainerID") %>&numLocationID=<%=rsChemicals("numLocationID")%>">
            <%= rsChemicals("strChemicalName") %></A>
    </TD>
<% if numCampusID = "0" then %>
	<td><%= rsChemicals("strCampusName") %></td>
<% end if 
	if numCampusID = "0" OR strBuildingLocationID = "0" then
%>
	<td><%= rsChemicals("strBuildingName") %></td>
<% 	end if
	if numlocationID <> "0" then %>	
	<TD>
        <%= rsChemicals("strSpecificLocation") %>
	</TD>
<% end if %>

<% if numCampusID = "0" OR numlocationID = "0" then %>	
	<TD>
        <%= rsChemicals("strStoreLocation") + ", " + rsChemicals("strStoreType") + ", " + rsChemicals("strStoreNotes") %>
	</TD>
<% 	end if 
'   end if
%>

<% ' if numlocationID <> "0" then 
'if rsChemicals("numStoreTypeID") = "1" then %>	
<% 
'end if
'end if %>

<!-- DLJ 23Jan8 moved UN Number and pg out of above if statement so they are shown in ALL search results -->
<!-- DLJ 23Jan8 changed order of fields and added DGClass -->
	<TD><%= rsChemicals("numQuantity") %></TD>
	<!--<TD><%= rsChemicals("strContainerSize") %></TD>  -->
	<TD><%= rsChemicals("numSize") %></TD>
	<TD><%= rsChemicals("strContainerUnits") %></TD>
	
	<TD><%= rsChemicals("strCas") %></TD> 
	<TD><%= rsChemicals("strGrade") %></TD>
	<TD><%= rsChemicals("strOtherInfo") %></TD>
    <TD><%= rsChemicals("strHazardous") %></TD>        
    <!--TD><% dim boolRADone
             boolRADone = rsChemicals("numRiskAssessmentId") 
             'if len(boolRADone)>0  then
			 if boolRADone>0  then
             %>Yes<%
             else
              %>No<%
             end if
             %></TD-->
    <TD><%= rsChemicals("strDangerousGoodsClass") %></TD>
	<TD><%= rsChemicals("strPG") %></TD>
	<!--<td><%= rsChemicals("strLoginId") %></td>-->

	
		<!--
			<% 'AA Feb 2014 - hide SSDG from general users
			if strLoginId <> "science" then %>
			<TD><%= rsChemicals("strSSDG") %></TD>
	<% end if %>
	-->

</TR>
    <% 
	end if
rsChemicals.MoveNext
	
	loop 

'AA Set up for CSV
	dim exp 
	exp = Split(strSQL, "FROM")
	
	dim newFields
	newFields = "SELECT tblChemicalContainer.numBarcode, tblChemicalContainer.strChemicalName, tblChemicalContainer.strSpecificLocation,"_
	&" tblChemicalContainer.numSize, tblChemicalContainer.strContainerUnits, tblChemicalContainer.numQuantity,"_
	&" tblChemicalContainer.strCas, tblChemicalContainer.strOtherInfo ,tblChemicalContainer.strContainerOwner,  "_
	
&" tblChemicalContainer.strGrade, tblChemicalContainer.strManufacturer, tblChemicalContainer.strproductNumber, "_
	& " tblChemicalContainer.strState, tblChemicalContainer.strExpiry "

	
	dim newSQL
	newSQL = newFields&" FROM "&exp(1)
	

' DLJ March 2014 only supervisor and admin logins can download CSV 

		set rsLocation = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT tblLocation.strLoginID, tblLocation.numLocationID "
		strSQL = strSQL +  "FROM tblLocation "
		strSQL = strSQL +  "WHERE tblLocation.numLocationID = " + numLocationID
		rsLocation.Open strSQL, conn, 3, 3

		'response.write(strLoginID)
		'response.write strSQL
		' first check it is a record set corresponding to a single location
		if Not rsLocation.EOF Then
		
	  		if ((strLoginID = rsLocation("strLoginID")) Or (strLoginID = "admin")) Then
		'	response.write(rsLocation("strLoginID"))

%>
    <form action="CSVCreator.asp" method=post enctype="application/x-www-form-urlencoded">
	<input type ="hidden" value="rsChemicals" name="data"/>
	<input type ="hidden" value="<%=newSQL %>" name="sql"/>
	<input type="submit" value="Export as CSV"/>
	</form>
<%

			end if
		End if	


%>



	
	<%
		rsChemicals.Close
	rsLocation.close
	set rsLocation = nothing
	set rsChemicals = nothing
	conn.close
	set conn = nothing	
	%>
</TABLE>
</BODY>
</HTML>