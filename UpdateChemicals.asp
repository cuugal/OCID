<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%
           Server.ScriptTimeout = 180 
          
Sub CleanUp() 

	set rsChemicals = nothing
	conn.close
	set conn = nothing

End Sub

Dim numLocationID
Dim numChemicalID
Dim rsChemicals, rsBuildingCampus
Dim strSQL, strSQL2
Dim conn 
Dim numRecords
Dim numRecordCounter
Dim strSortEditByName
Dim strBuildingLocationID
Dim numCampusId

'numLocationID = Request.Form("hdnNumLocationID")

numLocationID = cstr(session("numLocationID"))
strBuildingLocationID = cstr(session("numBuildingID"))
numCampusID = cstr(session("numCampusID")) 

if numLocationID = "0" then
	Response.Write "You must Choose a Location"
	Response.End
end if

'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

'numLocationID = Request.Form("hdnNumLocationID")
'numCampusID = Request.Form("hdncampuslocation")
strSortEditByName = Request.Form("chkSortEditByName")

'strBuildingLocationID = Request.Form("hdnbuildinglocation")



set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsChemicals = Server.CreateObject("ADODB.Recordset")

If Request.Form("UPDATE") = "SAVE Changes to Chemical Inventory" then 
	
	Dim strChemicalName
	Dim strSpecificLocation
	Dim numQuantity
	Dim strContainerOwner
	Dim strContainerSize
	Dim strCAS
	Dim strGrade
	Dim strHazardous
	Dim strDangerousGoodsClass
	Dim strOtherInfo
	Dim temp
	Dim strUnNumber
	Dim strPG
    dim inAction 
    dim splitContainerS
    dim splitContainerM
    dim spacePosition
    dim isCorrectSize
	dim isAbleToDelete
	isAbleToDelete=true
    isCorrectSize=true
    inAction = request.form("actionValue")
	numRecords = Request.Form("hdnRecordCount")
	For numRecordCounter = 1 to numRecords
		numChemicalID = Request.Form("hdnNumChemicalID" + cstr(numRecordCounter))
		
	set rsChemicals = Server.CreateObject("ADODB.Recordset")		
		
		If (Request.Form("chkDelete" + cstr(numRecordCounter))= "ON") then 
		    set rsSearchRecord=server.createobject("ADODB.Recordset")
			'sqlSearch="select count(numChemicalContainerID) as numChemicalContainerID from tblRiskAssessment where numChemicalContainerID= " & numChemicalID
			'rsSearchRecord.open sqlSearch, conn
			'if (rsSearchRecord.fields.item("numChemicalContainerID")>0) then
				'isAbleToDelete=false
			'   exit for
			'else
			   strSQL = "DELETE FROM tblChemicalContainer "
			   strSQL = strSQL +  "WHERE (numChemicalContainerID = " + numChemicalID + ")"
			   rsChemicals.Open strSQL, conn, 3, 3
            'end if			   
		Else
			strChemicalName = Request.Form("txtChemicalName" + cstr(numRecordCounter))		
			strSpecificLocation = Request.Form("txtSpecificLocation" + cstr(numRecordCounter))
			numQuantity = Request.Form("txtQuantity" + cstr(numRecordCounter))
			strContainerSize = Request.Form("txtContainerSize" + cstr(numRecordCounter))
            spacePosition=instr(1,strContainerSize," ")
            if spacePosition=0 then
            isCorrectSize=false
			exit for
            else
               splitContainerS=mid(strContainerSize,1,instr(1,strContainerSize," ")-1)
               splitContainerM=mid(strContainerSize,instr(strContainerSize," ")+1)
               if isnumeric(splitContainerS)=false then
                  isCorrectSize=false
                  exit For
                  'DLJ commented this out 24Oct2008 - it was preventing user from using non-standard units such
				  'elseif ucase(splitContainerM)<>ucase("mg") and ucase(splitContainerM)<>ucase("g") and ucase(splitContainerM)<>ucase("ml") and ucase(splitContainerM)<>ucase("l") and ucase(splitContainerM)<>ucase("kg") then
					 ' isCorrectSize=false
					'exit for  
			   end if
            end if
			strGrade = Request.Form("txtGrade" + cstr(numRecordCounter))
			strContainerOwner = Request.Form("txtOwner" + cstr(numRecordCounter))
			strUnNumber = Request.Form("txtUnNumber" + cstr(numRecordCounter))		
			strPG = Request.Form("txtPG" + cstr(numRecordCounter))		
			strCAS = Request.Form("txtCAS" + cstr(numRecordCounter)) 
			strHazardous = Request.Form("txtHazardous" + cstr(numRecordCounter))
			strSSDG = Request.Form("txtSSDG" + cstr(numRecordCounter))
			
			strSubsDG = Request.Form("txtsubsDG" + cstr(numRecordCounter))
			strHazchem = Request.Form("txtHazchem" + cstr(numRecordCounter))
			strPoisons = Request.Form("txtPoisons" + cstr(numRecordCounter))
			
			'Put together CAS number
			'strCASa = Request.Form("txtCASa" + cstr(numRecordCounter)) 
			'strCASb = Request.Form("txtCASb" + cstr(numRecordCounter)) 
			'strCASc = Request.Form("txtCASc" + cstr(numRecordCounter)) 
			'strCAS = strCASa + "-" + strCASb + "-" + strCASc
			
			if strHazardous = "on" then
				strHazardous = "Yes"
			else
				strHazardous = "No"
			end if
			
			if strSSDG = "on" then
				strSSDG = "Yes"
			else
				strSSDG = "No"
			end if
			
			strDangerousGoodsClass = Request.Form("txtDangerousGoodsClass" + cstr(numRecordCounter))
			
			strOtherInfo = Request.Form("txtOtherInfo" + cstr(numRecordCounter))   

			strSQL = "UPDATE tblChemicalContainer SET "
			
			if strChemicalName = "" then
				strSQL = strSQL +  "strChemicalName = NULL, "
			else
				strSQL = strSQL +  "strChemicalName = '" + InjectionEncode(strChemicalName) + "', "
			end if

			if strSpecificLocation = "" then
				strSQL = strSQL +  "strSpecificLocation = NULL, "
			else
				strSQL = strSQL +  "strSpecificLocation = '" + InjectionEncode(strSpecificLocation) + "', "
			end if
			
			if numQuantity = "" then
				strSQL = strSQL +  "numQuantity = NULL, "
			else
				strSQL = strSQL +  "numQuantity = '" + numQuantity + "', "
			end if
			
			if strContainerOwner = "" then
				strSQL = strSQL +  "strContainerOwner = NULL, "
			else
				strSQL = strSQL +  "strContainerOwner = '" + InjectionEncode(strContainerOwner) + "', "
			end if

	        'if inAction = "F" then
  					if strContainerSize = "" then
						strSQL = strSQL + "strContainerSize = NULL,"
					else
						strSQL = strSQL + "strContainerSize = '" + InjectionEncode(strContainerSize) + "', "
					end if

	        'elseif inAction = "T" then
					'if strContainerSize = "" then
					'	strSQL = strSQL + "strContainerSize = NULL,"
					'else
					'	strSQL = strSQL + "strContainerSize = '" + InjectionEncode(strContainerSize) + "', "
					'end if
			'end if
			if strCAS = "" then
				strSQL = strSQL + "strCAS = NULL,"
			else
				strSQL = strSQL + "strCAS = '" + strCAS + "', "
			end if
			
			if strGrade = "" then
				strSQL = strSQL + "strGrade = NULL,"
			else
				strSQL = strSQL + "strGrade = '" + InjectionEncode(strGrade) + "', "
			end if
			
			strSQL = strSQL + "strHazardous = '" + strHazardous + "', "
			strSQL = strSQL + "strSSDG = '" + strSSDG + "', "
			strSQL = strSQL + "strDangerousGoodsClass = '" + InjectionEncode(strDangerousGoodsClass) + "', "
			strSQL = strSQL + "strSubsDG = '" + InjectionEncode(strSubsDG) + "', "
			strSQL = strSQL + "strHazchem = '" + InjectionEncode(strHazchem) + "', "
			strSQL = strSQL + "strPoisons = '" + InjectionEncode(strPoisons) + "', "
			strSQL = strSQL + "strUnNumber = '" + strUnNumber + "', "
			strSQL = strSQL + "strPG = '" + strPG + "', "
			strSQL = strSQL + "strOtherInfo = '" + InjectionEncode(strOtherInfo) + "' "
			strSQL = strSQL + "WHERE (numChemicalContainerID = " + numChemicalID + ")"

				'Response.Write ("Function currently not working")
				'Response.Write strSQL
				'Response.End



			rsChemicals.Open strSQL, conn, 3, 3
		End If
		set rsChemicals = nothing
		
	Next
	
	Dim dtmLastUpdated
	dtmLastUpdated = DanDate(Date, "%d/%m/%Y" )
	dtmLastUpdated = cstr(dtmLastUpdated)
	if isAbleToDelete=false then
		   Response.write("This Record cannot be deleted. Please delete the record in Risk Assessment in order to delete this record! Error occur on record line: " & numRecordCounter & "<br/>" & "<a align='center' href='javascript:history.go(-1)'>Go Back</a>")
    elseif isCorrectSize=false then
    Response.write("Please ensure there is a space between the quantities and units (e.g 100 kg). Also ensure there is no empty space before the quantity." & "<br/> Check the record in line " & numRecordCounter & ". " & " <a align='center' href='javascript:history.go(-1)'>Go Back</a>")
	'Response.write("<br/>" & "Position: " & spacePosition & "<br/>" & "splitS: " & splitContainerS & "<br/>" & "splitM: " & splitContainerM)
	else
	
	set rsLocation = Server.CreateObject("ADODB.Recordset")
	strSQL = "UPDATE tblLocation "	
	strSQL = strSQL + "SET dtmLastUpdated = '" + dtmLastUpdated + "' "
	strSQL = strSQL + "WHERE (numLocationID = " + numLocationID + ")"
	

	rsLocation.Open strSQL, conn, 2, 3
	CleanUp() 

	Response.Write ("The Chemicals have been Updated")
    end if
	Response.End
'end of update part of code


else
	
	Dim strLoginID
	strLoginID = lcase(session("strLoginID"))
	if strLoginID <> "admin" then
		set rsLocation = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT tbllocation.strLoginID, tbllocation.numLocationID "
		strSQL = strSQL +  "FROM tblLocation "
		strSQL = strSQL +  "WHERE tblLocation.numLocationID = " + numLocationID
		rsLocation.Open strSQL, conn, 3, 3
		if (rsLocation("strLoginID") <> strLoginID) then
			Response.Write "You do not have permission to update chemicals at this location, please contact the Administrator if you should."
			Response.End
		end if
	end if

strCAS = Request.Form("txtCAS1") + "-" + Request.Form("txtCAS2") + "-" + Request.Form("txtCAS3")

strSQL = "SELECT tblChemicalContainer.numChemicalContainerID, tblChemicalContainer.strCAS, tblChemicalContainer.strContainerOwner, tblChemicalContainer.strChemicalName, tblLocation.strBuildingLocation, tblLocation.boolLicensedDepot, tblStoreType.strStoreType,"
strSQL = strSQL + "tblStoreLocation.strStoreLocation, tblLocation.strStoreNotes, tblLocation.numStoreTypeID, tblLocation.numBuildingID, tblLocation.numCampusID, tblLocation.strStoreManager,tblLocation.dtmLastUpdated, tblChemicalContainer.strSpecificLocation, "
strSQL = strSQL + "tblChemicalContainer.numQuantity, tblChemicalContainer.strDangerousGoodsClass, tblChemicalContainer.strContainerSize, tblChemicalContainer.strHazardous, tblChemicalContainer.strGrade, tblChemicalContainer.numLocationID, "
strSQL = strSQL + "tblChemicalContainer.strUnNumber, tblChemicalContainer.strPG, tblChemicalContainer.strOtherInfo,tblChemicalContainer.strSSDG,tblChemicalContainer.strsubsDG,tblChemicalContainer.strHazchem,tblChemicalContainer.strPoisons "
strSQL = strSQL + "FROM tblChemicalContainer, tblLocation, tblStoreType, tblStoreLocation "
strSQL = strSQL + "WHERE tblChemicalContainer.numLocationID = tblLocation.numLocationID AND "
'strSQL = strSQL + "((tblLocation.numLocationID)=[id]) AND "
strSQL = strSQL + "tblStoreType.numStoreTypeID = tblLocation.numStoreTypeID AND "
strSQL = strSQL + "tblLocation.numStoreLocationID = tblStoreLocation.numStoreLocationID "

'	strSQL = "SELECT * "
'	strSQL = strSQL +  "FROM tblLocation, tblChemicalContainer, tblStoreType "
'	strSQL = strSQL +  "WHERE tblLocation.numLocationID = tblChemicalContainer.numLocationID AND "
'	strSQL = strSQL +  "(tblChemicalContainer.numLocationID = " + numLocationID + ") "
'	strSQL = strSQL + "tblLocation.numStoreTypeID = tblStoreType.numStoreTypeID "

if numCampusID <> "0" then
	if strBuildingLocationID <> "0" then 
			rem a Builing location has been chosen
		if numLocationID = "0" then
			rem a room has NOT been chosen, ie search all rooms at in a building
			strSQL = strSQL + "AND (tblLocation.numBuildingID = " + strBuildingLocationID + ") "
		else
			rem a room has been chosen
			strSQL = strSQL + "AND (tblChemicalContainer.numLocationID = " + numLocationID + ") "
		end if
	else
		strSQL = strSQL + "AND (tblLocation.numCampusID = " + numCampusID + ") "
	end if 
end if
	
	rem IF THE SORT BY NAME CHECK BOX IS NOT SELECTED THEN ORDER BY LOCATION
	strSQL = strSQL + "ORDER BY "
	if strSortEditByName <> "on" then
	strSQL = strSQL +  "tblChemicalContainer.strSpecificLocation, "
	end if
	strSQL = strSQL + "tblChemicalContainer.strChemicalName, tblChemicalContainer.strContainerSize, tblChemicalContainer.strGrade"

'	set rsChemicals = Server.CreateObject("ADODB.Recordset")
	rsChemicals.Open strSQL, conn, 3, 3




if rsChemicals.EOF then
	Response.Write "<BR><DIV align='center'><FONT color='red' face='Arial'>  There are no results for that search , please make your search less specific (try using *, for wildcard search).</FONT></DIV>"
	Response.End
	


end if

	strSQL2 = "SELECT tblBuilding.strBuildingName, tblCampus.strCampusName FROM tblCampus, tblBuilding "
	strSQL2 = strSQL2 + "WHERE tblCampus.numCampusID = " + cstr(rsChemicals("numCampusID"))  + " AND "
	strSQL2 = strSQL2 + "tblBuilding.numBuildingID = " + cstr(rsChemicals("numBuildingID"))  + " AND tblCampus.numCampusID = tblBuilding.numCampusID"

	set rsBuildingCampus = Server.CreateObject("ADODB.Recordset")
	rsBuildingCampus.Open strSQL2, conn

'	if (rsChemicals.EOF) then
'		Response.Write ("There are no Chemicals at this location")
'		Response.End
'	End If

End If%>
<HTML>
<HEAD>

	<TITLE>Update Chemicals</TITLE>
</HEAD>

<BODY>
<% if numLocationID <> "0" then %>
<table border="0" width="100%" id="table1">
	<tr>
		<td width="518">

Supervisor: <%= rsChemicals("strStoreManager") %><br>
Location: <%= rsBuildingCampus("strCampusName") + ", " + rsBuildingCampus("strBuildingName") + ", " +rsChemicals("strStoreLocation")%><BR>
Last Updated: <%= rsChemicals("dtmLastUpdated") %><br/>
Store Type: <%=rsChemicals("strStoreType") + ", " + rsChemicals("strStoreNotes")%><br/>
Location ID: <%= rsChemicals("numLocationID") %>
<% end if %>







<FORM action="UpdateChemicals.asp" method=POST name=frmUpdateChemicals>
<TABLE WIDTH="100%" BORDER=0 ALIGN="center" VALIGN = "TOP" cellpadding="2">
<TR ALIGN="left" VALIGN="top" BGCOLOR="yellow">
	<TD>Delete</TD>
	<TD>Name</TD>
	<TD>Specific Location</TD>
	<td>Quantity</td>
	<TD>Size</TD>
	<TD>Grade</TD>
    <TD>CAS#</TD>
    <TD>Owner</TD>
	<TD>Other Info</td>
    <TD>Haz (Y/N)</td>
    <TD>DG Class</td>
    <%' if rsChemicals("numStoreTypeID") = "1" then %>
	<TD>UN Number</TD>
	<TD>PG (I, II, III)</TD>
<% 'end if %>


	



	
	<TD>SDG Class</td>
	<TD>Hazchem Code</td>
	<TD>Poison Schedule</td>
	<TD>SS (Y/N)</td>
</TR>
    <% 
    numRecordCounter = 0
    do while not rsChemicals.EOF
    numRecordCounter = numRecordCounter + 1 
    numChemicalID = rsChemicals("numChemicalContainerID")
     %>
    <INPUT type="hidden" name=hdnNumChemicalID<%=numRecordCounter%> value=<%=numChemicalID%>>

<TR>
	<TD align="center"><span style="font-size:12px;"><%=numRecordCounter%></span><INPUT type="checkbox" name=chkDelete<%=numRecordCounter%> value="ON"></TD>

	<TD><INPUT type="text" size=25 name=txtChemicalName<%=numRecordCounter%>
		 value="<%= rsChemicals("strChemicalName") %>"></TD>

	<TD><INPUT type="text" size=15 name=txtSpecificLocation<%=numRecordCounter%>
		 value="<%= rsChemicals("strSpecificLocation") %>"></TD>
	
	<TD><INPUT size=6 type="text" name=txtQuantity<%=numRecordCounter%> 
		value="<%= rsChemicals("numQuantity") %>"></TD>

	<TD><INPUT size=6 type="text" name=txtContainerSize<%=numRecordCounter%>
		value="<%= rsChemicals("strContainerSize") %>"></TD>

	<%'if rsChemicals("numStoreTypeID")= "1" then%>

	<!--moved line for container size out of remarekedout loop for clarity -->	
	<!-- removed 'readonly' from above ..ontainerSize readonly <=numRecordCo... -->
	<!-- input type=hidden name=actionValue value="T" -->
	<%'else%>
	
	<%'end if%>

	<TD><INPUT size=6 type="text" name=txtGrade<%=numRecordCounter%> 
	value="<%=rsChemicals("strGrade")%>"></TD>

	<TD><INPUT size=10 type="text" name=txtCAS<%=numRecordCounter%> 
	value="<%= rsChemicals("strCAS") %>"></TD>
    	<TD><INPUT type="text" size=12 name=txtOwner<%=numRecordCounter%> 
	value="<%=rsChemicals("strContainerOwner")%>"></TD>
    	<TD><INPUT type="text" size=15 name=txtOtherInfo<%=numRecordCounter%> 
	value="<%=rsChemicals("strOtherInfo")%>"></TD>
    	<TD align="center">
    <INPUT type="checkbox" name=txtHazardous<%=numRecordCounter%> 
	<%if (rsChemicals("strHazardous") = "Yes") Or (rsChemicals("strHazardous") = "Y")then%> CHECKED
	<%end if%>></TD>
    
	<TD><INPUT type="text" size=3 name=txtDangerousGoodsClass<%=numRecordCounter%> 
	value="<%=rsChemicals("strDangerousGoodsClass")%>"></TD>
<% 

if numlocationID <> "0" then 
'if rsChemicals("numStoreTypeID") = "1" then %>	
	<TD><INPUT size=6 type="text" name=txtUnNumber<%=numRecordCounter%> 
	value="<%= rsChemicals("strUnNumber") %>"></TD>
	
	<TD><INPUT size=3 type="text" name=txtPG<%=numRecordCounter%> 
	value="<%= rsChemicals("strPG") %>"></TD>
<% 
'end if
end if %>
	



	


	
		<TD>
		<INPUT type="text" size=4 name=txtsubsDG<%=numRecordCounter%> 
	value="<%=rsChemicals("strSubsDG")%>"></TD>
	
		<TD>
		<INPUT type="text" size=4 name=txtHazchem<%=numRecordCounter%> 
	value="<%=rsChemicals("strHazchem")%>"></TD>
	
		<TD>
		<INPUT type="text" size=4 name=txtPoisons<%=numRecordCounter%> 
	value="<%=rsChemicals("strPoisons")%>"></TD>
	

	<!-- added or if "Y" above-->
	<TD align="center">
    <INPUT type="checkbox" name=txtSSDG<%=numRecordCounter%> 
	<%if (rsChemicals("strSSDG") = "Yes") then%> CHECKED
	<%end if%>></TD>
	
</TR>
<% 
rsChemicals.MoveNext
loop 
CleanUp()
%>

<TR><TD colspan=10>
<INPUT type="reset" value="Undo Changes" name="RESET">
<!--INPUT type="submit" value="Update Chemical Inventory" name="UPDATE"-->
<!-- value property here is referenced in line 52. Could be better coding. -->
<INPUT type="submit" value="SAVE Changes to Chemical Inventory" name="UPDATE">
<INPUT type="hidden" name=hdnRecordCount value=<%=numRecordCounter%>>
<INPUT type="hidden"  name=hdnNumLocationID value=<%=numLocationID%>></TD></TR>
</TABLE>
</FORM>
</BODY>
</HTML>