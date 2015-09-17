

<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->



<% dim strDate
strDate = DanDate(Date, "%d/%m/%Y" )%>


<% Sub RefreshMenu()
	
	Dim Output
	Output = "<SCRIPT LANGUAGE=javascript><!-- " + chr(13)
	Output = Output + "parent.frames['Search'].location.reload() " + chr(13)
	Output = Output + "//--></SCRIPT>"
	
	Response.Write(OutPut)
	
End Sub

Dim checkAction
Dim checkArea

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>ChooseLoginID</title>
</head>

<body>

<div align="center">
<font face=Arial>
Edit Database:<BR>
<FONT COLOR=#f70932 SIZE="+1"><B>NOTE: If you delete a building or floor, all its chemical inventory is also deleted</B>.</FONT>
<P><BR>
<FORM action="EditPreferences.asp" method=POST name=frmEditPerferences>
	<select NAME=cboAction>
	<option value="0">Please select an action</option>
	<option value="1">Add</option>
	<option value="2">Edit</option>
	<option value="3">Delete</option>
    </select>

	<select NAME=cboArea size="1">
	<option value="0">Please select an edit area</option>
	<option value="8">Occupier</option>
	<option value="1">Building</option>
	<option value="2">Floor</option>
	<option value="7">Emergency Contact</option>
<!--
	<option value="3">Store Location</option>
	<option value="4">Store Type</option>
-->
    </select>
</td>
	
	&nbsp;&nbsp;<input type="submit" name="btnSubmit" value="Next">
</FORM></div>

<form name=frmEditNow action="EditPreferences.asp" method=POST>
<% call actionPerformed

Dim rsRecord
Dim rsBuilding
Dim rsStoreLocation
Dim rsStoreType
Dim conn
Dim strSQL, strSQL2
'Dim constr

Sub findCampus(numCheck)

'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsRecord = Server.CreateObject("ADODB.Recordset")
set rsBuilding = Server.CreateObject("ADODB.Recordset")
set rsStoreLocation = Server.CreateObject("ADODB.Recordset")
set rsStoreType = Server.CreateObject("ADODB.Recordset")

'if checkAction = 2 then
'	checkArea = checkArea + 1
'end if

'checkArea = checkArea - 1

if numCheck = 8 then
	strSQL = "SELECT * FROM tblOccupier"
	rsRecord.Open strSQL, conn, 2, 3 
	Response.Write "Existing Occupier Name : " %>
 	   
          <% do while not rsRecord.EOF %>
          <%= rsRecord("strOccupierName") %>
          
          <% numOccupier = rsRecord("numOccupierID")
           rsRecord.MoveNext
	loop
	rsRecord.Close
	set rsRecordLocation = nothing
'	connCampus.close
'	set connCampus = nothing
%>  
<% end if


if numCheck = 2 then
	strSQL = "SELECT * FROM tblCampus"
	rsRecord.Open strSQL, conn, 2, 3 %>
 	<select name="cboCampus">
          <option value="0">Please select</option>
          <% do while not rsRecord.EOF %>
          <option value="<%=rsRecord("numCampusID")%>"><%= rsRecord("strCampusName") %></option>
          <% rsRecord.MoveNext
	loop
	rsRecord.Close
	set rsRecordLocation = nothing
'	connCampus.close
'	set connCampus = nothing
%> </select> 

<% end if

if numCheck = 3 then
	strSQL2 = "SELECT * FROM tblBuilding"
	rsBuilding.Open strSQL2, conn, 2, 3 %>
 	<select name="cboBuilding">
          <option value="0">Please select</option>
          <% do while not rsBuilding.EOF %>
          <option value="<%=rsBuilding("numBuildingID")%>"><%= rsBuilding("strBuildingName") %></option>
          <% rsBuilding.MoveNext
	loop
	rsBuilding.Close
	set rsBuilding = nothing

%> </select>

<% end if

if numCheck = 4 then
	strSQL2 = "SELECT * FROM tblStoreLocation"
	rsStoreLocation.Open strSQL2, conn, 2, 3 %>
 	<select name="cboStoreLocation">
          <option value="0">Please select</option>
          <% do while not rsStoreLocation.EOF %>
          <option value="<%=rsStoreLocation("numStoreLocationID")%>"><%= rsStoreLocation("strStoreLocation") %></option>
          <% rsStoreLocation.MoveNext
	loop
	rsStoreLocation.Close
	set rsStoreLocation = nothing

%> </select>
<% end if

if numCheck = 5 then
	strSQL2 = "SELECT * FROM tblStoreType"
	rsStoreType.Open strSQL2, conn, 2, 3 %>
 	<select name="cboStoreType">
          <option value="0">Please select</option>
          <% do while not rsStoreType.EOF %>
          <option value="<%=rsStoreType("numStoreTypeID")%>"><%= rsStoreType("strStoreType") %></option>
          <% rsStoreType.MoveNext
	loop
	rsStoreType.Close
	set rsStoreType = nothing

%> </select>
<% end if

end sub

function editNewArea(areaName, iNum)
   
    response.write "Select a Occupier for an emergency contact : " 
    call findCampus(8)
    
    response.write "Emergency Contact Name 1 : "
	response.write "<input type=text name=txtECN1>" %><BR><BR><%
	
	response.write "Emergency Contact Position 1 : "
	response.write "<input type=text name=txtECPs1>" %><BR><BR><%
	
	response.write "Emergency Contact Phone 1 : "
	response.write "<input type=text name=txtECPh1> " %><BR><BR><%
	
	response.write "Emergency Contact Name 2 : "
	response.write "<input type=text name=txtECN2> " %><BR><BR><%
	
	response.write "Emergency Contact Position 2 : "
	response.write "<input type=text name=txtECPs2> " %><BR><BR><%
	
	response.write "Emergency Contact Phone 2 : "
	response.write "<input type=text name=txtECPh2> " %><BR><BR><%
	
    response.write "<input type=button name=btn" + areaName + " value='Edit " + areaName + "' onClick='javascript: document.frmEditNow.submit()'> <br>"
	response.write "<input type=hidden name=actionValue value='edit" + cstr(7) + "'>"

end function


'****************************************function to Edit the Emergency Contact details*******************************
function editAreaEC(areaName, inNum)
    dim id
    dim temp
    temp = Request.Form("cboCampus")  
    response.write "<input type=button name=btn" + areaName + " value='Edit " + areaName + "' onClick='javascript: document.frmEditNow.submit()'> <br>"
    response.write "<input type=hidden name=actionValue value=" + cstr(inNum) + ">"
	
end function
'*********************************************End of Function*************************************************************
'****************************************function to delete the Emergency Contact details*******************************
function deleteAreaEC(areaName, inNum)
    response.write "<input type=button name=btn" + areaName + " value='Delete " + areaName + "' onClick='javascript: document.frmEditNow.submit()'> <br>"
	response.write "<input type=hidden name=actionValue value=" + cstr(inNum) + ">"
	
end function
'*********************************************End of Function*************************************************************

'****************************************function to add the Emergency Contact details*******************************
function addAreaEC(areaName, inNum)
   
  	response.write "<input type=button name=btn" + areaName + " value='Add " + areaName + "' onClick='javascript: document.frmEditNow.submit()'> <br>"
	response.write "<input type=hidden name=actionValue value=" + cstr(inNum) + ">"
	
end function
'*********************************************End of Function*************************************************************
function addArea(areaName, inNum)
	response.write "<input type=text name=txt" + areaName + ">  "
	response.write "<input type=button name=btn" + areaName + " value='Add " + "'                   onClick='javascript: document.frmEditNow.submit()'> <br>"
	response.write "<input type=hidden name=actionValue value=" + cstr(inNum) + ">"
end function

function editArea(areaName, inNum)
	response.write "<input type=text name=txt" + areaName + ">  "
	response.write "<input type=button name=btn" + areaName + " value='Edit " + "'                onClick='javascript: document.frmEditNow.submit()'> <br>"
	response.write "<input type=hidden name=actionValue value='edit" + cstr(inNum) + "'>"
end function

function delArea(areaName, inNum)
'	response.write "<input type=text name=txt" + areaName + ">  "
	response.write "<input type=button name=btn" + areaName + " value='Delete " + "'			onClick='javascript: document.frmEditNow.submit()'> <br>"
	response.write "<input type=hidden name=actionValue value='del" + cstr(inNum) + "'>"
end function
' Deleted areaName + from addArea and editArea functions from button value
Sub actionPerformed()

checkAction = request("cboAction")
checkArea = request("cboArea")

	if checkAction = 1 then
		if checkArea = 1 then
		    Response.Write "Adding new building name" %> <BR><BR><%
			response.write "New Building Name: "
			call addArea("Campus", 1)
		end if
		if checkArea = 2 then
			response.write "New Floor Name: " %> <BR><BR><%
			response.write "Belongs to : "
			call findCampus(2)	
			Response.Write "Adding new floor name"
			call addArea("Building", 2)
'			call CleanUp	
		end If
		if checkArea = 3 then
		    Response.Write "Adding new store Location name" %> <BR><BR><%
			response.write "New Store Location: "
			call addArea("StoreLocation", 3)
			response.write "Belongs to : "
			call findCampus(3)	
'			call CleanUp	
		end if
		if checkArea = 4 then
			Response.Write "Adding new store type name" %> <BR><BR><%
			response.write "New Store Type: "
			call addArea("StoreType", 4)
		end if
		
		if checkArea = 7 then
			Response.Write "Adding new emergency contact" %> <BR><BR><%
		    response.write "Please Select the building:"
		    call findCampus(2)%><Br><Br><%
		    call addAreaEC("EmergencyContact", 7)
		end if
	   	if checkArea = 8 then
		   
		    ' Code to check if there is a existing occupier******************************************
		      dim rsChkOcc
		      dim strChkOcc
		      dim flg
		      dim conn
		      dim constr
		      
		      constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
              set conn = Server.CreateObject("ADODB.Connection")
              conn.open constr

		      flg = 0
		      
		      strChkOcc = "select * from tblOccupier" 
			  set rsChkOcc = Server.CreateObject("ADODB.Recordset")
			  rsChkOcc.Open strChkOcc,conn,2,3 
			  
			   while not rsChkOcc.EOF 
			      if rsChkOcc("strOccupierName")<>"" then
			         flg = 1       
			      end if
			         rsChkOcc.MoveNext     
			   wend
			       if flg = 1 then
						Response.Write "One occupier already exists ! cannot add more than one occupier. "
			       else
			            Response.Write "Adding new occupier " %><BR><BR><%
			            response.write "New Occupier : " 
	    				call addArea("Occupier", 8)
	    		   end if
	    		   
	    	 ' code ends here*************************************************************************	   
	    	end if
      
elseif checkAction = 2 then
		if checkArea = 1 then
		    Response.Write "Editing building name" %> <BR><BR><%
			response.write "Existing building Name: "
			call findCampus(2)	
			response.write "New Building Name: "
			call editArea("Campus", 1)
		end if
		if checkArea = 2 then
			Response.Write "Editing floor name" %> <BR><BR><%
			response.write "Existing floor Name: "
			call findCampus(3)	
			response.write "New Floor Name: "
			call editArea("Building", 2)
		end if
		if checkArea = 3 then
			Response.Write "Editing  store location name" %> <BR><BR><%
			response.write "Existing Store Location: "
			call findCampus(4)	
			response.write "New Store Location: "
			call editArea("StoreLocation", 3)
		end if
		if checkArea = 4 then
			Response.Write "Editing store type" %> <BR><BR><%
			Response.Write "Editing new store type" %> <BR><%
			response.write "Existing Store Type: "
			call findCampus(5)	
			response.write "New Store Type: "
			call editArea("StoreType", 4)
		end if
		if checkArea = 7 then
		   Response.Write "Editing  emergency contact" %> <BR><BR><%
		   response.write "Select a campus for an emergency contact :  " 
			call findCampus(2)
			call editAreaEC("EmergencyContact", 9)
		end if
		if checkArea = 8 then
			response.write "Editing Occupier Name " %> <BR><BR><%
			call findCampus(8)	%><BR><BR><%
			response.write "New Occupier Name: "
			call editArea("Occupier", 8)
		end if
		
	elseif checkAction = 3 then
	    if checkArea = 7 then
			response.write "Deleting emergency contact" %> <BR><BR><%
			response.write "Select a campus for an emergency contact : " 
			call findCampus(2)	
			call deleteAreaEC("Emergency Contact", 10)
	    
		elseif checkArea = 1 then
			response.write "Deleting building name" %> <BR><BR><%
			response.write "Existing Building Name: "
			call findCampus(2)	
'			response.write "New Building Name: "
			call delArea("Campus", 1)
		end if
		if checkArea = 2 then
			response.write "Deleting floor name" %> <BR><BR><%
			response.write "Existing Floor Name: "
			call findCampus(3)	
'			response.write "New Floor Name: "
			call delArea("Building", 2)
		end if
		if checkArea = 3 then
			response.write "Deleting store location" %> <BR><BR><%
			response.write "Existing Store Location: "
			call findCampus(4)	
'			response.write "New Store Location: "
			call delArea("StoreLocation", 3)
		end if
		if checkArea = 4 then
			response.write "Deleting store type" %> <BR><BR><%
			response.write "Existing Store Type: "
			call findCampus(5)	
'			response.write "New Store Type: "
			call delArea("StoreType", 4)
		end if
		if checkArea = 8 then
			response.write "Cannot delete the Occupier !"
			'call findCampus(8)	
'			response.write "New Store Type: "
			'call delArea("Occupier", 8)
		end if
	end if
end sub

Sub CleanUp()
conn.close
set conn = nothing
end sub
%>
</font>
</form>
</body>
</html>
<%
Dim inAction
Dim strSQL3
Dim rsAdd

inAction = request.form("actionValue")
'response.write (request.form("actionValue"))

'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

'if checkAction = 1 then
if inAction = "10" then
dim temp 
    dim rsFill
    
    'temp = Request.form("cboCampus")
    '*****************************************************************
        	      
		      %><a href = "DeleteEmergencyContact.asp?temp=<%=Request.Form("cboCampus")%>">Click here to access the Delete emergency contact form</a> 
		   <%
  
    '*****************************************************************
    
    
elseif inAction = "9" then
  
   %><a href = "EditEmergencyContact.asp?temp=<%=Request.Form("cboCampus")%>">Click here to access the Edit emergency contact form</a> <%
'******************Code to add new Ocuupier***************************************************************
elseif inAction = "8" then

dim rsAddOccupier
dim strSQLAC
dim numOID
dim OccupierName

OccupierName = (request.form("txtOccupier"))
if OccupierName <> "" then
	strSQLAC = "INSERT INTO tblOccupier (strOccupierName) VALUES('"& OccupierName &"')"
	set rsAddOccupier = Server.CreateObject("ADODB.Recordset")
	rsAddOccupier.Open strSQLAC, conn, dynaset, 3	
	set rsAddOccupier = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The Occupier Name has been added"
Response.End
call CleanUp
else
Response.Write "Please fill in a Occupier name"
response.End
end if
'******************Code ends here************************************************
'******************Code to add new emergency contact***************************************************************
else if inAction = "7" then
    
    dim rsChkCEC
    dim strSQLCEC
    temp= Request.form("cboCampus")
       
        strSQLCEC = "select numCampusId from tblEmergencyContact where numCampusID ="&temp
        set rsChkCEC = Server.CreateObject("ADODB.Recordset")
		rsChkCEC.Open strSQLCEC, conn, dynaset, 3	
		if rsChkCEC.EOF = false then
		   Response.Write "Cannot add more than one emergency contact !"
		   else 
		      
		      %><a href = "addEmergencyContact.asp?temp=<%=Request.Form("cboCampus")%>">Click here to access the emergency contact form</a> 
		   <%end if
  
'******************Code ends here***************************************************************
else if inAction = "1" then
Dim campusName
' Adding a BUILDING - NOTE: CAMPUS refers to building name
campusName = (request.form("txtCampus"))

if campusName <> "" then
	strSQL3 = "INSERT INTO tblCampus "
	strSQL3 = strSQL3 + "(strCampusName) "
	strSQL3 = strSQL3 + "VALUES ('" + InjectionEncode(campusName) + "')"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The building has been added"
Response.End
call CleanUp
else
Response.Write "Please fill in a building name"
response.End
end if

elseif inAction = "2" then

Dim buildingName
Dim numCampusID
' Adding a FLOOR - NOTE: BUILDING refers to floor name
buildingName = (request.form("txtBuilding"))
numCampusID = (request.form("cboCampus"))

if buildingName <> "" AND numCampusID <> "0" then
	strSQL3 = "INSERT INTO tblBuilding "
	strSQL3 = strSQL3 + "(strBuildingName, numCampusID) "
	strSQL3 = strSQL3 + "VALUES ('" + InjectionEncode(buildingName) + "'," + cstr(numCampusID) + ")"
'Response.Write strSQL3	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The floor has been added"
Response.End
call CleanUp
else
Response.Write "Please fill in a floor name and/or select a building"
response.End
end if

elseif inAction = "3" then

Dim storeLocationName
Dim numBuildingID

storeLocationName = (request.form("txtStoreLocation"))
numBuildingID = (request.form("cboBuilding"))

if storeLocationName <> "" AND numBuildingID <> "0" then
	strSQL3 = "INSERT INTO tblStoreLocation "
	strSQL3 = strSQL3 + "(strStoreLocation, numBuildingID) "
	strSQL3 = strSQL3 + "VALUES ('" + InjectionEncode(storeLocationName) + "', " + cstr(numBuildingID) + ")"

	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The store location has been added"
Response.End
call CleanUp
else
Response.Write "Please fill in a store location name and/or select a building"
response.End
end if

elseif inAction = "4" then

Dim storeTypeName
'Dim numBuildingID

storeTypeName = (request.form("txtStoreType"))
'numBuildingID = (request.form("cboBuilding"))

if storeTypeName <> "" then
	strSQL3 = "INSERT INTO tblStoreType "
	strSQL3 = strSQL3 + "(strStoreType) "
	strSQL3 = strSQL3 + "VALUES ('" + InjectionEncode(storeTypeName) + "')"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The store type has been added"
Response.End
call CleanUp
else
Response.Write "Please fill in a store type"
response.End
end if
end if
end if
end if
'****************************new edit code for the Emergency Contact*************************************



'************************************edit code for the Emergency Contact finishes here*******************

'****************************new edit code for the occupier*************************************
if inAction = "edit8" then

Dim editOccupierName
Dim numOccupier
'***********************************************************************************************
strSQL = "SELECT * FROM tblOccupier"
set rsRecord = Server.CreateObject("ADODB.Recordset")
rsRecord.Open strSQL, conn, 2, 3 
	'Response.Write "Existing Occupier Name : " %><% do while not rsRecord.EOF
           numOccupier = rsRecord("numOccupierID")
           rsRecord.MoveNext
	loop
	rsRecord.Close
	set rsRecordLocation = nothing
'	connCampus.close
'**********************************************************************************************
editOccupierName = (request.form("txtOccupier"))
'numOccupier = (request.form("cboOccupier"))

if editOccpierName <> " " then 'AND numOccupier <> "0" then
	strSQL3 = "UPDATE tblOccupier "	
	strSQL3 = strSQL3 + "SET strOccupierName = '" + InjectionEncode(editOccupierName) + "' "
	strSQL3 = strSQL3 + "WHERE (numOccupierID = " + cStr(numOccupier) + ")"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The occupier has been updated"
Response.End
call CleanUp
else
Response.Write "Please select a occupier to update and/or fill in a occupier name "
response.End
end if

'************************************edit code for the occupier finishes here*******************
elseif inAction = "edit1" then

Dim editCampusName
Dim numCampus

editCampusName = (request.form("txtCampus"))
numCampus = (request.form("cboCampus"))

if editCampusName <> "" AND numCampus <> "0" then
	strSQL3 = "UPDATE tblCampus "	
	strSQL3 = strSQL3 + "SET strCampusName = '" + InjectionEncode(editCampusName) + "' "
	strSQL3 = strSQL3 + "WHERE (numCampusID = " + cStr(numCampus) + ")"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The building has been updated"
Response.End
call CleanUp
else
Response.Write "Please select a building to update and/or fill in a building name "
response.End
end if

elseif inAction = "edit2" then

Dim editBuildingName
Dim numBuilding

editBuildingName = (request.form("txtBuilding"))
numBuilding = (request.form("cboBuilding"))

if editBuildingName <> "" AND numBuilding <> "0" then
	strSQL3 = "UPDATE tblBuilding "
	strSQL3 = strSQL3 + "SET strBuildingName = '" + InjectionEncode(editBuildingName) + "' "
	strSQL3 = strSQL3 + "WHERE (numBuildingID = " + cStr(numBuilding) + ")"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The building has been updated"
Response.End
call CleanUp
else
Response.Write "Please select a building to update and/or fill in a building name "
response.End
end if

elseif inAction = "edit3" then

Dim editLocationName
Dim numLocation

editLocationName = (request.form("txtStoreLocation"))
numLocation = (request.form("cboStoreLocation"))

if editLocationName <> "" AND numLocation <> "0" then
	strSQL3 = "UPDATE tblStoreLocation "
	strSQL3 = strSQL3 + "SET strStoreLocation = '" + InjectionEncode(editLocationName) + "' "
	strSQL3 = strSQL3 + "WHERE (numStoreLocationID = " + cStr(numLocation) + ")"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The store location has been updated"
Response.End
call CleanUp
else
Response.Write "Please select a store location to update and/or fill in a store location name "
response.End
end if

elseif inAction = "edit4" then

Dim editStoreTypeName
Dim numStoreType

editStoreTypeName = (request.form("txtStoreType"))
numStoreType = (request.form("cboStoreType"))

if editStoreTypeName <> "" AND numStoreType <> "0" then
	strSQL3 = "UPDATE tblStoreType "
	strSQL3 = strSQL3 + "SET strStoreType = '" + InjectionEncode(editStoreTypeName) + "' "
	strSQL3 = strSQL3 + "WHERE (numStoreTypeID = " + cStr(numStoreType) + ")"
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The store type has been updated"
Response.End
call CleanUp
else
Response.Write "Please select a store type to update and/or fill in a store type"
response.End
end if
elseif inAction = "edit7" then

Dim strEECName1
Dim strEECPosition1
Dim strEECPhone1
Dim strEECName2
Dim strEECPosition2
Dim strEECPhone2
dim rsESearch
dim strESQLSearch
campusName = (request.form("cboCampus"))
OccupierName = (request.form("cboOccupier"))


strEECName1 =(request.form("txtECN1"))
strEECPosition1 =(Request.Form("txtECPs1"))
strEECPhone1 = (Request.Form("txtECPh1"))
strEECName2 = (request.form("txtECN2"))
strEECPosition2 =(Request.Form("txtECPs2"))
strEECPhone2 = (Request.Form("txtECPh2"))

if editOccpierName <> " " then 
	 strESQLSearch = "Update tblEmergencyContact "_
	 &"SET "_  	 
	 &"strEmergencyContactName1 = '"& strEECName1 &"',"_
	 &" strEmergencyContactPosition1='"& strEECPosition1 &"',"_
	 &" strEmergencyContactPhone1= '"& strEECPhone1 &"',"_
	 &" strEmergencyContactName2= '"& strEECName2 &"',"_
	 &" strEmergencyContactPosition2= '"& strEECPosition2 &"',"_
	 &" strEmergencyContactPhone2='"& strEECPhone2 &"',"_
	 &" numOccupierID='"& OccupierName &"' where numCampusID='"& campusName & "' "
	 

	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strESQLSearch, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The Emergency Contact has been updated"
Response.End
call CleanUp
else
Response.Write "Please select a Building to update and/or fill in a occupier name "
response.End
end if
end if
'end if
if inAction = "del1" then

Dim delnumCampus

'editCampusName = (request.form("txtCampus"))
delnumCampus = (request.form("cboCampus"))

if delnumCampus <> "0" then
	strSQL3 = "Delete FROM tblCampus "
	strSQL3 = strSQL3 + "WHERE numCampusID = " + cstr(delnumCampus)
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The building has been deleted"
Response.End
call CleanUp
else
Response.Write "Please select a building to delete"
response.End
end if

elseif inAction = "del2" then

Dim delnumBuilding

'editCampusName = (request.form("txtCampus"))
delnumBuilding = (request.form("cboBuilding"))

if delnumBuilding <> "0" then
	strSQL3 = "Delete FROM tblBuilding "
	strSQL3 = strSQL3 + "WHERE numBuildingID = " + cstr(delnumBuilding)
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The floor has been deleted"
Response.End
call CleanUp
else
Response.Write "Please select a floor to delete"
response.End
end if

elseif inAction = "del3" then

Dim delnumStoreLocation

'editCampusName = (request.form("txtCampus"))
delnumStoreLocation = (request.form("cboStoreLocation"))

if delnumStoreLocation <> "0" then
	strSQL3 = "Delete FROM tblStoreLocation "
	strSQL3 = strSQL3 + "WHERE numStoreLocationID = " + cstr(delnumStoreLocation)
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The store location has been deleted"
Response.End
call CleanUp
else
Response.Write "Please select a store location to delete"
response.End
end if

elseif inAction = "del4" then

Dim delnumStoreType

'editCampusName = (request.form("txtCampus"))
delnumStoreType = (request.form("cboStoreType"))

if delnumStoreType <> "0" then
	strSQL3 = "Delete FROM tblStoreType "
	strSQL3 = strSQL3 + "WHERE numStoreTypeID = " + cstr(delnumStoreType)
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The store type has been deleted"
Response.End
call CleanUp
else
Response.Write "Please select a store type to delete"
response.End
end if

elseif inAction = "del8" then

Dim delnumOccupierId
delnumOccupierId = (request.form("cboOccupier"))

if delnumOccupierId <> "0" then
	strSQL3 = "Delete FROM tblOccupier "
	strSQL3 = strSQL3 + "WHERE numOccupierID = " + cstr(delnumOccupierId)
	
	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3, conn, dynaset, 3	
	set rsAdd = nothing
'conn.close
'set conn = nothing
RefreshMenu()
Response.Write "The Occupier has been deleted"
Response.End
call CleanUp
else
Response.Write "Please select a Occupier to delete"
response.End
end if

end if

select case inAction

    
case "19"
        tempo = Request.form("cboCampus") 
		Response.Redirect "editEmergencyContact.asp?tempo"
end select		
%>