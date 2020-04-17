<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%
Sub CleanUp() 

	set rsChemicals = nothing
	conn.close
	set conn = nothing

End Sub

Sub CleanUp2() 

	set rsStoreType = nothing
	conn.close
	set conn = nothing

End Sub

Dim numLocationID
Dim numChemicalID
dim strContainerSize
dim strContainerCheck
Dim strCAS
Dim rsStoreType
Dim rsChemicals
Dim strSQL, strSQL2
Dim numStoreTypeID
Dim conn 
'Dim constr
'	constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

'numLocationID = Request.Form("hdnNumLocationID")
numLocationID = cstr(session("numLocationID"))
strBuildingLocationID = cstr(session("numBuildingID"))
numCampusID = cstr(session("numCampusID")) 

if numLocationID = "0" then
	Response.Write "You must Choose a Location"
	Response.End
end if

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsStoreType = Server.CreateObject("ADODB.Recordset")

strSQL2 = strSQL2 + "SELECT numStoreTypeID FROM tblLocation WHERE numLocationID = " + cstr(numLocationID)
rsStoreType.Open strSQL2, conn, 3, 3
numStoreTypeID = rsStoreType("numStoreTypeID")
set rsStoreType = nothing

If Request.Form("ADD") = "Add New Chemical to the Location" then 
	
'	Dim numLocationID
'	Dim numChemicalID
'	Dim strCAS
'	Dim rsChemicals

'	Dim strSQL
'	Dim conn 
'	Dim constr
'	constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

	For each item in Request.Form
	
		if request.form("txtChemicalName") = "" then
			Response.Write "You must fill in the chemical name. Please click 'back' and fill in the chemical name."
'			Response.Write item
			Response.End
		end if
		
	Next	

'	set conn = Server.CreateObject("ADODB.Connection")
'	conn.open constr
	set rsChemicals = Server.CreateObject("ADODB.Recordset")
	
'Put together CAS number
	strCAS = Request.Form("txtCAS1") + "-" + Request.Form("txtCAS2") + "-" + Request.Form("txtCAS3")
'Put togather the Container size and the unit togather , here "unit" can be L = liters or Kg = kilograms 	
    strContainerCheck = Request.Form("txtContainerSize")
'********************************************************************************************************
 if NOT (IsNumeric(strContainerCheck)) then
    Response.Write "You must fill only numeric value in 'Container Size' field. Please click 'back' and fill in the Container Size."
    Response.Write item
	Response.End
 end if 
'********************************************************************************************************
   if Request.Form("txtPG") ="0" then
         strPGV = " "      
   else
      strPGV = Request.Form("txtPG")
      'change above from strPGV = request("txtPG") to strPGV = request.Form("txtPG")
   end if
   'response.write("Txtpg= " & request("txtPG"))
    strContainerSize = Request.Form("txtContainerSize") + " " + Request.Form("txtContainerUnit")
    
	strSQL = "INSERT INTO tblChemicalContainer "
	strSQL = strSQL + "(strChemicalName, strSpecificLocation, strContainerOwner, strGrade, strContainerSize, numQuantity, strCAS, strOtherInfo, strHazardous,strSSDG, strDangerousGoodsClass,strSubsDG,strHazchem,strPoisons , numLocationID, numSize, strContainerUnits "

	'if numStoreTypeID = "1" then
		strSQL = strSQL + ", strUnNumber, strPG"
	'end if

	strSQL = strSQL + ") VALUES ('" + InjectionEncode(Request.Form("txtChemicalName")) + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtSpecificLocation")) + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtOwner")) + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtGrade")) + "'"
	strSQL = strSQL + ", '" + strContainerSize + "'"
	strSQL = strSQL + ", '" + Request.Form("txtQuantity") + "'"
	strSQL = strSQL + ", '" + strCAS + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtOtherInfo")) + "'"

	strSQL = strSQL + ", '" + Request.Form("txtHazardous") + "'"
	strSQL = strSQL + ", '" + Request.Form("txtSSDG") + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtDangerousGoodsClass")) + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtsubsDG")) + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtHazchem")) + "'"
	strSQL = strSQL + ", '" + InjectionEncode(Request.Form("txtPoisons")) + "'"
	strSQL = strSQL + ", " + Request.Form("hdnNumLocationID") 

	strSQL = strSQL + ", " + Request.Form("txtContainerSize") +" "
	strSQL = strSQL + ", '" + Request.Form("txtContainerUnit") +"'"


		'if numStoreTypeID = "1" then
			strSQL = strSQL + ", '" + Request.Form("txtUnNumber") + "'"
			strSQL = strSQL + ", '" + strPGV + "'"
		'end if

	strSQL = strSQL + + ")"

	'Response.Write (strSQL)
	rsChemicals.Open strSQL, conn, 3, 3
		
	Dim dtmLastUpdated
	dtmLastUpdated = DanDate(Date, "%d/%m/%Y" )
	dtmLastUpdated = cstr(dtmLastUpdated)
	
	set rsLocation = Server.CreateObject("ADODB.Recordset")
	strSQL = "UPDATE tblLocation "	
	strSQL = strSQL + "SET dtmLastUpdated = '" + dtmLastUpdated + "' "
	strSQL = strSQL + "WHERE (numLocationID = " + cStr(numLocationID) + ")"
	
	rsLocation.Open strSQL, conn, 2, 3
	'rsLocation.Close
	
	CleanUp()
	Response.Write ("The Chemical has been Added")
	Response.End
   'end if 	
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
 
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<HTML>
<HEAD>
	<TITLE>Add Chemical</TITLE>
<script language="vbscript">
   sub GetUserName()
	dim V
		V= form.txtContainerSize.value
			if (Not IsNumeric(V)) then
			      msgbox "You typed none-numeric value in the 'Container Size' field ",vbexclamation
			exit sub 
	 	      
	        end if
   end sub
</script>

</HEAD>

<BODY>
<DIV align=center>
<BR><FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
<BR>
<FORM name = Form action="AddChemical.asp" method=POST name=frmAddChemical>
<input type="hidden" name=action value="abc">
<TABLE align=center border=0 cellPadding=1 cellSpacing=5>
    <TR>
		<TD colspan=3><STRONG><FONT color=red 
            face=""><div align="center">Add a Chemical</div></FONT></STRONG><BR><BR></TD>
			
			</TR>
    <TR>
        <TD>Chemical Name:</TD>
        <TD>
            <INPUT name=txtChemicalName style="HEIGHT: 22px; WIDTH: 265px" size="20"></TD>
			<td><FONT SIZE="-1">Chemical name that appears on container label.</FONT></td>
			</TR>

    <TR>
        <TD>Specific Location:</TD>
        <TD>
            <INPUT name=txtSpecificLocation style="HEIGHT: 22px; WIDTH: 265px" size="20"></TD>
			<TD><FONT SIZE="-1">Where in the lab or store is it kept eg Fridge A or Class 3 Cabinet # 1.</FONT> 
</td></TR>
    <TR>
        <TD>Grade:</TD>
        <TD>
            <INPUT name=txtGrade style="HEIGHT: 22px; WIDTH: 265px" size="20"></TD>
			<td><FONT SIZE="-1">eg AR.</FONT></td></TR>
    <TR>
        <TD>Container Size:</TD>
        <TD>
            <%'if numStoreTypeID = 1 then%>
            <INPUT name=txtContainerSize style="HEIGHT: 22; WIDTH: 72" size="20" maxlength="6">&nbsp;&nbsp;&nbsp;
        

		<select name=txtContainerUnit<%=numRecordCounter%> >
			<option value="">--</option>
			<option value="g">g</option>
			<option value="kg">kg</option>
			<option value="mL">mL</option>
			<option value="L">L</option>
			<option value="ug">ug</option>
			<option value="uL">uL</option>
			<option value="packs">packs</option>
			<option value="units">units</option>
			<option value="vials">vials</option>
			<option value="tablets">tablets</option>
			<option value="kits">kits</option>
			<option value="items">items</option>
			<option value="Cylinder-A">Cylinder-A</option>
			<option value="Cylinder-B">Cylinder-B</option>
			<option value="Cylinder-C">Cylinder-C</option>
			<option value="Cylinder-D">Cylinder-D</option>
			<option value="Cylinder-E">Cylinder-E</option>
			<option value="Cylinder-F">Cylinder-F</option>
			<option value="Cylinder-G">Cylinder-G</option>
		</select>
		</TD>
		<td><FONT SIZE="-1">Size of container, not how much is currently in it.</FONT>
			
			
			</td></TR>
			<% 'else %>
			<!--INPUT name=txtContainerSize style="HEIGHT: 22; WIDTH: 72" size="20" maxlength="6"-->
			<%'end if %>
    <TR>
        <TD>Quantity:</TD>
        <TD>
            <INPUT name=txtQuantity style="HEIGHT: 22px; WIDTH: 265px" size="20"></TD>
			<td><FONT SIZE="-1">Number of containers of that particular size and type of chemical.</FONT> </td></TR>
    <TR>
        <TD>CAS #:</TD>
        <TD>
            <!-- increased from maxlength=5 size=5 to 7 chars 28 June 2013 - CLEE > -->
			<INPUT name=txtCAS1 style="HEIGHT: 22px" maxlength="7" size="7"> - <INPUT name=txtCAS2 style="HEIGHT: 22px" maxlength=2 size=2 > - <INPUT name=txtCAS3 style="HEIGHT: 22px" maxlength=1 size=1 ></TD>
			<TD><FONT SIZE="-1">Chemical Abstracts Service registry number.</FONT></td></TR>
    <TR>
        <TD>Owner:</TD>
        <TD>
            <INPUT name=txtOwner style="HEIGHT: 22px; WIDTH: 265px" size="20"></TD>
			<td><FONT SIZE="-1">Person or Area.</FONT></td></TR>
    
<TR>
        <TD>Hazardous? :</TD>
       <TD>
	<INPUT name=txtHazardous type=radio value=Yes>YES
	<INPUT name=txtHazardous type=radio value=No>NO
	</TD>
	<td><FONT SIZE="-1">Statement of hazardous nature is found at the top of the MSDS.</FONT></td>


</TR>
<TR>
        <TD>Dangerous Goods Class:</TD>
        <TD>
            <INPUT name=txtDangerousGoodsClass style="HEIGHT: 22px" maxlength=4 size=4></TD>
			
        <td><FONT SIZE="-1">Dangerous Goods class number e.g. 4.3.</FONT></td>
      </TR>
<% 'if numStoreTypeID = "1" then %>
    <TR>
        <TD><font color="black" face="Arial">UN Number</font>:</TD>
        <TD>
            <INPUT name=txtUnNumber style="HEIGHT: 22px; WIDTH: 265px" maxlength="6" size="20"></TD>
			
        <TD><font size="-1">Chemical UN Number e.g. 1156</font></td>
      </TR>
    <TR>
        <TD>PG (I, II, III):</TD>
        <TD>
            <select size="1" name="txtPG">
			<option value="0">None</option>
			<option value="I">I</option>
			<option value="II">II</option>
			<option value="III">III</option>
			</select></TD>
			<td><FONT SIZE="-1">Chemical PG (Packing Group) Number e.g. I , II, III</FONT></td>
			</TR>
<% 'end if %>
	
<TR>
        <TD>Subsidiary DG Class</TD>
        <TD><FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
            <INPUT name=txtsubsDG style="HEIGHT: 22px" maxlength=4  size="4"></FONT></TD>
        <TD>&nbsp;</TD>
</TR>

<TR>
        <TD>Hazchem Code</TD>
        <TD>
            <FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
            <INPUT name=txtHazchem style="HEIGHT: 22px" maxlength=4 size="4"></FONT></TD>
    			<TD>&nbsp;</TD></TR>

<TR>
        <TD>Poisons Schedule</TD>
        <TD>
            <FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
            <INPUT name=txtPoisons style="HEIGHT: 22px" maxlength=4  size="4"></FONT></TD>
    			<TD>&nbsp;</TD></TR>

<TR>
        <TD>Other Information:</TD>
        <TD>
            <INPUT name=txtOtherInfo style="HEIGHT: 22px; WIDTH: 265px" size="20"></TD>
    			<TD><FONT SIZE="-1">Whatever you want eg Catalogue Number and Supplier.</FONT></TD></TR>
                  <tr>
		<FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
        <TD>Security Sensitive :</TD>
    <TD>
	<INPUT name=txtSSDG type=radio value=Yes>YES
	<INPUT name=txtSSDG type=radio value=No>NO	
	</TD><TD>
	<font size="2">Security Sensitive item e.g. SSAN (Security Sensitive Ammonium Nitrate) or Chemical of Security Concern</span></font></td>


</FONT>
	</tr>

    <TR>
        <TD colspan=2><INPUT type="reset" value="Clear Form" name=btnClear>&nbsp;&nbsp;
			<INPUT type="submit" value="Add New Chemical to the Location" name=ADD onclick ="call GetUserName()"   >
			<INPUT type="hidden"  name=hdnNumLocationID value=<%=numLocationID%>>
        </TD></TR>
        
      
        </Table>
</FORM>
</FONT></DIV>
</BODY>
</HTML>