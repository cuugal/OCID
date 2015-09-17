<%@ Language=VBScript %>
<!--#INCLUDE FILE="DbConfig.asp"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<% 
Dim rsChemicals, rsLocation, rsStoreType, rsBuildingCampus, rsStoreLocation
Dim strSQL, strSQL2, strSQL3, strSQL4
Dim strSQL5
Dim conn
Dim numChemicalContainerID, numLocationID
Dim strLocation

numChemicalContainerID = Request.QueryString("numChemicalID")
numLocationID = Request.QueryString("numLocationID")

strSQL = "SELECT * FROM tblChemicalContainer "
strSQL = strSQL + "WHERE numChemicalContainerID = " + numChemicalContainerID

strSQL2 = strSQL2 + "SELECT * FROM tblLocation "
strSQL2 = strSQL2 + "WHERE numLocationID = " + numLocationID

strSQL3 = strSQL3 + "SELECT tblStoreType.strStoreType FROM tblLocation, tblStoreType "
strSQL3 = strSQL3 + "WHERE tblLocation.numLocationID = " + numLocationID + " AND tblLocation.numStoreTypeID = tblStoreType.numStoreTypeID"

'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr

set rsChemicals = Server.CreateObject("ADODB.Recordset")
rsChemicals.Open strSQL, conn, 3, 3
set rsLocation = Server.CreateObject("ADODB.Recordset")
rsLocation.Open strSQL2, conn, 3, 3
set rsStoreType = Server.CreateObject("ADODB.Recordset")
rsStoreType.Open strSQL3, conn, 3, 3

strSQL4 = "SELECT tblBuilding.strBuildingName, tblCampus.strCampusName FROM tblCampus, tblBuilding "
strSQL4 = strSQL4 + "WHERE tblCampus.numCampusID = " + cstr(rsLocation("numCampusID"))  + " AND "
strSQL4 = strSQL4 + "tblBuilding.numBuildingID = " + cstr(rsLocation("numBuildingID"))  + " AND tblCampus.numCampusID = tblBuilding.numCampusID"

strSQL5 = "SELECT * FROM tblStoreLocation "
strSQL5 = strSQL5 + "WHERE numStoreLocationID = " + cstr(rsLocation("numStoreLocationID"))

set rsBuildingCampus = Server.CreateObject("ADODB.Recordset")
rsBuildingCampus.Open strSQL4, conn, 3, 3

set rsStoreLocation = Server.CreateObject("ADODB.Recordset")
rsStoreLocation.Open strSQL5, conn, 3, 3

strLocation = rsBuildingCampus("strCampusName") + ", " + rsBuildingCampus("strBuildingName") + ", " + rsStoreLocation("strStoreLocation")
%>
</HEAD>
<BODY>

<DIV align=center>
<BR>
<BR>
<TABLE align=center border=0 cellPadding=1 cellSpacing=10>
<FONT color=black face=Arial style="BACKGROUND-COLOR: #ffffff">
    <TBODY>
    <TR>
		<TD><STRONG><FONT color=red 
            face=""><FONT face=Arial><FONT>Chemical 
            Details</FONT></FONT></FONT></STRONG><FONT><FONT face=Arial><BR><BR></FONT></FONT></TD>
	</TR>
    <TR>
        <TD><FONT face=Arial><STRONG>Chemical 
            Name:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strChemicalName") %>
            </FONT></TD>
	</TR>
    <TR>
        <TD><FONT face=Arial><STRONG>Specific Location:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strSpecificLocation") %>
            </FONT></TD>
	</TR>
	<TR>
        <TD><FONT face=Arial><STRONG>Grade:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strGrade") %>
            </FONT></TD>
	</TR>
    <TR>
        <TD><FONT face=Arial><STRONG>Store Type:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsStoreType("strStoreType") %>
            </FONT></TD>
	</TR>
    <TR>
        <TD><FONT face=Arial><STRONG>Container 
            Size:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strContainerSize") %>
            </FONT></TD>
	</TR>
    <TR>
        <TD><FONT 
            face=Arial><STRONG>Quantity:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("numQuantity") %>
            </FONT></TD>
	</TR>
    <TR>
        <TD><FONT face=Arial><STRONG>CAS 
            #:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strCAS") %>
            </FONT></TD>
	</TR>
    <TR>
        <TD><FONT 
            face=Arial><STRONG>Owner:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
             </STRONG><%= rsChemicals("strContainerOwner") %>
           </FONT></TD>
	</TR>
    <TR>
        <TD><FONT 
            face=Arial><STRONG>Other Information:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
             </STRONG><%= rsChemicals("strOtherInfo") %>
           </FONT></TD>
	</TR>  
    <TR>
        <TD><FONT 
            face=Arial><STRONG>Location:</STRONG></FONT></TD>
        <TD>
            <FONT face=Arial><%= strLocation %></FONT></TD></TR>

    <TR>
        <TD><FONT 
            face=Arial><STRONG>Hazardous ?:</STRONG></FONT></TD>
        <TD>


<FONT face=Arial><%= rsChemicals("strHazardous")%></FONT>

</TD></TR>

    <TR>            
        <TD><FONT face=Arial><STRONG>DangerousGoodsClass:</STRONG></FONT></TD>
        <TD><FONT face=Arial><%= rsChemicals("strDangerousGoodsClass")%></FONT></TD>
	</TR>

    <% 'if rsLocation("numStoreTypeID") = "1" then %>
	<TR>
        <TD><FONT 
            face=Arial><STRONG>UN Number:</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strUnNumber") %>
            </FONT></TD>
	</TR>
	<TR>
        <TD><FONT 
            face=Arial><STRONG>PG (I, II, III):</STRONG></FONT></TD>
        <TD><FONT face=Arial><STRONG>
            </STRONG><%= rsChemicals("strPG") %>
            </FONT></TD>
	</TR>
<%' end if %>

    <TD><FONT face=Arial><STRONG>Subsidary DG Class</STRONG></FONT></TD>
        <TD><FONT face=Arial><%= rsChemicals("strSubsDG")%></FONT></TD></TR>

	<TD><FONT face=Arial><STRONG>Hazchem Code</STRONG></FONT></TD>
        <TD><FONT face=Arial><%= rsChemicals("strHazChem")%></FONT></TD></TR>

	<TD><FONT face=Arial><STRONG>Poisons Schedule</STRONG></FONT></TD>
        <TD><FONT face=Arial><%= rsChemicals("strPoisons")%></FONT></TD></TR>

	<TD><FONT face=Arial><STRONG>Security Sensitive</STRONG></FONT></TD>
        <TD><FONT face=Arial><%= rsChemicals("strSSDG")%></FONT></TD></TR>

    
    
    
</TBODY></FONT> 
</TABLE></FORM>
</DIV>

</BODY>
</HTML>
