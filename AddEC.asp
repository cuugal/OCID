<%@Language = VBScript%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
 <%
 dim numCampusID
 
 numCampusID = Request.QueryString("numCampusID") 

Dim strECName1
Dim strECPosition1
Dim strECPhone1
Dim strECName2
Dim strECPosition2
Dim strECPhone2
dim rsSearch
dim strSQL3
dim numCID
dim numOCuID
dim rsAdd
'********************************DATABASE CONNECTIVITY CODE ********************************
dim dcnDB ' As ADODB.Connection

Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
dcnDB.Open 



strECName1 =(request.form("txtECN1"))
strECPosition1 =(Request.Form("txtECPs1"))
strECPhone1 = (Request.Form("txtECPh1"))
strECName2 = (request.form("txtECN2"))
strECPosition2 =(Request.Form("txtECPs2"))
strECPhone2 = (Request.Form("txtECPh2"))


	strSQL3 = "INSERT INTO tblEmergencyContact (strEmergencyContactName1, strEmergencyContactPosition1,"_
	&"strEmergencyContactPhone1 ,strEmergencyContactName2 ,strEmergencyContactPosition2,"_
	&"strEmergencyContactPhone2,numCampusID)"_
	&" VALUES('"& strECName1 &"','"& strECPosition1 &"' ,'"& strECPhone1 &"' , '"& strECName2 &"',"_
	&"'"& strECPosition2 &"','"& strECPhone2 &"', '"& numCampusID &"' )"
 
 	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strSQL3,dcnDB
	
    Response.Write "The emergency contact has been added sucessfully !" 	

%>

</BODY>
</HTML>
