<%@Language = VBScript%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
<%
dim numCampusID
dim strEECName1
Dim strEECPosition1
Dim strEECPhone1
Dim strEECName2
Dim strEECPosition2
Dim strEECPhone2
dim rsESearch
dim strESQLSearch
dim strOccupier 
dim rsAdd


'********************************DATABASE CONNECTIVITY CODE ********************************
dim dcnDB ' As ADODB.Connection

Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
dcnDB.Open 


   numCampusID = Request.QueryString("numCampusID")
   
   strOccupier = Request.Form("CBoOccupier") 
		strEECName1 =(request.form("txtECN1"))
		strEECPosition1 =(Request.Form("txtECPs1"))
		strEECPhone1 = (Request.Form("txtECPh1"))
		strEECName2 = (request.form("txtECN2"))
		strEECPosition2 =(Request.Form("txtECPs2"))
		strEECPhone2 = (Request.Form("txtECPh2"))

   
   if numCampusID = "1" then
        strESQLSearch = "Delete from tblEmergencyContact where numCampusID=" &numCampusID
	 	set rsAdd = Server.CreateObject("ADODB.Recordset")
		rsAdd.Open strESQLSearch, dcnDB
		Response.Write "The Emergency Contact has been deleted."
		Response.End
elseif numCampusID = "2" then
     
'else
'Response.Write "Please select a Campus to update and/or fill in a occupier name "
'response.End
   		strESQLSearch = "Delete from tblEmergencyContact where numCampusID=" &numCampusID
	 	set rsAdd = Server.CreateObject("ADODB.Recordset")
		rsAdd.Open strESQLSearch, dcnDB
		Response.Write "The Emergency Contact has been deleted."
		Response.End

	set rsAdd = Server.CreateObject("ADODB.Recordset")
	rsAdd.Open strESQLSearch, dcnDB

	Response.Write "The Emergency Contact has been deleted."
	Response.End
end if  
%>

</BODY>
</HTML>
