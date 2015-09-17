<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
<%

'********************************DATABASE CONNECTIVITY CODE ********************************
dim dcnDB ' As ADODB.Connection

Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
dcnDB.Open 

'***********Setting up the recordset *************************************
dim rsQuery
dim rsFillCombo
dim strSQL
dim numCampusID

   numCampusID = Request.QueryString("temp")  
  ' if numCampusID = "1" then
        strSQL = "Select *  from tblEmergencyContact "
		Set rsQuery = Server.CreateObject("ADODB.Recordset")
		rsQuery.Open strSQL,dcnDB 
	
     	
        strSQL = "Select * from tblCampus where numCampusID ="&numCampusID
		Set rsFillCombo = Server.CreateObject("ADODB.Recordset")
		rsFillCombo.Open strSQL,dcnDB 


   dim strCampusName 
   strCampusName = rsFillCombo("strCampusName")

'************************fill combo code ends here****************************************
'*************************filling up the form fields***************
'if numCampusID = "1" then %>
  <form method="POST" action="AddEC.asp?numCampusID=<%=numCampusID%>">
<%'else%>
   
<%'end if %>
		<div align="center">
		<table border="0" width="52%" id="table1">
			<tr>
				<td align="center" colspan="2">
<font face=Arial>
				Add&nbsp; Emergency Contact for <%= strCampusName%>  campus</font></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center"><font face="Arial">
				&nbsp;</font></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Name 1</font></td>
				<td align="center"><font face="Arial">
				<input type="text" name="txtECN1" size="20"  ></font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Position 1</font></td>
				<td align="center"><font face="Arial">
				<input type="text" name="txtECPs1" size="20"  ></font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Phone 1</font></td>
				<td align="center"><font face="Arial">
				<input type="text" name="txtECPh1" size="20"  ></font></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Name 2</font></td>
				<td align="center"><font face="Arial">
				<input type="text" name="txtECN2" size="20" ></font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Position 2</font></td>
				<td align="center"><font face="Arial">
				<input type="text" name="txtECPs2" size="20" ></font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Phone 2</font></td>
				<td align="center"><font face="Arial">
				<input type="text" name="txtECPh2" size="20"  ></font></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center"><font face="Arial">
				<input type="submit" value="Add Emergency Contact" name="btnEdit"></font></td>
			</tr>
		</table>
	</div>
</form>

</BODY>
</HTML>
