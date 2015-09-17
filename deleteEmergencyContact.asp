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
   'if numCampusID = "1" then
    
        strSQL = "Select *  from tblEmergencyContact where numCampusID ="&numCampusID
        Set rsQuery = Server.CreateObject("ADODB.Recordset")
		rsQuery.Open strSQL,dcnDB 
if rsQuery.EOF  then 
       Response.Write "Record does not exists !" 	
else
    'elseif numCampusID = "2" then

        'strSQL = "Select *  from tblEmergencyContact where numCampusID ="&numCampusID
		'Set rsQuery = Server.CreateObject("ADODB.Recordset")
		'rsQuery.Open strSQL,dcnDB 		
		
    'end if
'*************************code to fill the combo******************************************
 strSQL = "Select * from tblCampus where numCampusID ="&numCampusID
		Set rsFillCombo = Server.CreateObject("ADODB.Recordset")
		rsFillCombo.Open strSQL,dcnDB 


   dim strCampusName 
   strCampusName = rsFillCombo("strCampusName")


'************************fill combo code ends here****************************************
'*************************filling up the form fields***************
'if numCampusID = 1 then %>
  <form method="POST" action="DeleteEC.asp?numCampusID=<%=numCampusID%>">
<%'else%>
   
<%'end if %>
		<div align="center">
		<table border="0" width="52%" id="table1">
			<tr>
				<td align="center" colspan="2">
<b><font face=Arial color="#FF0000">
				Delete&nbsp;Emergency Contact for <%= strCampusName%>  campus</font></b></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Name 1</font></td>
				<td align="center"><font face="Arial">
				<%=rsQuery("strEmergencyContactName1")%> </font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Position 1</font></td>
				<td align="center"><font face="Arial">
				<%=rsQuery("strEmergencyContactPosition1")%> </font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Phone 1</font></td>
				<td align="center"><font face="Arial">
				<%=rsQuery("strEmergencyContactPhone1")%> </font></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Name 2</font></td>
				<td align="center"><font face="Arial">
				<%=rsQuery("strEmergencyContactName2")%> </font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Position 2</font></td>
				<td align="center"><font face="Arial">
				<%=rsQuery("strEmergencyContactPosition2")%> </font></td>
			</tr>
			<tr>
				<td width="236" align="right"><font face="Arial">Emergency 
				Contact Phone 2</font></td>
				<td align="center"><font face="Arial">
				<%=rsQuery("strEmergencyContactPhone2")%> </font></td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center">&nbsp;</td>
			</tr>
			<tr>
				<td width="236" align="right">&nbsp;</td>
				<td align="center"><font face="Arial">
				<input type="submit" value="Delete Emergency Contact" name="btnEdit"></font></td>
			</tr>
		</table>
	</div>
</form>
 <%end if %>
</BODY>
</HTML>
