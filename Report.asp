<%@ Language = VBscript%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Quantities</title>
</head>

<body><%'--------------------------------- new updates from october 2006----------------------------------------------------------- %>
<%
dim dcnDB ' As ADODB.Connection
dim rsQueryC
dim rsQueryD
dim strSQL

Set dcnDB = Server.CreateObject("ADODB.Connection")
dcnDB.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
dcnDB.Open 

 Set rsQueryC = Server.CreateObject("ADODB.Recordset")
 rsQueryC.Open "Select * from tblLocation order by numLocationId",dcnDB  


while not rsQueryC.EOF  
%> <% 
dim cId
 cId = rsQueryC(0)
' Response.Write(cId) 
   '************************************************************************
  
set rsQueryD = Server.CreateObject("ADODB.Recordset")
	rsQueryD.Open "SELECT tblCampus.strCampusName,tblBuilding.strBuildingName,tblStoreLocation.strStoreLocation,tblChemicalContainer.strSpecificLocation,tblChemicalContainer.numQuantity"_
&" AS PG, tblChemicalContainer.strPG,tblChemicalContainer.strDangerousGoodsClass,tblLocation.BoolLicensedDepot,tblChemicalContainer.strContainerSize,tblLocation.numCampusID,tblStoreType.strStoreType,tblLocation.strStoreNotes,tblLocation.numLocationId"_
&" FROM ((((tblChemicalContainer RIGHT JOIN tblLocation"_
&" ON tblChemicalContainer.numLocationID = tblLocation.numLocationID)"_
&" INNER JOIN tblBuilding ON tblBuilding.numBuildingID = tblLocation.NumBuildingID) INNER JOIN"_
&" tblCampus ON tblCampus.numCampusID = tblLocation.NumCampusID) INNER JOIN tblStoreLocation ON "_
&" tblStoreLocation.numStoreLocationID = tblLocation.numStoreLocationID) INNER JOIN tblStoreType ON"_
&" tblStoreType.numStoreTypeID = tblLocation.numStoreTypeID "_
&" WHERE tblLocation.numStoreTypeId<>1 and tblLocation.numLocationId="& cId ,dcnDB

%>
 </table>
	 
	  <% 
	     	'*****************Applying the nested while loop for the required result
dim numNResult
dim strNNew
dim strNUnit

dim numNPI
dim numNAmountPI
dim numNTotalPI


dim numNPII
dim numNAmountPII	
dim numNTotalPII1

dim numNPIII
dim numNAmountPIII	
dim numNTotalPIII
dim strNDepoClass
dim boolNL
dim Nt
flgNI= 0
numNTotalPI= 0



 while not rsQueryD.EOF 
     %><% ' Response.Write(rsQueryD(12))%><% 
    if rsQueryD(12)= cId then
                 
	       			w = cstr(rsQueryD(0))' campus
					x=  cstr(rsQueryD(1))'building
					b = cstr(rsQueryD(2))'storeLocation
					p = (rsQueryD(10))'storeType
					z = (rsQueryD(11))'storeNotes
				    
				    
              	strNLocation = w +", " + x +", " + b + ", " + p + ", " + z
                    	
                strNDepoClass = (rsQueryD(6))	
     
                boolNL= rsQueryD(7)
                 
  	              if  rsQueryD(5) = "I" then
	                 
	                	numNResult =instr(1,rsQueryD(8)," ",vbTextCompare) 
 						strNnew = mid(rsQueryD(8),1,numNResult)  
						numNPI = cdbl(strNnew)
						
						strNUnit = Mid(rsQueryD(8), numNresult, Len(rsQueryD(8)))
						  
						       
							if  strNUnit = " mL" then
								numNPI = (numNPI/ 1000) 
							elseif strNUnit =  " g" then
							    
							    numNPI = (numNPI/ 1000)
							elseif strNUnit =  " L" then
							    numNPI = numNPI   
							elseif strNUnit =  " Kg" then
							    numNPI = numNPI      
							end if 
							
                      numNAmountPI = (numNPI * rsQueryD(4))	     
 	                 numNTotalpI = numNTotalpI + numNAmountPI
                end if
                
                if  rsQueryD(5) = "II" then
	                	numNResult =instr(1,rsQueryD(8)," ",vbTextCompare) 
 						strNnew = mid(rsQueryD(8),1,numNResult)  
						numNPII = cdbl(strNnew)
						strNUnit = Mid(rsQueryD(8), numNresult, Len(rsQueryD(8)))
						    
						       
							if  strNUnit = " mL" then
								numNPII = (numNPII/ 1000) 
							elseif strNUnit =  " g" then
							    numPII = (numNPII/ 1000)
							elseif strNUnit =  " L" then
							    numNPII = numNPII   
							elseif strNUnit =  " Kg" then
							    numNPII = numNPII      
							end if 
							
                      numNAmountPII = (numNPII * rsQueryD(4))	     
 	                  numNTotalpII = numNTotalpII + numNAmountPII
                end if
                if  rsQueryD(5) = "III" then
	               	numNResult =instr(1,rsQueryD(8)," ",vbTextCompare) 
 						strNnew = mid(rsQueryD(8),1,numNResult)  
						numNPIII = cdbl(strNnew)
						strNUnit = Mid(rsQueryD(8), numNresult, Len(rsQueryD(8)))
						 
						  
							if  strNUnit = " mL" then
								numNPIII = (numNPIII/ 1000)
							elseif strNUnit =  " g" then
							    numNPIII = (numNPIII/ 1000)
							elseif strNUnit =  " L" then
							    numNPIII = numNPIII  
							elseif strNUnit =  " Kg" then
							    numNPIII = numNPIII  
							end if 
							
                     numNAmountPIII = (numNPIII * rsQueryD(4))	     
 	                numNTotalpIII = numNTotalpIII + numNAmountPIII
                end if
    end if
          
           rsQueryD.MoveNext 
           ' n = n + 1
         wend
             %><%'Response.Write(n) 
			if rsQueryD.EOF = False then 
				 rsQueryD.MoveFirst 
			end if
		 Nt = numNTotalPI + numNTotalPII + numNTotalPIII
	    
	    if Nt >0 then 
      %>

  Store Name&nbsp;&nbsp;&nbsp; :&nbsp;&nbsp;&nbsp; <%=strNLocation %>
  
   <br>
	  <TABLE id=table5 style="WIDTH: 846px; HEIGHT: 52px" width=846 border=1><!-- MSTableType="nolayout" -->
  
  <TR>
    <TD align=middle width="35%" bgColor=#ffff00 height="23">&nbsp;</TD>
    <TD align=middle width="6%" bgColor=#ffff00 height="23">&nbsp;</TD>
    <TD align=middle width="11%" bgColor=#ffff00 height="23">&nbsp;</TD>
    <TD align=middle width="37%" bgColor=#ffff00 colspan="4" height="23">Quantities ( L / Kg 
	)</TD>
    </TR>
  <TR>
    <TD align=middle width="40%" bgColor=#ffff00>DG Store Name and Location</TD>
   
    <TD align=middle width="10%" bgColor=#ffff00>PG I</TD>
    <TD align=middle width="10%" bgColor=#ffff00>PG II</TD>
    <TD align=middle width="10%" bgColor=#ffff00>PG III</TD>
    <TD align=middle width="30%" bgColor=#ffff00>Total</TD>
    </TR>
	<TR>
  
	<TR>
   <TD align=middle width="40%" bgColor=#FFFFFF height="8"><%= strNDepoClass%></TD>
<%   
    '----------------- check if total value is > than 0 ----then only display-------------------
	
	 if numNTotalPI = 0 then %><TD align=middle width="10%" bgColor=#FFFFFF height="23">0</TD><%           
      else %> <TD align=middle width="11%" bgColor=#FFFFFF height="23"> <%=round((numNTotalPI),1)%></TD><% numNTotalPI = 0        
      end if 
      
      if numNTotalPII = 0 then %><TD align=middle width="10%" bgColor=#FFFFFF height="8">0</TD><%           
      else %> <TD align=middle width="11%" bgColor=#FFFFFF height="23"> <%=round((numNTotalPII),1)%></TD><% numNTotalPII = 0        
      end if 
      
      if numNTotalPIII = 0 then %><TD align=middle width="10%" bgColor=#FFFFFF height="8">0</TD><%           
      else %> <TD align=middle width="7%" bgColor=#FFFFFF height="23"> <%=round((numNTotalPIII),1)%></TD><% numNTotalPIII = 0        
      end if  
    
      %>
      
     <% if NT = 0 then %><TD align=middle width="30%" bgColor=#FFFFFF height="8">0</TD></TR><%           
      else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=round(NT,1) %></TD></TR><% T = 0        
      end if            
     end if    %>
  <%   	
		rsQueryC.MoveNext  
       	wend   
   	  	  
	'***********************************************************************
 %>

</body>

</html>