<%

'********************************DATABASE CONNECTIVITY CODE ********************************


'******************************************new Code**************************************************
dim dcnDB1 ' As ADODB.Connection

dim strSQL1

Set dcnDB1 = Server.CreateObject("ADODB.Connection")
dcnDB1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
dcnDB1.Open


strSQL1 = "SELECT tblChemicalContainer.numQuantity AS PG, tblChemicalContainer.strPG AS strPG, "_
&"tblChemicalContainer.strDangerousGoodsClass AS strDangertousGoodsClass, tblChemicalContainer.strContainerSize AS strContainerSize"_
&" FROM tblChemicalContainer "

set rsQueryB = Server.CreateObject("ADODB.Recordset")
rsQueryB.Open strSQL1,dcnDB1

%>








<%
'*****************Applying the nested while loop for the required result


f1 = false
f2 = false


do while not rsQueryB.EOF


strDGC = rsQueryB("strDangertousGoodsClass")
strDGC = mid(strDGC,1,1)
if strDGC="" then
strDGC="E"
end if
Select Case strDGC

'****************************************CASE 1****************************************************************
'**************************************************************************************************************
Case "1":    	              ' DG class 1
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else
if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then
numPI = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))
end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI= (numPI/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI = (numPI/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI = (numPI/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI = numPI
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" then
numPI = numPI
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI = (numPI * rsQueryB("PG"))
numTotalpI = numTotalpI + numAmountPI
else
numAmountPI = (numPI * 0)
numTotalpI = numTotalpI + numAmountPI
end if
end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then
numPII = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII = (numPII/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII = (numPII/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII = (numPII/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPII = numPII
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" then
numPII = numPII
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPII = (numPII * rsQueryB("PG"))
numTotalpII = numTotalpII + numAmountPII
else
numAmountPII = (numPII * 0)
numTotalpII = numTotalpII + numAmountPII
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then
numPIII = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII = (numPIII/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII = (numPIII/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII = (numPIII/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIII = numPIII
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII = numPIII
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPIII = (numPIII * rsQueryB("PG"))
numTotalpIII = numTotalpIII + numAmountPIII
else
numAmountPIII = (numPIII * 0)
numTotalpIII = numTotalpIII + numAmountPIII
end if

end if
t = numTotalPI + numTotalPII + numTotalPIII
if t <> 0 then
f1 = true
else
f1 = false
end if

'****************************************END OF CASE 1 ********************************************************
'**************************************************************************************************************

'****************************************CASE 2****************************************************************
'**************************************************************************************************************
Case "2":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI2 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else
if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPI2 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI2= (numPI2/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI2 = (numPI2/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI2 = (numPI2/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI2 = numPI2
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI2 = numPI2
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI2 = (numPI2 * rsQueryB("PG"))
numTotalpI2 = numTotalpI2 + numAmountPI2
else
numAmountPI = (numPI * 0)
numTotalpI2 = numTotalpI2 + numAmountPI2
end if
end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII2 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII2 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII2 = (numPII2/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII2 = (numPII2/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII2 = (numPII2/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPII2 = numPII2
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII2 = numPII2
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPII2 = (numPII2 * rsQueryB("PG"))
numTotalpII2 = numTotalpII2 + numAmountPII2
else
numAmountPII = (numPII * 0)
numTotalpII2 = numTotalpII2 + numAmountPII2
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII2 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then


numPIII2 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII2 = (numPIII2/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII2 = (numPIII2/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII2 = (numPIII2/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIII2 = numPIII2
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII2 = numPIII2
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPIII2 = (numPI2 * rsQueryB("PG"))
numTotalpIII2 = numTotalpIII2 + numAmountPIII2
else
numAmountPIII = (numPIII * 0)
numTotalpIII2 = numTotalpIII2 + numAmountPIII2
end if

end if
t2 = numTotalPI2 + numTotalPII2 + numTotalPIII2
if t2 <> 0 then
f2 = true
else
f2 = false
end if

'****************************************END OF CASE 2*******************************************************
'************************************************************************************************************

'****************************************CASE 3***  ( DG CLASS 3)  *************************************************************
'**************************************************************************************************************
' Used trim() function in case 3 to see if it had any effect in fixing problem with calc. - NO
' put in mg factoring to see if this fixes result
Case "3":
if  rsQueryB("strPG") = "I" then
strContSize = Trim(rsQueryB("strContainerSize"))
numResult =instr(1,strContSize," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(strContSize)>0 then
strnew = mid(strContSize,1,numResult)
if IsNumeric(strnew) then
numPI3 = cdbl(strnew)
strUnit = Mid(strContSize, numresult, Len(strContSize))
end if
end if

else
if not isnull(strContSize) and len(strContSize)>0 then
strstr = strContSize
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then
numPI3 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))
end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI3= (numPI3/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI3 = (numPI3/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI3 = (numPI3/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI3 = numPI3
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI3 = numPI3
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI3 = (numPI3 * rsQueryB("PG"))
numTotalpI3 = numTotalpI3 + numAmountPI3
else
numAmountPI3 = (numPI3 * 0)
numTotalpI3 = numTotalpI3 + numAmountPI3
end if
end if

if  rsQueryB("strPG") = "II" then
strContSize = Trim(rsQueryB("strContainerSize"))
numResult =instr(1,strContSize," ",vbTextCompare)

if numResult > 0 then

if strContSize<>" " and len(strContSize)>0 then
strnew = mid(strContSize,1,numResult)
if IsNumeric(strnew) then
numPII3 = cdbl(strnew)
strUnit = Mid(strContSize, numresult, Len(strContSize))
end if
end if

else

if not isnull(strContSize) and len(strContSize)>0 then

strstr = strContSize
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII3 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII3 = (numPII3/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII3 = (numPII3/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII3 = (numPII3/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPII3 = numPII3
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII3 = numPII3
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPII3 = (numPII3 * rsQueryB("PG"))
numTotalpII3 = numTotalpII3 + numAmountPII3
else
numAmountPII3 = (numPII3 * 0)
numTotalpII3 = numTotalpII3 + numAmountPII3
end if
end if

if  rsQueryB("strPG") = "III" then
strContSize = Trim(rsQueryB("strContainerSize"))
numResult =instr(1,strContSize," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(strContSize)>0 then
strnew = mid(strContSize,1,numResult)
if IsNumeric(strnew) then
numPIII3 = cdbl(strnew)
strUnit = Mid(strContSize, numresult, Len(strContSize))
end if
end if

else

if not isnull(strContSize) and len(strContSize)>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then


numPIII3 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII3 = (numPIII3/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII3 = (numPIII3/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII3 = (numPIII3/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPIII3 = numPIII3
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII3 = numPIII3
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIII3 = (numPIII3 * rsQueryB("PG"))
numTotalpIII3 = numTotalpIII3 + numAmountPIII3
else
numAmountPIII3 = (numPI3 * 0)
numTotalpIII3 = numTotalpIII3 + numAmountPIII3
end if

end if
t3 = numTotalPI3 + numTotalPII3 + numTotalPIII3
if t3 <> 0 then
f3 = true
else
f3= false
end if

'****************************************END OF CASE 3*******************************************************
'************************************************************************************************************

'****************************************CASE 4****************************************************************
'**************************************************************************************************************
Case "4":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI4 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else
if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPI4 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI4= (numPI4/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI4 = (numPI4/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI4 = (numPI4/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI4 = numPI4
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then 'NOTE  KG update all others ***************
numPI4 = numPI4
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI4 = (numPI4 * rsQueryB("PG"))
numTotalpI4 = numTotalpI4 + numAmountPI4
else
numAmountPI4 = (numPI4 * 0)
numTotalpI4 = numTotalpI4 + numAmountPI4
end if
end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII4 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII4 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII4 = (numPII4/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII4 = (numPII4/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII4 = (numPII4/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPII4 = numPII4
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII4 = numPII4
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPII4 = (numPII4 * rsQueryB("PG"))
numTotalpII4 = numTotalpII4 + numAmountPII4
else
numAmountPII4 = (numPII4 * 0)
numTotalpII4 = numTotalpII4 + numAmountPII4
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII4 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then


numPIII4 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII4 = (numPIII4/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII4 = (numPIII4/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII4 = (numPIII4/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIII4 = numPIII4
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII4 = numPIII4
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIII4 = (numPIII4 * rsQueryB("PG"))
numTotalpIII4 = numTotalpIII4 + numAmountPIII4
else
numAmountPIII4 = (numPIII4 * 0)
numTotalpIII4 = numTotalpIII4 + numAmountPIII4
end if

end if
t4 = numTotalPI4 + numTotalPII4 + numTotalPIII4
if t4 <> 0 then
f4 = true
else
f4 = false
end if

'****************************************END OF CASE 4*******************************************************
'************************************************************************************************************
'****************************************CASE 5****************************************************************
'**************************************************************************************************************
Case "5":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI5 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else
if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPI5 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI5= (numPI5/ 1000)

elseif strUnit =  " g" or strUnit = " G" then
numPI5 = (numPI5/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI5 = (numPI5/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI5 = numPI5
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI5 = numPI5
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI5 = (numPI5 * rsQueryB("PG"))
numTotalpI5 = numTotalpI5 + numAmountPI5
else
numAmountPI5 = (numPI5 * 0)
numTotalpI5 = numTotalpI5 + numAmountPI5
end if
end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII5 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then
numPII5 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII5 = (numPII5/ 1000)

elseif strUnit =  " g" or strUnit = " G" then
numPII5 = (numPII5/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII5 = (numPII5/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPII5 = numPII5
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII5 = numPII5
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPII5 = (numPII5 * rsQueryB("PG"))
numTotalpII5 = numTotalpII5 + numAmountPII5
else
numAmountPII5 = (numPII5 * 0)
numTotalpII5 = numTotalpII5 + numAmountPII5
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII5 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then


numPIII5 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII5 = (numPIII5/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII5 = (numPIII5/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII5 = (numPIII5/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPIII5 = numPIII5
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII5 = numPIII5
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIII5 = (numPIII5 * rsQueryB("PG"))
numTotalpIII5 = numTotalpIII5 + numAmountPIII5
else
numAmountPIII5 = (numPIII5 * 0)
numTotalpIII5 = numTotalpIII5 + numAmountPIII5
end if

end if
t5 = numTotalPI5 + numTotalPII5 + numTotalPIII5
if t5 <> 0 then
f5 = true
else
f5 = false
end if

'****************************************END OF CASE 5*******************************************************
'************************************************************************************************************
'****************************************CASE 6****************************************************************
'**************************************************************************************************************
Case "6":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI6 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else
if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPI6 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI6= (numPI6/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI6 = (numPI6/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI6 = (numPI6/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI6 = numPI6
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI6 = numPI6
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI6 = (numPI6 * rsQueryB("PG"))
numTotalpI6 = numTotalpI6 + numAmountPI6
else
numAmountPI6 = (numPI6 * 0)
numTotalpI6 = numTotalpI6 + numAmountPI6
end if
end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII6 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII6 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII6 = (numPII6/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII6 = (numPII6/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII6 = (numPII6/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPII6 = numPII6
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII6 = numPII6
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPII6 = (numPII6 * rsQueryB("PG"))
numTotalpII6 = numTotalpII6 + numAmountPII6
else
numAmountPII6 = (numPII6 * 0)
numTotalpI6 = numTotalpI6 + numAmountPI6
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII6 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then


numPIII6 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII6 = (numPIII6/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII6 = (numPIII6/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII6 = (numPIII6/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPIII6 = numPIII6
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII6 = numPIII6
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPIII6 = (numPIII6 * rsQueryB("PG"))
numTotalpIII6 = numTotalpIII6 + numAmountPIII6
else
numAmountPIII6 = (numPIII6 * 0)
numTotalpIII6 = numTotalpIII6 + numAmountPIII6
end if

end if
t6 = numTotalPI6 + numTotalPII6 + numTotalPIII6
if t6 <> 0 then
f6 = true
else
f6 = false
end if

'****************************************END OF CASE 6*******************************************************
'************************************************************************************************************
'****************************************CASE 7****************************************************************
'**************************************************************************************************************
Case "7":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI7 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else
if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPI7 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI7= (numPI7/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI7 = (numPI7/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI7 = (numPI7/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI7 = numPI7
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI7 = numPI7
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPI7 = (numPI7 * rsQueryB("PG"))
numTotalpI7 = numTotalpI7 + numAmountPI7
else
numAmountPI7 = (numPI7 * 0)
numTotalpI7 = numTotalpI7 + numAmountPI7
end if
end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII7 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII7 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII7 = (numPII7/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII7 = (numPII7/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII7 = (numPII7/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPII7 = numPII7
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII7 = numPII7
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPII7 = (numPII7 * rsQueryB("PG"))
numTotalpII7 = numTotalpII7 + numAmountPII7
else
numAmountPII7 = (numPII7 * 0)
numTotalpII7 = numTotalpII7 + numAmountPII7
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII7 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then


numPIII7 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII7 = (numPIII7/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII7 = (numPIII7/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII7 = (numPIII7/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPIII7 = numPIII7
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII7 = numPIII7
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPIII7 = (numPIII7 * rsQueryB("PG"))
numTotalpIII7 = numTotalpIII7 + numAmountPIII7
else
numAmountPIII7 = (numPIII7 * 0)
numTotalpIII7 = numTotalpIII7 + numAmountPIII7
end if

end if
t7 = numTotalPI7 + numTotalPII7 + numTotalPIII7
if t7 <> 0 then
f7 = true
else
f7 = false
end if

'****************************************END OF CASE 7*******************************************************
'************************************************************************************************************
'****************************************CASE 8**************************************************************
'************************************************************************************************************

Case "8":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI8 = cdbl(strnew)

strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))

end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)

strnew = mid(strstr,1,numResult)

if IsNumeric(strnew) then

numPI8 = cdbl(strnew)

strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI8= (numPI8/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI8 = (numPI8/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI8 = (numPI8/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI8 = numPI8
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI8 = numPI8
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPI8 = (numPI8 * rsQueryB("PG"))
numTotalpI8 = numTotalpI8 + numAmountPI8
else
numAmountPI8 = (numPI8 * 0)
numTotalpI8 = numTotalpI8 + numAmountPI8
end if

end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII8 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
'Response.Write strstr
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII8 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII8 = (numPII8/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII8 = (numPII8/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII8 = (numPII8/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPII8 = numPII8
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII8 = numPII8
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPII8 = (numPII8 * rsQueryB("PG"))
numTotalpII8 = numTotalpII8 + numAmountPII8
else
numAmountPII8 = (numPII8 * 0)
numTotalpII8 = numTotalpII8 + numAmountPII8
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII8 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")

strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPIII8 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml"  or strUnit =  " ML" then
numPIII8 = (numPIII8/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII8 = (numPIII8/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII8 = (numPIII8/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIII8 = numPIII8
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII8 = numPIII8
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIII8 = (numPIII8 * rsQueryB("PG"))
numTotalpIII8 = numTotalpIII8 + numAmountPIII8
else
numAmountPIII8 = (numPIII8 * 0)
numTotalpIII8 = numTotalpIII8 + numAmountPIII8
end if

end if
t8 = numTotalPI8 + numTotalPII8 + numTotalPIII8
if t8 <> 0 then
f8 = true
else
f8= false
end if

'************************************ END OF CASE 8******************************************************************
'************************************************************************************************************
'****************************************CASE 9**************************************************************
'************************************************************************************************************

Case "9":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPI9 = cdbl(strnew)

strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))

end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)

strnew = mid(strstr,1,numResult)

if IsNumeric(strnew) then

numPI9 = cdbl(strnew)

strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPI9= (numPI9/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPI9 = (numPI9/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPI9 = (numPI9/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPI9 = numPI9
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPI9 = numPI9
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPI9 = (numPI9 * rsQueryB("PG"))
numTotalpI9 = numTotalpI9 + numAmountPI9
else
numAmountPI9 = (numPI9 * 0)
numTotalpI9 = numTotalpI9 + numAmountPI9
end if

end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPII9 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
'Response.Write strstr
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPII9 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPII9 = (numPII9/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPII9 = (numPII9/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPII9 = (numPII9/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPII9 = numPII9
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPII9 = numPII9
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPII9 = (numPII9 * rsQueryB("PG"))
numTotalpII9 = numTotalpII9 + numAmountPII9
else
numAmountPII9 = (numPII9 * 0)
numTotalpII9 = numTotalpII9 + numAmountPII9
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIII9 = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")

strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPIII9 = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIII9 = (numPIII9/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIII9 = (numPIII9/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIII9 = (numPIII9/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIII9 = numPIII9
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIII9 = numPIII9
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIII9 = (numPIII9 * rsQueryB("PG"))
numTotalpIII9 = numTotalpIII9 + numAmountPIII9
else
numAmountPIII9 = (numPIII9 * 0)
numTotalpIII9 = numTotalpIII9 + numAmountPIII9
end if

end if
t9 = numTotalPI9 + numTotalPII9 + numTotalPIII9
if t9 <> 0 then
f9 = true
else
f9= false
end if
'************************************ END OF CASE 9******************************************************************
'************************************************************************************************************
'****************************************CASE N**************************************************************
'************************************************************************************************************

Case "N":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIN = cdbl(strnew)

strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))

end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)

strnew = mid(strstr,1,numResult)

if IsNumeric(strnew) then

numPIN = cdbl(strnew)

strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIN= (numPIN/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIN = (numPIN/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIN = (numPIN/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPIN = numPIN
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIN = numPIN
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPIN = (numPIN * rsQueryB("PG"))
numTotalpIN = numTotalpIN + numAmountPIN
else
numAmountPIN = (numPI8 * 0)
numTotalpIN = numTotalpIN + numAmountPIN
end if

end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIIN = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
'Response.Write strstr
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPIIN = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIIN = (numPIIN/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIIN = (numPIIN/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIIN = (numPIIN/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIIN = numPIIN
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIIN = numPIIN
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIIN = (numPIIN * rsQueryB("PG"))
numTotalpIIN = numTotalpIIN + numAmountPIIN
else
numAmountPIIN = (numPIIN * 0)
numTotalpIIN = numTotalpIIN + numAmountPIIN
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIIIN = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")

strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPIIIN = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml"  or strUnit =  " ML" then
numPIIIN = (numPIIIN/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIIIN = (numPIIIN/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIIIN = (numPIIIN/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIIIN = numPIIIN
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIIIN = numPIIIN
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIIIN = (numPIIIN * rsQueryB("PG"))
numTotalpIIIN = numTotalpIIIN + numAmountPIIIN
else
numAmountPIIIN = (numPIII8 * 0)
numTotalpIIIN = numTotalpIIIN + numAmountPIIIN
end if

end if
tN = numTotalPIN + numTotalPIIN + numTotalPIIIN
if tN <> 0 then
fN = true
else
fN= false
end if

'************************************ END OF CASE N******************************************************************
'************************************************************************************************************
'****************************************CASE E**************************************************************
'************************************************************************************************************

Case "E":
if  rsQueryB("strPG") = "I" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIE = cdbl(strnew)

strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))

end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)

strnew = mid(strstr,1,numResult)

if IsNumeric(strnew) then

numPIE = cdbl(strnew)

strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIE= (numPIE/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIE = (numPIE/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIE = (numPIE/ 1000000)
elseif strUnit =  " L" or strUnit = " l" then
numPIE = numPIE
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIE = numPIE
end if

if IsNumeric(rsQueryB("PG"))	then
numAmountPIE = (numPIE * rsQueryB("PG"))
numTotalpIE = numTotalpIE + numAmountPIE
else
numAmountPIE = (numPIE * 0)
numTotalpIE = numTotalpIE + numAmountPIE
end if

end if

if  rsQueryB("strPG") = "II" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult > 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIIE = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")
strstr = putSpace(strstr)
'Response.Write strstr
numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPIIE = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml" or strUnit = " ML" then
numPIIE = (numPIIE/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIIE = (numPIIE/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIIE = (numPIIE/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIIE = numPIIE
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIIE = numPIIE
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIIE = (numPIIE * rsQueryB("PG"))
numTotalpIIE = numTotalpIIE + numAmountPIIE
else
numAmountPIIE = (numPIIE * 0)
numTotalpIIE = numTotalpIIE + numAmountPIIE
end if
end if

if  rsQueryB("strPG") = "III" then

numResult =instr(1,rsQueryB("strContainerSize")," ",vbTextCompare)

if numResult> 0 then

if rsQueryB("strContainerSize")<>" " and len(rsQueryB("strContainerSize"))>0 then
strnew = mid(rsQueryB("strContainerSize"),1,numResult)
if IsNumeric(strnew) then
numPIIIE = cdbl(strnew)
strUnit = Mid(rsQueryB("strContainerSize"), numresult, Len(rsQueryB("strContainerSize")))
end if
end if

else

if not isnull(rsQueryB("strContainerSize")) and len(rsQueryB("strContainerSize"))>0 then

strstr = rsQueryB("strContainerSize")

strstr = putSpace(strstr)

numResult =instr(1,strstr," ",vbTextCompare)
strnew = mid(strstr,1,numResult)
if IsNumeric(strnew) then

numPIIIE = cdbl(strnew)
strUnit = Mid(strstr, numresult, Len(strstr))

end if
end if
end if

if  strUnit = " mL" or strUnit = " ml"  or strUnit =  " ML" then
numPIIIE = (numPIIIE/ 1000)
elseif strUnit =  " g" or strUnit = " G" then
numPIIIE = (numPIIIE/ 1000)
ElseIf strUnit = " mg" Or strUnit = " MG" Then
numPIIIE = (numPIIIE/ 1000000)

elseif strUnit =  " L" or strUnit = " l" then
numPIIIE = numPIIIE
elseif strUnit =  " Kg" or strUnit = " kg" Or strUnit = " KG" Then
numPIIIE = numPIIIE
end if
if IsNumeric(rsQueryB("PG"))	then
numAmountPIIIE = (numPIIIE * rsQueryB("PG"))
numTotalpIIIE = numTotalpIIIE + numAmountPIII8
else
numAmountPIIIE = (numPIIIE * 0)
numTotalpIIIE = numTotalpIIIE + numAmountPIII8
end if

end if
tE = numTotalPIE + numTotalPIIE + numTotalPIIIE
if tE <> 0 then
fE = true
else
fE= false
end if

'************************************ END OF CASE E******************************************************************

end select




rsQueryB.MoveNext
loop
rsQueryB.MoveFirst


%>


<%if f1 = true or f2 = true or f3 = true or f4 = true  or f5 = true or f6 = true or f7 = true or f8 = true or f9 = true or fN=true or fE=true then%>
<BR>
<TABLE id=table5 style="WIDTH: 846px; HEIGHT: 105px" width=846 border=1>
    <TR>
        <TD align=middle width="25%" bgColor=#ffff00>&nbsp;</TD>
        <TD align=middle width="25%" bgColor=#ffff00>&nbsp;</TD>
        <TD align=middle width="25%" bgColor=#ffff00>&nbsp;</TD>
        <TD align=middle width="25%" bgColor=#ffff00 colspan="4">Quantities ( L / Kg )</TD>
    </TR>
    <TR>
        <!-- check alignment of category with packing group -->
        <TD align=middle width="25%" bgColor=#FFFF00>DG Class</TD>
        <TD align=middle width="25%" bgColor=#ffff00>Category 1</TD>
        <TD align=middle width="25%" bgColor=#ffff00>Category 2</TD>
        <TD align=middle width="25%" bgColor=#ffff00>Category 3</TD>
        <TD align=middle width="25%" bgColor=#ffff00>Total</TD>
    </TR>

    <%
    f1 = false
    f2 = false
    f3 = false
    f4 = false
    f5 = false
    f6 = false
    f7 = false
    f8 = false
    f9 = false
    fN = false
    fE = false
    end if%>

    <% if t <>0 then %>

    <%'----------------------------First ROW ------------------------------------------------------------------------%>
    <TD align=middle width="25%" bgColor=#FFFFFF>Class 1</TD>
    <%if numTotalPI = 0 or numTotalPI = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI),2)%></TD><% numTotalPI = 0
    end if

    if numTotalPII = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII),2)%></TD><% numTotalPII = 0
    end if

    if numTotalPIII = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII),2)%></TD><% numTotalPIII = 0
    end if
    if T = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=round((T),2) %></TD></TR><% T = 0
    end if
    end if
    '----------------------------SECOND ROW ------------------------------------------------------------------------

    if t2 <>0 then%>

    <TD align=middle width="25%" bgColor=#FFFFFF>Class 2</TD>
    <%if numTotalPI2 = 0 or numTotalPI2 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI2),2)%></TD><% numTotalPI2 = 0
    end if

    if numTotalPII2 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII2),2)%></TD><% numTotalPII2 = 0
    end if

    if numTotalPIII2 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII2),2)%></TD><% numTotalPIII2 = 0
    end if
    if T2 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T2),2) %></TD></TR><% T2 = 0
    end if
    end if

    '----------------------------THIRD ROW ------------------------------------------------------------------------

    if t3 <>0 then%>

    <TD align=middle width="25%" bgColor=#FFFFFF>Class 3</TD>
    <%if numTotalPI3 = 0 or numTotalPI3 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI3),2)%></TD><% numTotalPI3 = 0
    end if

    if numTotalPII3 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII3),2)%></TD><% numTotalPII3 = 0
    end if

    if numTotalPIII3 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII3),2)%></TD><% numTotalPIII3 = 0
    end if
    if T3 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T3),2) %></TD></TR><% T3 = 0
    end if
    end if
    '----------------------------FOURTH ROW ------------------------------------------------------------------------

    if t4 <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class 4</TD>
    <%if numTotalPI4 = 0 or numTotalPI4 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI4),2)%></TD><% numTotalPI4 = 0
    end if

    if numTotalPII4 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII4),2)%></TD><% numTotalPII4 = 0
    end if

    if numTotalPIII4 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII4),2)%></TD><% numTotalPIII4 = 0
    end if
    if T4 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T4),2) %></TD></TR><% T4 = 0
    end if
    end if
    '----------------------------FIFTH ROW ------------------------------------------------------------------------

    if t5 <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class 5</TD>
    <%if numTotalPI5 = 0 or numTotalPI5= "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI5),2)%></TD><% numTotalPI5 = 0
    end if

    if numTotalPII5 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII5),2)%></TD><% numTotalPII5 = 0
    end if

    if numTotalPIII5 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII5),2)%></TD><% numTotalPIII5 = 0
    end if
    if T5 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T5),2) %></TD></TR><% T5 = 0
    end if
    end if
    '----------------------------SIXTH ROW ------------------------------------------------------------------------

    if t6 <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class 6</TD>
    <%if numTotalPI6 = 0 or numTotalPI6 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI6),2)%></TD><% numTotalPI6 = 0
    end if

    if numTotalPII6 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII6),2)%></TD><% numTotalPII6 = 0
    end if

    if numTotalPIII6 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII6),2)%></TD><% numTotalPIII6 = 0
    end if
    if T6 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T6),2) %></TD></TR><% T6 = 0
    end if
    end if
    '----------------------------SEVENTH ROW ------------------------------------------------------------------------

    if t7 <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class 7</TD>
    <%if numTotalPI7 = 0 or numTotalPI7 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI7),2)%></TD><% numTotalPI7 = 0
    end if

    if numTotalPII7 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII7),2)%></TD><% numTotalPII7 = 0
    end if

    if numTotalPIII7 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII7),2)%></TD><% numTotalPIII7 = 0
    end if
    if T7 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T7),2) %></TD></TR><% T7 = 0
    end if
    end if
    '----------------------------EIGHT ROW ------------------------------------------------------------------------

    if t8 <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class 8</TD>
    <%if numTotalPI8 = 0 or numTotalPI8 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI8),2)%></TD><% numTotalPI8 = 0
    end if

    if numTotalPII8 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII8),2)%></TD><% numTotalPII8 = 0
    end if

    if numTotalPIII8 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII8),2)%></TD><% numTotalPIII8 = 0
    end if
    if T8 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T8),2) %></TD></TR><% T8 = 0
    end if
    end if
    '----------------------------NINE ROW ------------------------------------------------------------------------

    if t9 <>0 then%>

    <TD align=middle width="25%" bgColor=#FFFFFF>Class 9</TD>
    <%if numTotalPI9 = 0 or numTotalPI9 = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPI9),2)%></TD><% numTotalPI9 = 0
    end if

    if numTotalPII9 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPII9),2)%></TD><% numTotalPII9 = 0
    end if

    if numTotalPIII9 = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIII9),2)%></TD><% numTotalPIII9 = 0
    end if
    if T9 = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((T9),2) %></TD></TR><% T9 = 0
    end if
    end if

    '----------------------------NONE ROW ------------------------------------------------------------------------

    if tN <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class NONE</TD>
    <%if numTotalPIN = 0 or numTotalPIN = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIN),2)%></TD><% numTotalPIN = 0
    end if

    if numTotalPIIN = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIIN),2)%></TD><% numTotalPIIN = 0
    end if

    if numTotalPIIIN = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=round((numTotalPIIIN),2)%></TD><% numTotalPIIIN = 0
    end if
    if TN = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=Round((TN),2) %></TD></TR><% TN = 0
    end if
    end if

    '----------------------------Empty ROW ------------------------------------------------------------------------

    if tE <>0 then%>


    <TD align=middle width="25%" bgColor=#FFFFFF>Class Empty</TD>
    <%if numTotalPIE = 0 or numTotalPIE = "" then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=((numTotalPIE))%></TD><% numTotalPIE = 0
    end if

    if numTotalPIIE = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=((numTotalPIIE))%></TD><% numTotalPIIE = 0
    end if

    if numTotalPIIIE = 0 then %><TD align=middle width="7%" bgColor=#FFFFFF>0</TD><%
    else %> <TD align=middle width="25%" bgColor=#FFFFFF> <%=((numTotalPIIIE))%></TD><% numTotalPIIIE = 0
    end if
    if TE = 0 then %><TD align=middle width="25%" bgColor=#FFFFFF>0</TD></TR><%
    else %> <TD align=middle width="7%" bgColor=#FFFFFF> <%=(TE) %></TD></TR><% TE = 0
    end if
    end if


    %>






</TABLE>




<%

function putSpacetemp(strString)
str = strString
str2 = strString

cnt = Len(str)
cnt2 = 1
While cnt > 0
strnew = Mid(str, 1, cnt)
str = strnew
cnt = cnt - 1

If IsNumeric(str) Then
quantity = str
str = quantity + " " + unit
putSpace = str
Exit function

Else
While cnt2 <= Len(str2)
If cnt2 > 1 Then
strnew = Mid(str2, cnt2 - 1, Len(str2))
Else
strnew = Mid(str2, cnt2, Len(str2))
End If
str2 = strnew
cnt2 = cnt2 + 1
Wend
If Not IsNumeric(str2) Then
If Len(str2) = 2 Then
str2 = Mid(str2, 2, Len(str2))
unit = str2
ElseIf Len(str2) = 3 Then
str2 = Mid(str2, 3, Len(str2))
unit = str2
end if
unit = str2
end if
end if

wend
end function


function putSpace(str)
Dim a(7)
Dim b(7)
Dim unit
Dim qty
Dim cnt
Dim cntlen
Dim i
Dim j
Dim k
Dim flg
i = 0
j = 0
k = 0

flg = False
cntlen = 1
cnt = Len(str)
While cnt >= cntlen
strnew = Mid(str, 1, cntlen)
cntlen = cntlen + 1
If IsNumeric(strnew) Then
a(i) = strnew
i = i + 1
Else
If flg = False Then
stru = Mid(str, cntlen - 1, Len(str))
b(j) = stru
j = j + 1
flg = True
End If
End If

Wend
k = i
' printing the 2 arrays
For i = 0 To k - 1
strqty = a(i)
Next



strUnit =  b(0)


s = cstr(strQty)+" "+cstr(strUnit)
strQTY = ""
strUnit =""
putSpace = s



end function
%>
