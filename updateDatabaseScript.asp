<!--#INCLUDE FILE="DbConfig.asp"-->
<%

dim text,size, unit, length, i, spaceCheck, sql
set conn=server.createobject("ADODB.Connection")
conn.open constr
set rs=server.createobject("ADODB.Recordset")
sql="select numChemicalContainerID, strContainerSize, numQuantity from tblChemicalContainer"
rs.open sql, conn

do until rs.eof

spaceCheck=instr(1, rs.fields.item("strContainerSize")," ")

length=len(rs.fields.item("strContainerSize"))

'it would be good if this script could trim the leading spaces from a strContainerSize
if spaceCheck=0 then
updateSQL="update tblChemicalContainer set strContainerSize='" & putSpace(length,rs.fields.item("strContainerSize")) & "' where numChemicalContainerID=" & rs.fields.item("numChemicalContainerID")
conn.execute updateSQL
size=""
unit=""
end if
if isnull(rs.fields.item("strContainerSize")) then
updateSize="update tblChemicalContainer set strContainerSize='0 g' where numChemicalContainerID=" & rs.fields.item("numChemicalContainerID")
'response.write(updateSize & "<br/>")
conn.execute updateSize
end if
if isnull(rs.fields.item("numQuantity")) then
updateQuantity="update tblChemicalContainer set numQuantity=0 where numChemicalContainerID=" & rs.fields.item("numChemicalContainerID")
conn.execute updateQuantity
'response.write(updateQuantity & "<br/>")
end if
rs.movenext
loop
rs.close

Response.write("Update Completed! All Size and unit are separated. All Null value in quantity changed to 0 and null value in size and unit changed to '0 g' ..")

'DLJ added  Or ((mid(str,i,1)) = ".") 23Oct2008

function putSpace(length,str)
for i=1 to length

	if (isnumeric(mid(str,i,1))) Or ((mid(str,i,1)) = ".") then
	size=size+mid(str,i,1)
	else
	unit=unit+mid(str,i,1)
	end If
next
putSpace=size + " " + unit
end function
%>