<%
Dim constr
constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

Function InjectionEncode(str)
	InjectionEncode=Replace(str,"'","''")
End Function

%>