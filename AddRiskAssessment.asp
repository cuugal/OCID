<%@ Language=VBScript %>
<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%	
Sub CleanUp() 

	set rsRA = nothing
	conn.close
	set conn = nothing

End Sub

Dim numChemicalID
Dim conn 
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr


If Request.Form("ADD") = "Add the New Risk Assessment" then 
	
	Dim rsRA
	Dim strSQL

	set rsRA = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "INSERT INTO tblRiskAssessment ("
	For each item in Request.Form
	If (item = "ADD") then
		'Do nothing
	Else
		strSQL = strSQL + item + ", "
	End if
	Next
	strSQL = Left(strSQL, (Len(strSQL)-2))
	strSQL = strSQL + ") VALUES ("
	For each item in Request.Form
	If (item = "ADD") then
		'Do nothing
	Else
	If (item = "numChemicalContainerID") then
		strSQL = strSQL + InjectionEncode(Request.Form(item)) + ", "
	Else
		strSQL = strSQL + "'" + InjectionEncode(Request.Form(item)) + "', "
	End if
	End If
	Next
	strSQL = Left(strSQL, (Len(strSQL)-2))
	strSQL = strSQL + ")"
	'Response.write(strSQL)
	rsRA.Open strSQL, conn, 3, 3
	CleanUp()
	Response.Write ("The Risk Assessment has been Added")
	Response.End
	
Else

	numLocationID = Request.QueryString("numLocationID")
	
	'Dim strLoginID
	'strLoginID = lcase(session("strLoginID"))
	'if strLoginID <> "admin" then
	'	set rsLocation = Server.CreateObject("ADODB.Recordset")
	'	strSQL = "SELECT tbllocation.strLoginID, tbllocation.numLocationID "
	'	strSQL = strSQL +  "FROM tblLocation "
	'	strSQL = strSQL +  "WHERE tblLocation.numLocationID = " + numLocationID
	'	rsLocation.Open strSQL, conn, 3, 3
	'	if (rsLocation("strLoginID") <> strLoginID) then
	'		Response.Write "You do not have permission to Add Risk Assessments to chemicals at the chosen location, please contact the Administrator if you should."
	'		Response.End
	'	end if
	'end if
	
	Dim strChemicalName
	Dim strStoreManager
	numChemicalID = Request.QueryString("numChemicalID")
	strChemicalName = Request.QueryString("strChemicalName")
	strStoreManager = Request.QueryString("strStoreManager")

End If
%>




<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<% dim strDate
strDate = DanDate(Date, "%d/%m/%Y" )
%>
<link href="ocid.css" rel="stylesheet" type="text/css" />
</HEAD>
<BODY>
<FORM action="AddRiskAssessment.asp" name=frmAddRiskAssessment method=post>
<INPUT value="<%= numChemicalID%>" type="hidden" name="numChemicalContainerID">
<INPUT value="<%= strChemicalName%>" type="hidden" name="strChemicalName">
<INPUT value="<%= numLocationID%>" type="hidden" name="numLocationID">


<FONT face=Arial size=5><STRONG>New Risk Assessment for the Use of a Hazardous Chemical</STRONG></FONT>
<FONT face=Arial size=2>

<P>To be completed for <b>each work activity</b> involving the <b>hazardous chemical</b></P>



<table border="0" cellpadding="0" cellspacing="0" width="98%">
    <TR>
        <TD>Substance:  <b><%= strChemicalName %></b></TD>
    </TR>
    <TR>
        <TD>Name of Assessor:  <INPUT name=strAssessorsName></TD>
		<TD>Supervisor:  <b><%= strStoreManager %></b></TD>
    </TR>
	<tr>
		<TD>Location of Use:  <INPUT name=strLocationOfUse style="HEIGHT: 22px; WIDTH: 149px" ></TD>
		<TD>Date of Assessment: <INPUT name=dtmDateOfAssessment value="<%= strDate %>" style="HEIGHT: 22px; WIDTH: 149px" ></TD>
	</tr>
</TABLE>

<TABLE>
	<tr>
		<td>Other persons involved in assessment:</td>
	</tr>
	<TR>
		<TD colspan="2"><textarea cols="70" rows="4" name="strOtherPersons" wrap="soft"></textarea></TD>
	</TR>
</TABLE>

<table border="0" cellpadding="0" cellspacing="0" width="98%">
	<TR>
		<TD bgcolor="#dddddd"><b>Note to Supervisors on Consultation:</b> Work health and safety (WHS) legislation requires that staff involved in the work activity 
		<u>must be consulted</u> during risk assessments, when decisions are made about the measures to be taken to eliminate or control health and safety risks, and when risk assessments are reviewed.</TD>
	</TR>
</TABLE>

<br/>

<table border="0" cellpadding="0" cellspacing="0" width="98%">
	<TR>
		<TD colSpan=4 bgcolor = #dddddd><STRONG>1. DESCRIPTION OF HAZARD</STRONG></TD>
	</TR>
	<TR>
		<TD>Work activity description. Include quantities and concentrations of the substance(s) used.</TD>
	</TR>
	<TR>
		<TD colSpan=5><textarea cols="70" rows="4" name="strWorkActivity" wrap="soft"></textarea></TD>
	</TR>
	<TR>
		<TD>Note any hazardous reaction product(s) formed during the work activity.  Ensure that the control measures for these products are also included in this assessment.</TD>
	</TR>
	<TR>
		<TD colSpan=5><textarea cols="70" rows="4" name="strHazardousProducts" wrap="soft"></textarea></TD>
	</TR>
</table>

<br/>

<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
	<TR>
		<TD colspan=4 bgcolor = #dddddd><STRONG>2. HAZARDOUS NATURE OF SUBSTANCE(S)</STRONG></TD>
	</TR>
	<TR>
		<TD colspan=4>Referring to the Safety Data Sheet (SDS), complete the following:</TD>
	</TR>
	<tr>
		<TD><INPUT name=strHazardExplosive type=checkbox>Explosive</td>
		<td><INPUT name=strFlammableGas type=checkbox>Flammable Gas</td>
		<td><INPUT name=strGasUnderPressure type=checkbox>Gas Under Pressure</TD>
		<TD><INPUT name=strHazardFlamable type=checkbox>Flammable Liquid</td>
	</tr>
	<tr>
		<td><INPUT name=strFlammableSolid type=checkbox>Flammable Solid</td>
		<td><INPUT name=strHazardSpontaneouslyCombustable type=checkbox>Self-Reactive Substance</TD>
		<TD><INPUT name=strPyrophoricSubstance type=checkbox>Pyrophoric Substance</td>
		<td><INPUT name=strHazardOxidiser type=checkbox>Oxidiser</td>
	</tr>
	<tr>
	    <TD><INPUT name=strHazardDangerousWhenWet type=checkbox>Dangerous When Wet</TD>
		<TD><INPUT name=strOrganicPeroxide type=checkbox>Organic Peroxide</td>
		<td><INPUT name=strHazardCorrosive type=checkbox>Corrosive</td>
		<td><INPUT name=strHazardAcuteToxicity type=checkbox>Toxic</TD>
	</tr>
	<tr>
		<TD><INPUT name=strHazardChronicToxicity type=checkbox>Cumulative Effects</td>
		<td><INPUT name=strHazardAsphyxiant type=checkbox>Asphyxiant</td>
		<td><INPUT name=strHazardIrritant type=checkbox>Irritant</TD>
		<TD><INPUT name=strHazardSensitiser type=checkbox>Sensitiser</td><td>
	</tr>
	<tr>
		<TD><INPUT name=strHazardMutagenic type=checkbox>Mutagen</td>
		<td><INPUT name=strHazardCarcinogen type=checkbox>Carcinogen</TD>
		<TD><INPUT name=strHazardTeratogen type=checkbox>Toxic to Reproduction</td>
		<td><INPUT name=strHazardHarmfulToEnvironment type=checkbox>Aquatic Toxicity</td>
		<td><!--INPUT name=strHazardRadioactive type=checkbox>Radioactive --></TD>
	</tr>
</table>

<table>
	<tr><td><br/></td></tr>
	<TR>
		<TD>What specific health effects can the substance cause?</TD>
	</tr>
	<tr>
		<td colspan=3>Examples: Burns to skin, toxicity, chronic toxicity, systemic poisoning, asthma, cancer, respiratory irritation, skin irritation, dermatitis, eye damage, target organ system toxicity, asphyxiation, harm from explosion, burns from fire</td>
	</tr>
	<tr>
		<TD><textarea name="strSpecificHealthEffects" cols="70" rows="4"></textarea></TD>
	</TR>
	<tr><td><br/></td></tr>
	<TR>
		<TD><B>Hazard Level of Substance(s):</B></TD>
	</tr>
</TABLE>

<TABLE>
	<tr>
		<TD>
		<INPUT name=strHazardLevel type=radio value="high">High&nbsp;&nbsp; 
		<INPUT name=strHazardLevel type=radio value="medium">Medium&nbsp;&nbsp; 
		<INPUT name=strHazardLevel type=radio value="low">Low
		</TD>
	</TR>
</TABLE>

<br/>

<table border="0" cellpadding="0" cellspacing="0" width="98%">
	<TR>
		<TD colSpan=4 bgcolor = #dddddd><STRONG>3. EXPOSURE TO THE SUBSTANCE(S) IN THIS WORK ACTIVITY</STRONG></TD>
	</TR>
	<TR>
		<TD colSpan=3><B>How often is the work activity performed each semester?</B><INPUT name=strDurationOfExposure></TD>
	</TR>
	<tr><td><br/></td></tr>
	<TR>
        <TD>Note the <B>level of exposure </B>(with existing controls):</TD>
	</tr>	
    <tr>
		<TD colSpan=3>
		<INPUT name=strLvlOfExposure type=radio value="not significant">Not significant&nbsp;&nbsp;
		<INPUT name=strLvlOfExposure type=radio value="low">Low&nbsp;&nbsp;
		<INPUT name=strLvlOfExposure type=radio value="medium">Medium&nbsp;&nbsp;          
		<INPUT name=strLvlOfExposure type=radio value="high">High&nbsp;&nbsp;     
		<INPUT name=strLvlOfExposure  type=radio value="uncertain">Uncertain
		</TD>
	</TR>
	<tr><td><br/></td></tr>
	<TR>
        <TD>Note the likely <B>route(s) of exposure </B>(with existing controls):</TD>
	</tr> 
	<tr>
		<TD>
		<INPUT name=strRouteInhalation type=checkbox>Inhalation
		<INPUT name=strRouteSkinContact type=checkbox>Skin contact
		<INPUT name=strRouteInjection type=checkbox>Injection/needlestick
        <INPUT name=strRouteIngestion type=checkbox>Ingestion
        <INPUT name=strRouteEyeContact type=checkbox>Eye contact
		</td>
	</TR>
</TABLE>

<br/>

<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
	<TR>
		<TD colspan="4" bgcolor = "#dddddd"><STRONG>4. SAFETY CONTROL MEASURES SELECTED</STRONG></TD>
	</TR>
	<TR>
		<TD colspan="4">Note the controls (both existing and new) needed to minimise the risk of exposure during this work activity.</TD>
	</TR>
	<tr>
		<TD colspan="4"><b>Engineering Controls</b></td>
	</tr>
	<tr>
		<TD><INPUT name=strControlFumeCupboard type=checkbox>Fume cupboard</TD>
		<TD><INPUT name=strControlLocalExhaustVentilation type=checkbox>Local exhaust ventilation</TD>
		<TD><INPUT name=strControlGeneralVentilation type=checkbox>General ventilation</TD>
		<TD></TD>
	</tr>
	<tr><td><br/></td></tr>
	<tr>
		<TD colspan="4"><b>Administrative Controls</b></td>
	</tr>
	<TR>
		<TD><INPUT name=strControlTraining type=checkbox>Training/induction</TD>
		<TD><INPUT name=strControlRestrictedAccess type=checkbox>Restricted access</TD>
		<TD><INPUT name=strControlColleagueInAttendance type=checkbox>Colleague in attendance</TD>
		<TD><INPUT name=strSafeWorkProcedures type=checkbox>Safe work procedures</TD>
	</tr>
	<tr><td><br/></td></tr>
	<tr>
		<TD colspan="4"><b>Personal Protective Controls</b></td>
	</tr>
	<TR>
		<TD><INPUT name=strControlLabCoat type=checkbox>Lab coat</TD>
		<TD><INPUT name=strControlSafetyGlasses type=checkbox>Safety glasses</TD>
			<!--TD><INPUT name=strControlGloves type=checkbox><FONT face=Arial size=2>gloves</FONT></TD-->
		<TD><INPUT name=strControlFaceshield type=checkbox>Face shield</TD>
		<TD><INPUT name=strControlRespirator type=checkbox>Respirator</TD>
	</TR>
		<TD colspan="4">Glove type:<INPUT name=strGloveType size="70"></TD>
	</TR>
	<tr>
        <TD colspan="4">Other safety control measures:<INPUT name=strControlOther size="70"></TD>	
	</tr>
</TABLE>

<br/>

<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
	<TR>
		<TD colspan="4" bgcolor = "#dddddd"><FONT face=Arial><STRONG>5.  EMERGENCY FACILIITIES</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD colspan="4">Note the emergency facilities that must be available during the work activity</TD>
	</TR>
    <TR>
		<!--TD><FONT face=Arial size=2><INPUT name=strFacilitiesSkillkit type=checkbox>Spillkit</FONT></TD-->
        <TD><INPUT name=strFacilitiesEyeWashStation type=checkbox>Eye wash station</TD>
        <TD><INPUT name=strFacilitiesAntidoteKeptOnHand type=checkbox>Antidote kept on-hand</TD> 
        <TD><INPUT name=strFireBlanket type=checkbox>Fire blanket</TD>
		<TD></TD>
		<!--TD><FONT face=Arial size=2><INPUT name=strFacilitiesHealthSurveillance type=checkbox>Health surveillance</FONT></TD--> 	
	</TR>
    <TR>
		<TD><INPUT name=strFacilitiesFirstAidKit type=checkbox>First aid kit</TD>
        <TD><INPUT name=strFacilitiesSafetyShower type=checkbox>Safety shower</TD>
        <TD><INPUT name=strExposureMonitoring type=checkbox>Exposure monitoring</TD>
		<TD></TD>
		<!--TD><FONT face=Arial size=2><INPUT name=strFacilitiesEvacuationProcedures type=checkbox>Evacuation/fire induction</FONT></TD-->
    </TR>
	<tr>
		<TD>Spill kit type:<INPUT name=strSpillkitType></TD>		
		<TD>Extinguisher type:<INPUT name=strExtinguisherType></TD>
		<TD>Other: <INPUT name=strFacilitiesOther></TD>
		<TD></TD>
	</tr>
</TABLE>

<BR>

<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
	<TR>
		<TD colspan="3" bgcolor = #dddddd><FONT face=Arial><STRONG>6.  ESTIMATED RISK</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD colspan="3">The <b>estimated risk</b> is based on the <b>nature of the hazard</b> and the <b>degree of exposure</b></TD>
	</TR>
	<TR>
		<TD colspan="3">Select the option that best describes the level of estimated risk</TD>
	</TR>
	<TR>
		<TD>
		<INPUT name=strRiskSignificant type=radio value=False>Risks are NOT Significant<BR>
		<INPUT name=strRiskSignificant type=radio value=True>Risks are significant, since the proposed controls are not adequate (if so, repeat this assessment when the risks have been adequately controlled)<br/>
		<INPUT name=strRiskControlled type=radio value=True>Risks will be adequately controlled<BR>
		<INPUT name=strRiskControlled type=radio value=False>Risks are uncertain and more information required (if so, repeat this assessment when more information is obtained)<br/>
		<hr></TD>
	</TR>
</TABLE>

A detailed assessment may be required where complex chemical processes or exposures are involved.<BR>
Risk assessment must be reviewed in 2 years or if the job or substance changes or new information becomes available.<BR>
</FONT>
<P>
<!-- <INPUT name=reset type=reset value="Clear form">&nbsp;&nbsp;  --><INPUT name=ADD type=submit value="Add the New Risk Assessment">
</P>
</FORM>
</BODY>
</HTML>
