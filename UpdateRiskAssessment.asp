<%@ Language=VBScript %>
<!--#INCLUDE FILE="date.inc"-->
<!--#INCLUDE FILE="DbConfig.asp"-->

<%	
Response.Buffer = True
Sub CleanUp() 


	set rsRA = nothing
	conn.close
	set conn = nothing

End Sub

Dim numChemicalID
Dim numRiskAssessmentID
Dim rsRA
Dim strSQL
Dim conn 
Dim action
'Dim constr
'constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")

set conn = Server.CreateObject("ADODB.Connection")
conn.open constr
set rsRA = Server.CreateObject("ADODB.Recordset") 

action = Request.Form("action")


If action = "SAVE Changes to Risk Assessment" then 
	
	numRiskAssessmentID = Request.Form("hdnRiskAssessmentID")
	
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open constr
	set rsRA = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "UPDATE tblRiskAssessment SET "
	For each item in Request.Form

	'remove the item called numChemicalContainerID from the query, since it could be null if the chemical container has beed deleted. Should remove redundant IF statement below DLJ 21July2015
	If ((item = "action") OR (item = "hdnRiskAssessmentID" ) OR (item = "numChemicalContainerID")) then

	'If ((item = "action") OR (item = "hdnRiskAssessmentID" )) then
		'Do nothing
	Else
		strSQL = strSQL + item + " = "
		If (item = "numChemicalContainerID") then
		strSQL = strSQL + InjectionEncode(Request.Form(item)) + ", "
		Else
		strSQL = strSQL + "'" + InjectionEncode(Request.Form(item)) + "', "
		End if
		
	End if
	Next

	strSQL = Left(strSQL, (Len(strSQL)-2))
	strSQL = strSQL + " WHERE (numRiskAssessmentID = " + numRiskAssessmentID +")"
	'Response.write(strSQL)
	rsRA.Open strSQL, conn, 3, 3
	CleanUp()
	Response.Write ("The Risk Assessment has been Updated")
	Response.End
	
Else 
	if action = "Delete this assessment" then
	
	numRiskAssessmentID = Request.Form("hdnRiskAssessmentID")

	strSQL = "DELETE From tblRiskAssessment WHERE (numRiskAssessmentID = "
	strSQL = strSQL + numRiskAssessmentID + ")"
	
	rsRA.Open strSQL, conn, 3, 3
	CleanUp()
	Response.Write ("The Risk Assessment has been Deleted")
	Response.End
	end if
end if

numRiskAssessmentID = Request.QueryString("numRiskAssessmentID")

numLocationID = Request.QueryString("numLocationID")

strSQL = "SELECT * FROM tblRiskAssessment WHERE numRiskAssessmentID = "
strSQL = strSQL + numRiskAssessmentID

rsRA.Open strSQL, conn, 3, 3


	Dim strChemicalName
	Dim strStoreManager


	strChemicalName = Request.QueryString("strChemicalName")
	strStoreManager = Request.QueryString("strStoreManager")
	numChemicalContainerID = Request.QueryString("numChemicalID")

%>
<HTML>
<HEAD>
<link href="ocid.css" rel="stylesheet" type="text/css" />
</HEAD>
<BODY>
<form action="UpdateRiskAssessment.asp" method="post">



<FONT face=Arial size=5><STRONG>Risk Assessment for the Use of a Hazardous Chemical</STRONG></FONT>
<FONT face=Arial size=2>

<P>To be completed for <b>each work activity</b> involving the <b>hazardous chemical</b></P>




<table border="0" cellpadding="0" cellspacing="0" width="98%">
    <TR>
        <TD>Substance:  <b><%= strChemicalName %></b></TD>
    </TR>
    <TR>
        <TD>Risk Assessment No.:  <b><%= numRiskAssessmentID %></b></TD>
    </TR>
    <TR>
        <TD>Name of Assessor:  <INPUT name=strAssessorsName value="<%=rsRA("strAssessorsName")%>"></TD>
        <!--TD>Name of Assessor:  <%= strAssessorsName %></TD-->
		<TD>Supervisor:  <b><%= strStoreManager %></b></TD>
    </TR>
	<tr>
		<TD style="padding-top: 0.2em;">Location of Use:  <INPUT name=strLocationOfUse value="<%=rsRA("strLocationOfUse")%>" style="width: 149px;" ></TD>
		<TD>Date of Assessment:  <INPUT name=dtmDateOfAssessment value="<%=rsRA("dtmDateOfAssessment")%>" style="width: 149px;"></TD>
	</tr>
</TABLE>

<TABLE>
	<tr>
		<td>Other persons involved in assessment:</td>
	</tr>
	<TR>
		<TD colspan=2><textarea cols="70" rows="4" name="strOtherPersons" wrap="soft"><%=rsRA("strOtherPersons")%></textarea></TD>
	</TR>
</TABLE>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=98%>
	<TR>
		<TD bgcolor="#dddddd"><b>Note to Supervisors on Consultation:</b> Work health and safety (WHS) legislation requires that staff involved in the work activity 
		<u>must be consulted</u> during risk assessments, when decisions are made about the measures to be taken to eliminate or control health and safety risks, and when risk assessments are reviewed.</TD>
	</TR>
</TABLE>




<br/>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=98%>
	<TR>
		<TD colspan="4" bgcolor = #dddddd><STRONG>1. DESCRIPTION OF HAZARD</STRONG></TD>
	</TR>
	<TR>
		<TD>Work activity description. Include quantities and concentrations of the substance(s) used.</TD>
	</TR>
	<TR>
		<TD colspan=5><textarea cols="70" rows="4" name="strWorkActivity" wrap="soft"><%=rsRA("strWorkActivity")%></textarea></TD>
	</TR>
	<TR>
		<TD>Note any hazardous reaction product(s) formed during the work activity.  Ensure that the control measures for these products are also included in this assessment.</TD>
	</TR>
	<TR>
		<TD colspan=5><textarea cols="70" rows="4" name="strHazardousProducts" wrap="soft"><%=rsRA("strHazardousProducts")%></textarea></TD>
	</TR>
</table>

<br/>





<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
	<TR>
		<TD colspan="4" bgcolor = #dddddd><STRONG>2. HAZARDOUS NATURE OF SUBSTANCE(S)</STRONG></TD>
	</TR>
	<TR>
		<TD colspan="4">Referring to the Safety Data Sheet (SDS), complete the following:</TD>
	</TR>
	<tr>
		<TD><INPUT name=strHazardExplosive  <% If (rsRA("strHazardExplosive") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Explosive</td>
		<td><INPUT name=strFlammableGas  <% If (rsRA("strFlammableGas") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Flammable Gas</td>
		<td><INPUT name=strGasUnderPressure  <% If (rsRA("strGasUnderPressure") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Gas Under Pressure</TD>
		<TD><INPUT name=strHazardFlamable  <% If (rsRA("strHazardFlamable") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Flammable Liquid</td>
	</tr>
	<tr>
		<td><INPUT name=strFlammableSolid  <% If (rsRA("strFlammableSolid") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Flammable Solid</td>
		<td><INPUT name=strHazardSpontaneouslyCombustable  <% If (rsRA("strHazardSpontaneouslyCombustable") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Self-Reactive Substance</TD>
		<TD><INPUT name=strPyrophoricSubstance  <% If (rsRA("strPyrophoricSubstance") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Pyrophoric Substance</td>
		<td><INPUT name=strHazardOxidiser  <% If (rsRA("strHazardOxidiser") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Oxidiser</td>
	</tr>
	<tr>
	    <TD><INPUT name=strHazardDangerousWhenWet  <% If (rsRA("strHazardDangerousWhenWet") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Dangerous When Wet</TD>
		<TD><INPUT name=strOrganicPeroxide  <% If (rsRA("strOrganicPeroxide") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Organic Peroxide</td>
		<td><INPUT name=strHazardCorrosive  <% If (rsRA("strHazardCorrosive") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Corrosive</td>
		<td><INPUT name=strHazardAcuteToxicity  <% If (rsRA("strHazardAcuteToxicity") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Toxic</TD>
	</tr>
	<tr>
		<TD><INPUT name=strHazardChronicToxicity  <% If (rsRA("strHazardChronicToxicity") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Cumulative Effects</td>
		<td><INPUT name=strHazardAsphyxiant  <% If (rsRA("strHazardAsphyxiant") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Asphyxiant</td>
		<td><INPUT name=strHazardIrritant  <% If (rsRA("strHazardIrritant") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Irritant</TD>
		<TD><INPUT name=strHazardSensitiser  <% If (rsRA("strHazardSensitiser") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Sensitiser</td><td>
	</tr>
	<tr>
		<TD><INPUT name=strHazardMutagenic  <% If (rsRA("strHazardMutagenic") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Mutagen</td>
		<td><INPUT name=strHazardCarcinogen  <% If (rsRA("strHazardCarcinogen") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Carcinogen</TD>
		<TD><INPUT name=strHazardTeratogen  <% If (rsRA("strHazardTeratogen") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Toxic to Reproduction</td>
		<td><INPUT name=strHazardHarmfulToEnvironment  <% If (rsRA("strHazardHarmfulToEnvironment") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Aquatic Toxicity</td>
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
		<TD><textarea name="strSpecificHealthEffects" cols="70" rows="4"><%=rsRA("strSpecificHealthEffects")%></textarea></TD>
	</TR>
	<tr><td><br/></td></tr>
	<TR>
		<TD><B>Hazard Level of Substance(s):</B></TD>
	</tr>
</TABLE>

<TABLE>
	<tr>
		<TD>
		<INPUT name=strHazardLevel type=radio value="high"  <% If (rsRA("strHazardLevel") = "high") then 
					Response.Write " CHECKED "
				END IF %>  >High&nbsp;&nbsp; 
	<INPUT name=strHazardLevel type=radio value="medium"  <% If (rsRA("strHazardLevel") = "medium") then 
					Response.Write " CHECKED "
				END IF %>  >Medium&nbsp;&nbsp; 
	<INPUT name=strHazardLevel type=radio value="low"  <% If (rsRA("strHazardLevel") = "low") then 
					Response.Write " CHECKED "
				END IF %>  >Low
		</TD>
	</TR>
</TABLE>

<br/>


<TABLE border=0 cellPadding=0 cellSpacing=0 width=98%>
	<TR>
		<TD colspan="4" bgcolor = #dddddd><STRONG>3. EXPOSURE TO THE SUBSTANCE(S) IN THIS WORK ACTIVITY</STRONG></TD>
	</TR>
	<TR>
		<TD colspan=3><B>How often is the work activity performed each semester?</B><INPUT name=strDurationOfExposure value="<%=rsRA("strDurationOfExposure")%>"></TD>
	</TR>
	<tr><td><br/></td></tr>
	<TR>
        <TD>Note the <B>level of exposure </B>(with existing controls):</TD>
	</tr>	
    <tr>
		<TD colspan=3>
		<INPUT name=strLvlOfExposure type=radio value="not significant"  <% If (rsRA("strLvlOfExposure") = "not significant") then 
					Response.Write " CHECKED "
				END IF %>  >Not significant&nbsp;&nbsp;
	<INPUT name=strLvlOfExposure type=radio value="low"  <% If (rsRA("strLvlOfExposure") = "low") then 
					Response.Write " CHECKED "
				END IF %>  >Low&nbsp;&nbsp;
	<INPUT name=strLvlOfExposure type=radio value="medium"  <% If (rsRA("strLvlOfExposure") = "medium") then 
					Response.Write " CHECKED "
				END IF %>  >Medium&nbsp;&nbsp;          
	<INPUT name=strLvlOfExposure type=radio value="high"  <% If (rsRA("strLvlOfExposure") = "high") then 
					Response.Write " CHECKED "
				END IF %>  >High&nbsp;&nbsp;     
	<INPUT name=strLvlOfExposure  type=radio value="uncertain"  <% If (rsRA("strLvlOfExposure") = "uncertain") then 
					Response.Write " CHECKED "
				END IF %>  >Uncertain
		</TD>
	</TR>
	<tr><td><br/></td></tr>
	<TR>
        <TD>Note the likely <B>route(s) of exposure </B>(with existing controls):</TD>
	</tr> 
	<tr>
		<TD>
		<INPUT name=strRouteInhalation <% If (rsRA("strRouteInhalation") = "on") then 
					Response.Write " CHECKED "
				END IF %>  type=checkbox>Inhalation
		<INPUT name=strRouteSkinContact  <% If (rsRA("strRouteSkinContact") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Skin contact
		<INPUT name=strRouteInjection  <% If (rsRA("strRouteInjection") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Injection/needlestick
        <INPUT name=strRouteIngestion  <% If (rsRA("strRouteIngestion") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Ingestion
        <INPUT name=strRouteEyeContact  <% If (rsRA("strRouteEyeContact") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Eye contact
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
		<TD><INPUT name=strControlFumeCupboard  <% If (rsRA("strControlFumeCupboard") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Fume cupboard</TD>
	<TD><INPUT name=strControlLocalExhaustVentilation  <% If (rsRA("strControlLocalExhaustVentilation") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Local exhaust ventilation</td>
	<TD><INPUT name=strControlGeneralVentilation  <% If (rsRA("strControlGeneralVentilation") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>General ventilation</TD>
		<TD></TD>
	</tr>
	<tr><td><br/></td></tr>


	<tr>
		<TD colspan="4"><b>Administrative Controls</b></td>
	</tr>
	<TR>
		<TD><INPUT name=strControlTraining  <% If (rsRA("strControlTraining") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Training/induction</TD>
	<TD><INPUT name=strControlRestrictedAccess  <% If (rsRA("strControlRestrictedAccess") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Restricted access</TD>
	<TD><INPUT name=strControlColleagueInAttendance  <% If (rsRA("strControlColleagueInAttendance") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Colleague in attendance</TD>
    <TD><INPUT name=strSafeWorkProcedures  <% If (rsRA("strSafeWorkProcedures") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Safe work procedures</TD>
	</tr>
	<tr><td><br/></td></tr>


	<tr>
		<TD colspan="4"><b>Personal Protective Controls</b></td>
	</tr>
	<TR>
		<TD><INPUT name=strControlLabCoat  <% If (rsRA("strControlLabCoat") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT>Lab coat</FONT></TD>
		<TD><INPUT name=strControlSafetyGlasses <% If (rsRA("strControlSafetyGlasses") = "on") then 
					Response.Write " CHECKED "
				END IF %>  type=checkbox><FONT>Safety glasses</FONT></TD>
		<!--TD><INPUT name=strControlGloves <% If (rsRA("strControlGloves") = "on") then 
					Response.Write " CHECKED "  
				END IF %>  type=checkbox><FONT face=Arial size=2>gloves</FONT></TD-->
		<TD><INPUT name=strControlFaceshield  <% If (rsRA("strControlFaceShield") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Face shield</TD>
        <TD><INPUT name=strControlRespirator  <% If (rsRA("strControlRespirator") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Respirator</TD>
	</TR>
		<TD colspan="4">Glove type:<INPUT name=strGloveType value="<%=rsRA("strGloveType")%>" ></TD>
	</TR>
	<tr>
        <TD colspan="4">Other safety control measures:<INPUT name=strControlOther value="<%=rsRA("strControlOther")%>" size="70"></TD>	
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
        <TD><INPUT name=strFacilitiesEyeWashStation  <% If (rsRA("strFacilitiesEyeWashStation") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Eye wash station</TD>
        <TD><INPUT name=strFacilitiesAntidoteKeptOnHand  <% If (rsRA("strFacilitiesAntidoteKeptOnHand") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Antidote kept on-hand</TD>
        <TD><INPUT name=strFireBlanket  <% If (rsRA("strFireBlanket") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Fire blanket</TD>
		<TD></TD>
		<!-- the following controls are now deprecated -->
		<!--TD><INPUT name=strFacilitiesSpillkit  <% If (rsRA("strFacilitiesSpillkit") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>spillkit</TD-->
		<!--td><INPUT name=strFacilitiesHealthSurveillance  <% If (rsRA("strFacilitiesHealthSurveillance") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>health surveillance</td-->

		<!--TD><INPUT name=strFacilitiesEvacuationProcedures  <% If (rsRA("strFacilitiesEvacuationProcedures") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>evacuation/fire induction</TD-->
	</TR>
    <TR>
		<TD><INPUT name=strFacilitiesFirstAidKit  <% If (rsRA("strFacilitiesFirstAidKit") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>First aid kit</TD>
        <TD><INPUT name=strFacilitiesSafetyShower  <% If (rsRA("strFacilitiesSafetyShower") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Safety shower</TD>
		<TD><INPUT name=strExposureMonitoring  <% If (rsRA("strExposureMonitoring") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Exposure monitoring</TD>
		<TD></TD>
    </TR>
	<tr>
		<TD>Spill kit type:<INPUT name=strSpillkitType value="<%=rsRA("strSpillkitType")%>"></TD>
		<TD>Extinguisher type:<INPUT name=strExtinguisherType  value="<%=rsRA("strExtinguisherType")%>"></TD>
        <TD colspan=3>Other:<INPUT name=strFacilitiesOther  value="<%=rsRA("strFacilitiesOther")%>" ></TD>
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
	<INPUT name=strRiskSignificant type=radio value=False <% If (rsRA("strRiskSignificant") = "False") then 
					Response.Write " CHECKED "
				END IF %>>Risks Are NOT Significant<br />
	<INPUT name=strRiskSignificant type=radio value=True <% If (rsRA("strRiskSignificant") = "True") then 
					Response.Write " CHECKED "
				END IF %>>Risks are significant, since the proposed controls are not adequate (if so, repeat this assessment when the risks have been adequately controlled)<br />
	<INPUT name=strRiskControlled type=radio value=True <% If (rsRA("strRiskControlled") = "True") then 
					Response.Write " CHECKED "
				END IF %>>Risks will be Adequately Controlled<br />
	<INPUT name=strRiskControlled type=radio value=False <% If (rsRA("strRiskControlled") = "False") then 
					Response.Write " CHECKED "
				END IF %>>Risks are uncertain and more information required (if so, repeat this assessment when more information is obtained)<br />
		 </TD>
	</TR>
</TABLE>

<BR>

<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
	<TR>
		<TD colspan="6" bgcolor = "#dddddd"><FONT face=Arial><STRONG>7.  DECLARATION</STRONG></FONT></TD>
	</TR>
	
	<TR>
		<TD colspan="6">Sign below once all recommended controls and emergency facilities are available.</TD>
	</TR>
	
	<TR>
		<TD colspan="6" bgcolor = "#dddddd">Assessment developed by:</TD>
	</TR>
	<TR>
		<TD>Assessors name: <br/>&nbsp </TD><TD>     </TD>
		<TD>Signature:  <br/>&nbsp</TD><TD>&nbsp;&nbsp;&nbsp;&nbsp;</TD>
		<TD>Date: <br/>&nbsp </TD><TD>&nbsp;&nbsp;&nbsp;&nbsp;</TD>
	</TR>

	<TR>
		<TD colspan="6" bgcolor = "#dddddd">Assessment approval: </ br> I am satisfied that the risks will be adequately controlled and that the resources required are available.</TD>
	</TR>
	<TR>
		<TD>Supervisor of person performing the activity: <br/>&nbsp </TD><TD>     </TD>
		<TD>Signature: <br/>&nbsp </TD><TD>&nbsp;&nbsp;&nbsp;&nbsp;     </TD>
		<TD>Date: <br/>&nbsp </TD><TD>&nbsp;&nbsp;&nbsp;&nbsp;     </TD>
	</TR>

	<TR>
		<TD colspan="6" bgcolor="#dddddd">
		A detailed assessment may be required where complex chemical processes or exposures are involved.<BR>
		Risk assessment must be reviewed in 2 years or if the job or substance changes or new information becomes available.
		</TD>
	</TR>

</TABLE>


<input type="hidden" name="hdnRiskAssessmentID" value="<%= numRiskAssessmentID %>">
<input type="hidden" name="numChemicalContainerID" value="<%= numChemicalContainerID %>">
<INPUT value="<%= strChemicalName%>" type="hidden" name="strChemicalName">
<INPUT value="<%= numLocationID%>" type="hidden" name="numLocationID">

<P>
<input type="reset" value="Undo Changes">&nbsp;&nbsp;
<!--INPUT name=action type=submit value="Update"-->
<!-- also made change to line 32 which uses value property in code "SAVE Changes to Risk Assessment". Could be better coding. -->
<INPUT name=action type=submit value="SAVE Changes to Risk Assessment">&nbsp;&nbsp; 
<INPUT name=action type=submit value="Delete this assessment">
</P>

</FORM>
</BODY>
</HTML>