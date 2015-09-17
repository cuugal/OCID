<%@ Language=VBScript %>
<!--#INCLUDE FILE="DbConfig.asp"-->
<%	
Sub CleanUp() 


	set rsRA = nothing
	conn.close
	set conn = nothing

End Sub

Dim rsRA
Dim strSQL
Dim conn
Dim numRiskAssessmentID

numRiskAssessmentID = Request.QueryString("numRiskAssessmentID")
strSQL = "SELECT * FROM tblRiskAssessment WHERE numRiskAssessmentID = "
strSQL = strSQL + numRiskAssessmentID

set conn = Server.CreateObject("ADODB.Connection")
conn.open Application("strDSN")
set rsRA = Server.CreateObject("ADODB.Recordset") 
rsRA.Open strSQL, conn, 3, 3


	Dim strChemicalName
	Dim strStoreManager
	Dim numLocationId

	strChemicalName = Request.QueryString("strChemicalName")
	strStoreManager = Request.QueryString("strStoreManager")
	numLocationId = Request.QueryString("numLocationId")
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

</HEAD>
<BODY>
<FORM >
<INPUT value=<%= numChemicalID%> type="hidden" name="numChemicalContainerID">
<INPUT value=<%= strChemicalName%> type="hidden" name="strChemicalName">
<INPUT value=<%= numLocationID%> type="hidden" name="numLocationID">
<FONT face=Arial size=5><STRONG>Risk Assessment for the Use of a Hazardous Chemical</STRONG></FONT>
<p>To be completed for <b>each work activity</b> involving the <b>hazardous chemical</b></p>

<FONT face = Arial size=2>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=98%>
    <TR>
        <TD ><FONT face=Arial size=2>Substance:</FONT></TD>
        <TD><FONT face=Arial size=3><b><%= strChemicalName %></b></FONT></TD> 
    </TR>
    <TR>
        <TD ><FONT face=Arial><FONT size=2>Name of Assessor:</FONT></FONT></TD>
        <TD><INPUT name=strAssessorsName value="<%=rsRA("strAssessorsName")%>"></td>
		<TD><FONT face=Arial size=2>Supervisor:</FONT></TD>
        <TD><FONT face=Arial size=2><%= strStoreManager %></FONT></TD>  
    </TR>
	
	<tr>
		<TD><FONT face=Arial><FONT size=2>Location of Use:</FONT></FONT></TD>
        <TD>
            <INPUT name=strLocationOfUse value="<%=rsRA("strLocationOfUse")%>" style="HEIGHT: 22px; WIDTH: 149px" >
			</td>

		<TD><FONT face=Arial><FONT size=2>Date of Assessment:</FONT></FONT></TD>
        <TD>
            <INPUT name=dtmDateOfAssessment value="<%=rsRA("dtmDateOfAssessment")%>" style="HEIGHT: 22px; WIDTH: 149px" >
			</td>
		</tr>
		<tr>
			<td><FONT face=Arial size=2>Other persons involved in assessment:</font>
			</td>
		</tr>
		<TR>
			<TD colSpan=2><textarea cols="57" rows="4" name="strOtherPersons" wrap="soft"><%=rsRA("strOtherPersons")%></textarea>
			</TD>
		</TR>
</TABLE>
</FONT>
<br/>

<TABLE border=0 cellPadding=0 cellSpacing=0 width=98%>
<TR><TD colSpan=4 bgcolor = #dddddd><FONT face=Arial><STRONG>1. Description of Hazard</STRONG></FONT>
<TR><TD><FONT face=Arial size=2><B>Note the work activity undertaken, ensuring that you include quantities and concentrations of the substance(s) used.</B></FONT></TD></TR>
</TD><TR>
	<TD colSpan=5><textarea cols="57" rows="4" name="strWorkActivity" wrap="soft"><%=rsRA("strWorkActivity")%></textarea></TD>

 </TR>
 
 <TR><TD><FONT face=Arial size=2><B>Note any hazardous reaction product(s) formed during the work activity.  Ensure that the control measures for these products are also included.</B></FONT></TD></TR>
</TD><TR>
	<TD colSpan=5><textarea cols="57" rows="4" name="strHazardousProducts" wrap="soft"><%=rsRA("strHazardousProducts")%></textarea></TD>

 </TR>

 </table>
<br/>

<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
<TR><TD colspan=3 bgcolor = #dddddd><FONT face=Arial><STRONG>2. Hazardous Nature of Substance(s)</STRONG></FONT></TD></TR>
<TR><TD colspan=3><FONT face=Arial size=2><B>Referring to the Safety Data Sheet (SDS), complete the following:</B></FONT></TD></TR>
<tr><td colspan=3">
	<table width="50%">
	<tr>
		<TD>
		<INPUT name=strHazardExplosive  <% If (rsRA("strHazardExplosive") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Explosive
		</td><td>
		<INPUT name=strFlammableGas  <% If (rsRA("strFlammableGas") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Flammable Gas
		</td><td>
		<INPUT name=strGasUnderPressure  <% If (rsRA("strGasUnderPressure") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Gas Under Pressure
		</TD>
	</tr>
	<tr>
		<TD>
		<INPUT name=strHazardFlamable  <% If (rsRA("strHazardFlamable") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Flammable Liquid
		</td><td>
		<INPUT name=strFlammableSolid  <% If (rsRA("strFlammableSolid") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Flammable Solid
		</td><td>
		<INPUT name=strHazardSpontaneouslyCombustable  <% If (rsRA("strHazardSpontaneouslyCombustable") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Self-Reactive Substance
		</TD>
	</tr>
	<tr>
		<TD>
		<INPUT name=strPyrophoricSubstance  <% If (rsRA("strPyrophoricSubstance") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Pyrophoric Substance
		</td><td>
		<INPUT name=strHazardOxidiser  <% If (rsRA("strHazardOxidiser") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Oxidiser
		</td><td>
		<INPUT name=strHazardDangerousWhenWet  <% If (rsRA("strHazardDangerousWhenWet") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Dangerous When Wet
		</TD>
	</tr>
	<tr>
		<TD>
		<INPUT name=strOrganicPeroxide  <% If (rsRA("strOrganicPeroxide") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Organic Peroxide
		</td><td>
		<INPUT name=strHazardCorrosive  <% If (rsRA("strHazardCorrosive") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Corrosive
		</td><td>
		<INPUT name=strHazardAcuteToxicity  <% If (rsRA("strHazardAcuteToxicity") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Toxic
		</TD>
	</tr>
	<tr>
		<TD>
		<INPUT name=strHazardChronicToxicity  <% If (rsRA("strHazardChronicToxicity") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Cumulative Effects
		</td><td>
		<INPUT name=strHazardAsphyxiant  <% If (rsRA("strHazardAsphyxiant") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Asphyxiant
		</td><td>
		<INPUT name=strHazardIrritant  <% If (rsRA("strHazardIrritant") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Irritant
		</TD>
	</tr>
	<tr>
		<TD>
		<INPUT name=strHazardSensitiser  <% If (rsRA("strHazardSensitiser") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Sensitiser
		</td><td>
		<INPUT name=strHazardMutagenic  <% If (rsRA("strHazardMutagenic") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Mutagen
		</td><td>
		<INPUT name=strHazardCarcinogen  <% If (rsRA("strHazardCarcinogen") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Carcinogen
		</TD>
	</tr>
	<tr>
		<TD>
		<INPUT name=strHazardTeratogen  <% If (rsRA("strHazardTeratogen") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Toxic to Reproduction
		</td><td>
		<INPUT name=strHazardHarmfulToEnvironment  <% If (rsRA("strHazardHarmfulToEnvironment") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Aquatic Toxicity
		</td><td>
		<INPUT name=strHazardRadioactive  <% If (rsRA("strHazardRadioactive") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>Radioactive
		</TD>
	</tr>
	</table>
</td></tr>
<tr><td><br/></td></tr>
<TR>
	<TD><FONT face=Arial size=2>What specific health effects can the substance cause?</font></TD>
</tr>
<tr>
    <TD><FONT face=Arial size=2><textarea name="strSpecificHealthEffects" cols="57" rows="4"><%=rsRA("strSpecificHealthEffects")%></textarea></TD>
</TR>
<tr><td colspan=3>Examples: Burns to skin, toxicity, chronic toxicity, systemic poisoning, asthma, cancer, respiratory irritation, skin irritation, dermatitis, eye damage, 
target organ system toxicity, asphyxiation, harm from explosion, burns from fire</td>
</tr>
<tr><td><br/></td></tr>
<TR>
	<TD><FONT face=Arial size=2><B>Hazard Level of Substance(s):</B></font></TD>
</tr>
<tr>
    <TD><FONT face=Arial size=2>
	<INPUT name=strHazardLevel type=radio value=high  <% If (rsRA("strHazardLevel") = "high") then 
					Response.Write " CHECKED "
				END IF %>  ><FONT face=Arial size=2>high&nbsp;&nbsp; 
	<INPUT name=strHazardLevel style="LEFT: 58px; TOP: 1px" type=radio value=medium  <% If (rsRA("strHazardLevel") = "medium") then 
					Response.Write " CHECKED "
				END IF %>  >medium&nbsp;&nbsp; 
	<INPUT name=strHazardLevel type=radio value=low  <% If (rsRA("strHazardLevel") = "low") then 
					Response.Write " CHECKED "
				END IF %>  >low</TD>

</TR>

<tr><td colspan=3><b>Note to Supervisors on Consultation:</b> Work health and safety (WHS) legislation requires that staff involved in the work activity must be consulted during risk assessments,
when decisions are made about the measures to be taken to eliminate or control health and safety risks, and when risk assessments are reviewed.</td></tr>
</TABLE>


<br/>



<TABLE border=0 cellPadding=0 cellSpacing=0 width=98%>
<TR><TD colSpan=4 bgcolor = #dddddd><FONT face=Arial><STRONG>3. Exposure to the substance(s) in this work activity</STRONG></FONT>
<TR>
     <TD><FONT face=Arial size=2><B>How often is the work activity performed each semester?</B></FONT></TD></tr>
     <tr><TD colSpan=3><INPUT name=strDurationOfExposure value="<%=rsRA("strDurationOfExposure")%>"></TD>
</TR>
<tr><td><br/></td></tr>
<TR>
        <TD><FONT face=Arial size=2>Note the <B>Level of Exposure </B></FONT>(with existing controls):</TD></tr>	
        <tr><TD colSpan=3><FONT face=Arial size=2>
	<INPUT name=strLvlOfExposure type=radio value="not significant"  <% If (rsRA("strLvlOfExposure") = "not significant") then 
					Response.Write " CHECKED "
				END IF %>  >not significant&nbsp;&nbsp;
	<INPUT name=strLvlOfExposure type=radio value=low  <% If (rsRA("strLvlOfExposure") = "low") then 
					Response.Write " CHECKED "
				END IF %>  >low&nbsp;&nbsp;
	<INPUT name=strLvlOfExposure type=radio value=medium  <% If (rsRA("strLvlOfExposure") = "medium") then 
					Response.Write " CHECKED "
				END IF %>  >medium&nbsp;&nbsp;          
	<INPUT name=strLvlOfExposure type=radio value=high  <% If (rsRA("strLvlOfExposure") = "high") then 
					Response.Write " CHECKED "
				END IF %>  >high&nbsp;&nbsp;     
	<INPUT name=strLvlOfExposure  type=radio value=uncertain  <% If (rsRA("strLvlOfExposure") = "uncertain") then 
					Response.Write " CHECKED "
				END IF %>  >uncertain</TD>
</TR></FONT>
<tr><td><br/></td></tr>
<TR>
        <TD><FONT face=Arial size=2>Note the <B>Likely Route(s) of Exposure </B></FONT>(with existing controls):</TD></tr> 
        <tr><TD>
		<INPUT name=strRouteInhalation <% If (rsRA("strRouteInhalation") = "on") then 
					Response.Write " CHECKED "
				END IF %>  type=checkbox><FONT face=Arial size=2>inhalation</FONT>
		<INPUT name=strRouteSkinContact  <% If (rsRA("strRouteSkinContact") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>skin contact</FONT>
		<INPUT name=strRouteInjection  <% If (rsRA("strRouteInjection") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>injection/needlestick</FONT>
        <INPUT name=strRouteIngestion  <% If (rsRA("strRouteIngestion") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>ingestion</FONT>
        <INPUT name=strRouteEyeContact  <% If (rsRA("strRouteEyeContact") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>eye contact</FONT>
		</td>
</TR>
</TABLE>

<br/>




<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
<TR><TD colspan=3 bgcolor = #dddddd><FONT face=Arial><STRONG>4. Safety Control Measures Selected</STRONG></FONT></TD></TR>
<TR><TD colspan=3><FONT face=Arial size=2>note the controls (both existing and new) needed to minimise the risk of exposure during this work activity</FONT></TD></TR>
<tr><td><b>Engineering Controls</b></td></tr>
<tr>
	<TD><INPUT name=strControlFumeCupboard  <% If (rsRA("strControlFumeCupboard") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>fume cupboard</FONT></TD>
	<TD><INPUT name=strControlLocalExhaustVentilation  <% If (rsRA("strControlLocalExhaustVentilation") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>local exhaust ventilation</FONT> </td>
	<TD><INPUT name=strControlGeneralVentilation  <% If (rsRA("strControlGeneralVentilation") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>general ventilation</FONT></TD>
</tr>
<tr><td><br/></td></tr>
<tr><td><b>Administrative Controls</b></td></tr>
<TR>  
	<TD><INPUT name=strControlTraining  <% If (rsRA("strControlTraining") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>training/induction</FONT></TD>
	<TD><INPUT name=strControlRestrictedAccess  <% If (rsRA("strControlRestrictedAccess") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>restricted access</FONT></TD>
	<TD><INPUT name=strControlColleagueInAttendance  <% If (rsRA("strControlColleagueInAttendance") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>colleague in attendance</FONT></TD>
	</tr>
	<tr>
    <TD><INPUT name=strSafeWorkProcedures  <% If (rsRA("strSafeWorkProcedures") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>safe work procedures</FONT></TD>
	<TD></TD>
		<td></td>
</TR>

<tr><td><br/></td></tr>
<tr><td><b>Personal Protective Controls</b></td></tr>
<TR><FONT face=Arial size=2>
        <TD><INPUT name=strControlLabCoat  <% If (rsRA("strControlLabCoat") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>lab coat</FONT></TD>
		<TD><INPUT name=strControlSafetyGlasses <% If (rsRA("strControlSafetyGlasses") = "on") then 
					Response.Write " CHECKED "
				END IF %>  type=checkbox><FONT face=Arial size=2>safety glasses</FONT></TD>
		
		</tr><tr>
		<TD><INPUT name=strControlGloves <% If (rsRA("strControlGloves") = "on") then 
					Response.Write " CHECKED "
				END IF %>  type=checkbox><FONT face=Arial size=2>gloves</FONT></TD>
		<TD><FONT face=Arial size=2>glove type: </font><INPUT name=strGloveType value="<%=rsRA("strGloveType")%>" ></FONT></TD>
		</tr><tr>
		<TD><INPUT name=strControlFaceshield  <% If (rsRA("strControlFaceShield") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>face shield</FONT></TD>
        <TD><INPUT name=strControlRespirator  <% If (rsRA("strControlRespirator") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>respirator</FONT></TD>
		<td></td>
</TR>
<TR><TD><br/></TD></TR>
	<tr>
        <TD colspan="4"><FONT face=Arial size=2>Other safety control measures: 
		</td></tr>
		<tr><td>
		<INPUT name=strControlOther value="<%=rsRA("strControlOther")%>" size="70"></FONT></TD>	
	</tr>

</TABLE>
<br/>




<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
<TR><TD colspan=3 bgcolor = #dddddd><FONT face=Arial><STRONG>5.  Emergency Facilities</STRONG></FONT></TD></TR>
<TR><TD colspan=3><FONT face=Arial size=2>Note the emergency facilities that must be available during the work activity</FONT></TD></TR>

    <TR>
		<TD><FONT face=Arial size=2><INPUT name=strFacilitiesSpillkit  <% If (rsRA("strFacilitiesSpillkit") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>spillkit</FONT></TD>
        <TD><FONT face=Arial size=2>spill kit type:<INPUT name=strSpillkitType value="<%=rsRA("strSpillkitType")%>"></FONT></TD>
		<td></td>
	</tr><tr>
        <TD><FONT face=Arial size=2><INPUT name=strFacilitiesEyeWashStation  <% If (rsRA("strFacilitiesEyeWashStation") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>eye wash station</FONT></TD>
        <TD><FONT face=Arial size=2><INPUT name=strFacilitiesAntidoteKeptOnHand  <% If (rsRA("strFacilitiesAntidoteKeptOnHand") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>antidote kept on-hand</FONT></TD>
        <td><INPUT name=strFacilitiesHealthSurveillance  <% If (rsRA("strFacilitiesHealthSurveillance") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox><FONT face=Arial size=2>health surveillance</FONT></td>
	</TR>
    <TR>
		<TD><FONT face=Arial size=2><INPUT name=strFacilitiesFirstAidKit  <% If (rsRA("strFacilitiesFirstAidKit") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>first aid kit</FONT></TD>
        <TD><FONT face=Arial size=2><INPUT name=strFacilitiesSafetyShower  <% If (rsRA("strFacilitiesSafetyShower") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>safety shower</FONT></TD>
        <TD><FONT face=Arial size=2><INPUT name=strExposureMonitoring  <% If (rsRA("strExposureMonitoring") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>exposure monitoring</FONT></TD>
    </TR>
	<tr>
		<TD><FONT face=Arial size=2>extinguisher type:<INPUT name=strExtinguisherType  value="<%=rsRA("strExtinguisherType")%>"></FONT></TD>
        <TD><FONT face=Arial size=2><INPUT name=strFireBlanket  <% If (rsRA("strFireBlanket") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>fire blanket</FONT></TD>
		<TD><FONT face=Arial size=2><INPUT name=strFacilitiesEvacuationProcedures  <% If (rsRA("strFacilitiesEvacuationProcedures") = "on") then 
					Response.Write " CHECKED "
				END IF %> type=checkbox>evacuation/fire induction</FONT></TD>
	</tr><tr>
         <TD colspan=3><FONT face=Arial size=2>other:<INPUT name=strFacilitiesOther  value="<%=rsRA("strFacilitiesOther")%>" ></FONT></TD>
	</tr>

</TABLE>
<BR>



<TABLE border=0 cellPadding=0 cellSpacing=0 width="98%">
<TR><TD colspan=3 bgcolor = #dddddd><FONT face=Arial><STRONG>5.  Estimated Risk</STRONG></FONT></TD></TR>
<TR><TD colspan=3><FONT face=Arial size=2>The <b>estimated risk</b> is based on the <b>nature of the hazard</b> and the <b>degree of exposure</b></FONT></TD></TR>
<TR><TD colspan=3><FONT face=Arial size=2>Select the option that best describes the level of estimated risk</FONT></TD></TR>

<TR><TD>
<FONT face=Arial size=2><INPUT name=strRiskSignificant type=radio value=False <% If (rsRA("strRiskSignificant") = "False") then 
					Response.Write " CHECKED "
				END IF %>>Risks Are Not Significant</FONT><BR>
<FONT face=Arial size=2><INPUT name=strRiskControlled type=radio value=True <% If (rsRA("strRiskControlled") = "True") then 
					Response.Write " CHECKED "
				END IF %>>Risks will be Adequately Controlled</FONT><BR>
<FONT face=Arial size=2><INPUT name=strRiskSignificant type=radio value=True <% If (rsRA("strRiskSignificant") = "True") then 
					Response.Write " CHECKED "
				END IF %>>Risks are significant, since the proposed controls are not adequate (if so, repeat this assessment when the risks have been adequately controlled)</FONT><br/>
<FONT face=Arial size=2><INPUT name=strRiskControlled type=radio value=False <% If (rsRA("strRiskControlled") = "Flse") then 
					Response.Write " CHECKED "
				END IF %>>Risks are uncertain and more information required (if so, repeat this assessment when more information is obtained)</FONT><br/>
<hr>
</TD></TR>
</TABLE>

<font size ="-1">A detailed assessment may be required where complex chemical processes or exposures are involved.</font><BR>
<font size="-2">Risk assessment must be reviewed in 2 years or if the job or substance changes or new information becomes available.</font><BR>

<P>
<!-- <INPUT name=reset type=reset value="Clear form">&nbsp;&nbsp;  --><INPUT name=ADD type=submit value="Add the New Risk Assessment">
</P>
<% CleanUp() %>
</FORM>
</BODY>
</HTML>
