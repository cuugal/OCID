 <%
 '
    ' Export Recordset to CSV
    ' http://salman-w.blogspot.com/2009/07/export-recordset-data-to-csv-using.html
    '
    ' This sub-routine Response.Writes the content of an ADODB.RECORDSET in CSV format
    ' The function closely follows the recommendations described in RFC 4180:
    ' Common Format and MIME Type for Comma-Separated Values (CSV) Files
    ' http://tools.ietf.org/html/rfc4180
    '
    ' @RS: A reference to an open ADODB.RECORDSET object
    '
	Response.ContentType = "text/csv"

	Response.AddHeader "Content-Disposition", "attachment;filename=export.csv"

	dim query
	dim strSQL
	dim constr
	query = request.form("data")
	strSQL = request.form("sql")
	
	constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("data/Chemicals.mdb")
	
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open constr

	set RS = Server.CreateObject("ADODB.Recordset")
	'response.write(strSQL)
	RS.Open strSQL, conn


    if RS.EOF then
    
        '
        ' There is no data to be written
        '
        response.end
    
    end if

    dim RX
    set RX = new RegExp
        RX.Pattern = "\r|\n|,|"""

    dim i
    dim Field
    dim Separator

    '
    ' Writing the header row (header row contains field names)
    '

    Separator = ""
    for i = 0 to RS.Fields.Count - 1
        Field = RS.Fields(i).Name
        if RX.Test(Field) then
            '
            ' According to recommendations:
            ' - Fields that contain CR/LF, Comma or Double-quote should be enclosed in double-quotes
            ' - Double-quote itself must be escaped by preceeding with another double-quote
            '
            Field = """" & Replace(Field, """", """""") & """"
        end if
        Response.Write Separator & Field
        Separator = ","
    next
    Response.Write vbNewLine

    '
    ' Writing the data rows
    '

    do until RS.EOF
        Separator = ""
        for i = 0 to RS.Fields.Count - 1
            '
            ' Note the concatenation with empty string below
            ' This assures that NULL values are converted to empty string
            '
            Field = RS.Fields(i).Value & ""
            if RX.Test(Field) then
                Field = """" & Replace(Field, """", """""") & """"
            end if
            Response.Write Separator & Field
            Separator = ","
        next
        Response.Write vbNewLine
        RS.MoveNext
    loop
	%>