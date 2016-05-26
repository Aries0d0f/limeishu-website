<%@Language=VBScript%>
<%
'---- SchemaEnum Values ----
Const adSchemaProviderSpecific = -1
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCollations = 3
Const adSchemaColumns = 4
Const adSchemaCheckConstraints = 5
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaKeyColumnUsage = 8
Const adSchemaReferentialContraints = 9
Const adSchemaTableConstraints = 10
Const adSchemaColumnsDomainUsage = 11
Const adSchemaIndexes = 12
Const adSchemaColumnPrivileges = 13
Const adSchemaTablePrivileges = 14
Const adSchemaUsagePrivileges = 15
Const adSchemaProcedures = 16
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaProviderTypes = 22
Const adSchemaViews = 23
Const adSchemaViewColumnUsage = 24
Const adSchemaViewTableUsage = 25
Const adSchemaProcedureParameters = 26
Const adSchemaForeignKeys = 27
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29

%>
<%
T1 = trim(request("T1")) & "." & trim(request("T2"))
tblname = request.querystring("tblname")		' passes the single table name to display the table contebnts
FormName = "tech-support-asp-access-dsn.asp"
dim fldname(99)
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>MSACCESS DSN Test Script</title>
</head>

<body>

<p>This test script has been uploaded to your site by web hosting technical support because of a support request (IDS-8022157) for your hosting account.</p>
If this test script executes correctly you will see a form below which will allow you to 
Test a DSN connection to your MSAccess data base.  Once submitted you will receive a listing of the tables within the database.  Otherwise you will see an error describing the failure.
<br><hr><br>

<%
' if the form is not submiting then run the form
if Request.Form("submit") <> "TEST" then
%>


Please enter the following required information and click the TEST button
<form method="POST" action="tech-support-asp-access-dsn.asp">
<table border="0" width="400" id="table1">
  <tr>
    <td width="120" align="right">&nbsp;</td>
    <td width="270">&nbsp;</td>
  </tr>
  <tr>
    <td width="120" align="right">Account User ID </td>
    <td width="270"><input type="text" name="T1" size="20"></td>
  </tr>
  <tr>
    <td width="120" align="right">DSN to Test</td>
    <td width="270"><input type="text" name="T2" size="20"></td>
  </tr>
  <tr>
    <td width="120">&nbsp;</td>
    <td width="270"><input type="submit" value="TEST" name="Submit"></td>
  </tr>
</table>
</form>


<%
else
	Set testconn = Server.CreateObject ("ADODB.Connection")
' Open the connection to the database. System DSN used here.  Remember, userid.DSN!
	cnstr = "DSN=" & T1
'response.write cnstr
'response.flush

	testconn.Open cnstr
	Response.Write "I have connected to your Database using this DSN:  (<b><font color=#000080>" & trim(request("T2")) & "</font></b>).<BR>"
	response.write "<font color=""#800000"">"
	response.write "<b>This is the connection information used to connect:</font><br><br>"
	response.write "<font color=""#000080"" face=""Arial""><b>"
	response.write "set testconn=server.createobject(""adodb.connection"")<br>"
	response.write "Set TestRs = Server.CreateObject (""ADODB.RecordSet"")<br>"
	response.write "cnstr = ""DSN=" & T1 & """<br>"
	response.write "testconn.Open cnstr"
	response.write "</font><br><hr><br>"

  ' Open the database schema to query the list of tables. Extract the
  ' list in a Recordset object
	Set TestRs = testconn.OpenSchema (adSchemaTables)
	tempctr = 0
  ' Loop through the list and print the table names
	set testconn1=server.createobject("adodb.connection")
	
	cnstr = "DSN=" & T1 
'response.write cnstr
'response.flush
	testconn1.Open cnstr
	Do While Not TestRs.EOF
		if TestRs ("TABLE_TYPE") = "TABLE" then
		    Response.Write "<b>" & TestRs ("TABLE_NAME") & "</b>"
			' create a connection to the current table and count the records
'response.write "<hr>*" & TestRs("TABLE_NAME") & "*<hr>"
'response.flush
			SQL = "SELECT * FROM [" & TestRs("TABLE_NAME") & "]"
			set TestRs1 = server.CreateObject("adodb.recordset")
'response.write "<hr>*" & SQL  & "*<hr>"
'response.flush
			TestRs1.Open SQL, testconn1
			cntr=0
			do until TestRs1.EOF
				cntr = cntr +1
				testrs1.MoveNext
			loop
			Response.Write "     <font color=#800000>Number of records</font> - "
			Response.Write cntr ' & "    <a href=" & FormName & "?R1=dsn&T1=" & T1 & "&con=list&B1=Submit&tblname=" & TestRs ("TABLE_NAME") & ">View Data</a>"
			Response.Write "<br>"
	TestRs1.Close
	Set TestRs1 = Nothing
		
	 end if
    TestRs.MoveNext
	Loop

	testconn1.Close
	Set testconn1 = Nothing
  ' Close and destroy the recordset and connection objects
	TestRs.Close
	Set TestRs = Nothing
	

  testconn.Close
  Set testconn = Nothing

%>
<table>
  <tr><td>

<% 'end if
' ***************************************************************************************
' Display data for the specified table (first get field list then display data)
' ***************************************************************************************
if con = "list" then
%>  
</td></tr></table>
<table><tr><td>
<B><%=tblname%></B> Data:
</td></tr></table>
<table>
  <tr><td>
<table><tr><td>

</td></tr></table>
<table bgcolor=LemonChiffon border=1><tr>
</td></tr></table>

<%
  End IF
%>
</td></tr></table>  
<%
if err.number <> 0 then
Response.write err.number & "<br>"
Response.write err.Description & "<br>"
Response.write err.Source & "<br>"
end if
%>
<%end if%>


<br><hr><br><br>

<b>List of common errors:</b><br><br>These are provided for troubleshooting in case you get any errors listed above.<br><br>

<table width="75%" border="0" align="center" id="table2">
  <tr bgcolor="#CCCCCC"> 
    <td width="10%" valign="top"><b>Error:</b></td>
    <td width="90%">General error Unable to open registry key 'Temporary 
	(volatile) Jet DSN<br>
    </td>
  </tr>
  <tr> 
    <td width="10%" valign="top"><b>Cause:</b></td>
    <td width="90%"> 
      Could be a FILE permissions error but most likely a empty or corrupt 
		database. This error is due to not being able to read database header 
		information.</td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr>
    </td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td width="10%" valign="top"><b>Error:</b></td>
    <td width="90%"> <font face="Arial" size="2">Data source name not found and 
	no default driver specified</font> <br>
    </td>
  </tr>
  <tr> 
    <td width="10%" valign="top"><b>Cause:</b></td>
    <td width="90%"> 
      The DSN is not valid. The DSN may not be created yet or is corrupt. Check 
		NT Tran log for when this DSN was created. If tran not in log then the 
		tran could be hanging and will need to be escalated.</td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr>
    </td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td width="10%" valign="top"><b>Error:</b></td>
    <td width="90%"> <font face="Arial" size="2">Could not find file 
	'(unknown)'.</font><br>
    </td>
  </tr>
  <tr> 
    <td width="10%" valign="top"><b>Cause:</b></td>
    <td width="90%"> 
      The file that the DSN is looking for does not exist. Verify that the path 
		and file name is correct in the DSN.</td>
  </tr>
</table>


<table width="75%" border="0" align="center" id="table3">
  <tr bgcolor="#CCCCCC"> 
    <td width="10%" valign="top"><b>Error:</b></td>
    <td width="90%">Please use updateable script<br>
    </td>
  </tr>
  <tr> 
    <td width="10%" valign="top"><b>Cause:</b></td>
    <td width="90%"> 
      Could be a FILE permissions error but most likely a coding issue. This 
		error is due to not being able to write data to the database. reset 
		perms and then test. If still giving the same error then it will be a 
		coding issue..</td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr>
    </td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td width="10%" valign="top"><b>Error:</b></td>
    <td width="90%"> Selected collating sequence not supported by the operating 
	system. <br>
    </td>
  </tr>
  <tr> 
    <td width="10%" valign="top"><b>Cause:</b></td>
    <td width="90%"> 
      This could be an issue where the specified language does not exist on the 
		server or a bad DSN.</td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <hr>
    </td>
  </tr>
  <tr bgcolor="#CCCCCC"> 
    <td width="10%" valign="top"><b>Error:</b></td>
    <td width="90%"> <b><font face="Arial" size="2">Unspecified error</font></b><br>
    </td>
  </tr>
  <tr> 
    <td width="10%" valign="top"><b>Cause:</b></td>
    <td width="90%"> 
      This is a general problem opening the database. Try removing the mdb file 
		and / or running a kill site.</td>
  </tr>
</table>


</body>

</html>
