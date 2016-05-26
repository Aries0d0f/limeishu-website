<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>MSSQL SERVER CONNECTION TEST</title>

</head>

<body>
<%

sqlserver = "sql700.mssqlservers.com"
userid = "limeis"
passwd = "aq7kk+1207"

'---- SchemaEnum Values ----  This section is required to read table names
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
'---- END SchemaEnum Values ---- 


set testconn=server.createobject("adodb.connection")
	set getservconn=server.createobject("adodb.connection")
	Set TestRs = Server.CreateObject ("ADODB.RecordSet")
%>

 
<div align="center">
  <center>
   
<br>
<table border="1" width="500" bgcolor="#CCCCCC" style="border-collapse: collapse" bordercolor="#111111" cellpadding="5" cellspacing="5"><tr><td>
<%testdsn="Driver={SQL Server};Server=" & sqlserver & ";Database=" & userid & ";UID=" & userid & ";PWD=" & passwd
' Open the connection to the database. use database selected
testconn.Open testdsn
Response.Write "I have connected to your database, user tables and views are listed below:<BR>"
%>
<center><b><u><font color="#000080">
User Tables, and Views </font></u></b><br></center>
<%' on error resume next
  ' Open the database schema to query the list of tables. Extract the
  ' list in a Recordset object
  Set TestRs = testconn.OpenSchema (adSchemaTables)
 Set RS_item = testconn.OpenSchema (adSchemaTables)
  ' Loop through the list and print the table names
  Do While Not TestRs.EOF
    if TestRs ("TABLE_TYPE") = "SYSTEM TABLE" then
    else
	  if TestRs ("TABLE_TYPE") = "TABLE" then
	    Response.Write "<BR><b>" & TestRs ("TABLE_NAME") 
	    Response.Write "&nbsp;&nbsp;-&nbsp;&nbsp;</b>" & TestRs ("TABLE_TYPE")
	  
	  else
	    Response.Write "<BR>" & TestRs ("TABLE_NAME") 
	    Response.Write "&nbsp;&nbsp;-&nbsp;&nbsp;" & TestRs ("TABLE_TYPE")
	  end if
	 end if

  TestRs.MoveNext
  Loop

  ' Close and destroy the recordset and connection objects
  TestRs.Close
  Set TestRs = Nothing

  testconn.Close
  Set testconn = Nothing


%>
</td></tr></table>

</body>
</html>