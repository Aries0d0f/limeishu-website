<!-- #include file="DBconn.asp" -->
<%

DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")

sql= "insert into t (a,b) values('1','2')"
if rs.State=1 then rs.Close
rs.Open sql,conn,adOpenStatic	

response.write "ok<br>"

if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>
aaaaa
