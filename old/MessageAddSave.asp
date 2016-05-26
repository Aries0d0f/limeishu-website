<!-- #include file="DBconn.asp" -->
<%
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")

	sql="select max(c_Message) as c_MessageKey from Message"
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic		
	if not rs.eof then
		c_Message=rs("c_MessageKey")+1
	end if
	if isnull(c_Message) then c_Message=1
	
	sql=" insert into Message (c_Message,c_Title,c_User,c_Mail,c_City,c_Content,c_DateTime) values ("
'	sql=sql & newkey
	sql=sql & c_Message 
	sql=sql & ",'" &  request("c_Title")& "'" 	
	sql=sql & ",'" &  request("c_User")  & "'"	
	sql=sql & ",'" &  request("c_Mail")  & "'"	
	sql=sql & ",'" &  request("c_City")  & "'"	
	sql=sql & ",'" &  replace(request("c_Content"),"'","`")  & "'"	
	sql=sql & ",'" &  now()  & "'"		
	sql=sql & ")"
	'response.write sql
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic	

reStr="Message.asp"
response.redirect reStr
 


if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>