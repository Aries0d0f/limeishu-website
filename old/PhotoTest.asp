<!-- #include file="DBconn.asp" -->
<%
'if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")


	sql=" insert into Photo(c_Photo,c_FileName,c_Type,c_TypeEx,c_Width,c_Height,c_Year,c_Name) values ("
	sql=sql & "1000"
	sql=sql & ",'11'"
	sql=sql & ",1"
	sql=sql & ",'11'"	
	sql=sql & ",'11'"	
	sql=sql & ",'11'"	
	sql=sql & ",'11'"	
	sql=sql & ",'11'"	
	
	sql=sql & ")"
	response.write sql
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic	



if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>