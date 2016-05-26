<!-- #include file="DBconn.asp" -->
<%
if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")
	
	sql="select max(c_photo) as newkey from Photo "
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic	
	newkey=rs("newkey")
	if isnull(newkey) then 
		newkey=1
	else 
		newkey=rs("newkey")+1
	end if	
	
	sql=" insert into Photo(c_photo,c_FileName,c_Type,c_TypeEx,c_Width,c_Height,c_Year,c_Name) values ("
	sql=sql & newkey
	sql=sql & ",'" & request("c_FileName") & "'"
	sql=sql & ",'" &  request("c_Type")& "'" 	
	sql=sql & ",'" &  request("c_TypeEx")  & "'"	
	sql=sql & ",'" &  request("c_Width")  & "'"	
	sql=sql & ",'" &  request("c_Height")  & "'"	
	sql=sql & ",'" &  request("c_Year")  & "'"	
	sql=sql & ",'" &  request("c_Name")  & "'"	
	
	sql=sql & ")"
	'response.write sql
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic	

reStr="PhotoUpLoadList.asp?"
reStr=reStr & "c_FileName="&request("c_FileName")
reStr=reStr & "&c_Type="&request("c_Type")
reStr=reStr & "&c_TypeEx="&request("c_TypeEx")
reStr=reStr & "&c_Width="&request("c_Width")
reStr=reStr & "&c_Height="&request("c_Height")
'reStr=reStr & "&c_Year="&request("c_Year")
'reStr=reStr & "&c_Name="& request("c_Name") 

response.redirect reStr
 


if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>