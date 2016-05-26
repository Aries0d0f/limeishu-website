<!-- #include file="DBconn.asp" -->
<%
'if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")

	sql=" update  Photo set "
	sql=sql & "c_Type='" &  request("c_Type")	 & "'"	
	if 	rtrim(request("c_FileName"))<>"" then sql=sql & ",c_FileName='" & request("c_FileName") & "'"
	sql=sql & ",c_TypeEx='" &  request("c_TypeEx")  & "'"	
	sql=sql & ",c_Width='" &  request("c_Width")  & "'"	
	sql=sql & ",c_Height='" &  request("c_Height")  & "'"	
	sql=sql & ",c_Year='" &  request("c_Year")  & "'"	
	sql=sql & ",c_Name='" &  request("c_Name")  & "'"	
	
	sql=sql & " where  c_Photo=" & request("c_Photo")
	'response.write sql
	'response.End()
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