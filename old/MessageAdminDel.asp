<!-- #include file="DBconn.asp" -->
<%
if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")

if request("c_Message")<>"" then 
	sql=" delete  from  Message where c_Message=" & request("c_Message")
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic
end if
response.redirect "MessageAdmin.asp"

if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>

