<!-- #include file="DBconn.asp" -->
<%
if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")
rs.CursorLocation=3
rs.PageSize=10
page=request("Page")
if page="" then page=1		

sql=" select * from  Message order by c_DateTime desc "
if rs.State=1 then rs.Close
rs.Open sql,conn,adOpenStatic

	pageNum=" 頁次 \ "
	for i=1 to rs.PageCount 
		if i=int(page) then 
			pageNum= pageNum & "[" & i & "]"
		else
			pageNum= pageNum & " <A href=MessageAdmin.asp?page=" & i
			pageNum= pageNum & ">[" & i & "]</A> " 

		end if
	next
	pageNum=pageNum 

if not rs.eof then
rs.AbsolutePage=page 

list=""
rowCount=1
while not rs.eof  and rowCount<=rs.PageSize
	rowCount=rowCount+1
	list=list& " <TR bgColor='#cadde1'>" & chr(13)
	list=list& " <TD height=26><B>"&rs("c_Title")&" </B></TD>" & chr(13)
	list=list& " <TD align='center'><A href='mailto:"&rs("c_Mail")&" '> "&rs("c_User")&"  </A></TD> " & chr(13)
	list=list& " <TD align='center' height=26>"&rs("c_Datetime")&"</TD>" & chr(13)	
	list=list& " <TD align='center' height=26><a href='MessageAdminDel.asp?c_Message="&rs("c_Message") &"'>刪除</a></TD>" & chr(13)		
	list=list& " </TR>" & chr(13)
	list=list& "  <TR bgColor='#ffffff'>" & chr(13)
	list=list& " <TD colSpan=4> "&rs("c_Content")&"&nbsp; </TD>	" & chr(13)
	rs.movenext
	
wend

end if

if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>


<TABLE width="780" border="0" align="center" cellPadding="5" cellSpacing="1" bgColor="#999999" id="Table5">
  <TR bgColor="#FFFFFF">
    <TD height="12" colspan="4" align="center"><div align="left"><%=pageNum%></div></TD>
  </TR>
  <TR bgColor="#669999">
    <TD align="center" width="40%"><FONT color="#ffffff">主題</FONT></TD>
    <TD align="center" width="23%"><FONT color="#ffffff">發表人</FONT></TD>
    <TD align="center" width="26%" height="12"><FONT color="#ffffff">張貼時間</FONT></TD>
    <TD width="11%" align="center" bgcolor="#669999"><font color="#FFFFFF">刪除</font></TD>
  </TR>
  <%=list%>
  <TR bgColor="#FFFFFF">
    <TD height="12" colspan="4" align="center"><div align="left"><%=pageNum%></div></TD>
  </TR>  
</TABLE>
<p>&nbsp;</p>
<p align="center"><font face="新細明體"><a href="FuncList.asp">回功能清單</a></font> </p>
