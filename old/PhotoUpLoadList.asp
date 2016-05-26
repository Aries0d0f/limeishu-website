<!-- #include file="DBconn.asp" -->
<%
if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")

c_Type=rtrim(request("c_Type"))
c_TypeEx=rtrim(request("c_TypeEx"))
c_Name=rtrim(request("c_Name"))
c_Year=rtrim(request("c_Year"))

if c_Type="" then 
	c_Type=1
else 
	c_Type=int(c_Type)
end if
if c_TypeEx="" then
	 c_TypeEx=1
else 
	c_TypeEx=int(c_TypeEx)
end if

sql = "select * from Photo  where c_Type=" & c_Type & " and c_TypeEx=" & c_TypeEx
if  c_Name<>"" then 
	sql=sql & " and c_Name like '%"& c_Name  &"%'"
end if
if  c_Year<>"" then 
	sql=sql & " and c_Year like '%"& c_Year  &"%'"
end if
sql=sql & " order by c_Name,c_Year "
'response.write sql
'response.End()

if rs.State=1 then rs.Close
rs.Open sql,conn,adOpenStatic	


rstr=""

while not rs.eof 
	rstr=rstr & "<tr>" & chr(13)
	
	rs_c_Type=rs("c_Type") 
	if rs_c_Type=1 then c_TypeStr="油畫作品"
	if rs_c_Type=2 then c_TypeStr="書信作品"
	if rs_c_Type=3 then c_TypeStr="照片集錦"
	if rs_c_Type=4 then c_TypeStr="文物作品"			
	rstr=rstr & "<td height=24><font size=2>" & c_TypeStr & "&nbsp;</font></td>" & chr(13)	
	
	rs_c_TypeEx=rs("c_TypeEx") 
	if rs_c_TypeEx=1 then c_TypeExStr="油畫"
	if rs_c_TypeEx=2 then c_TypeExStr="水墨"
	if rs_c_TypeEx=3 then c_TypeExStr="雕刻"
	if rs_c_TypeEx=4 then c_TypeExStr="字畫"
	if rs_c_TypeEx=5 then c_TypeExStr="書信"
	if rs_c_TypeEx=6 then c_TypeExStr="素描"
	if rs_c_TypeEx=8 then c_TypeExStr="水彩"	
	if rs_c_TypeEx=7 then c_TypeExStr="其他"	
	rstr=rstr & "<td><font size=2>" & c_TypeExStr & "&nbsp;</font></td>" & chr(13)	
	
	rstr=rstr & "<td><font size=2>" & rs("c_Name") & "&nbsp;</font></td>" & chr(13)
	rstr=rstr & "<td><font size=2>" & rs("c_Year") & "&nbsp;</font></td>" & chr(13)
	rstr=rstr & "<td align=center><font  size=2>  <a href='db/" & rs("c_FileName") & "' target='_blank' broder=0 > <img src='db/" & rs("c_FileName") & "' height=80  border='0' ></a></font></td>" & chr(13)
	rstr=rstr & "<td align=center><div align=center><font size=2> " & chr(13)
	rstr=rstr & "<a href='PhotoUploadMod.asp?c_Photo="& rs("c_Photo")  &"'> 修改 </a>" & chr(13)
	rstr=rstr & "/ <a href='PhotoUploadDel.asp?c_Photo="& rs("c_Photo")  &"'> 刪除 </a> </font></div></td>" & chr(13)
	rstr=rstr & "</tr>" & chr(13)

	rs.movenext
wend 
	
	

if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >

<HTML>
	<HEAD>
		<title>PhotoUpload</title>
		<meta http-equiv="Content-Type" content="text/html; charset=big5">	
	</HEAD>
	<body bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<FORM  id="Form1" action="PhotoUpLoadList.asp">
			<FONT face="新細明體">				<BR>
				<TABLE width="780" border="1" align="center" cellPadding="0" cellspacing="0" bordercolor="#000000" borderColorDark="#ffffff" 
					 id="Table1">
					<TR>
						<TD align="center" bgColor="#edeeec" colSpan="2"><font face="新細明體">作品</font>上傳
-資料清單							</TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px; HEIGHT: 15px" align="right" bgColor="#edeeec"><div align="right"><font size="2">類別</font></div></TD>
						<TD style="HEIGHT: 15px">						  <font size="2">
					    <select name="c_Type" id="c_Type">
				          <option value="1" <% if c_Type=1 then response.write "selected" %>>畫作作品</option>
				          <option value="2" <% if c_Type=2 then response.write "selected" %>>書信作品</option>
				          <option value="3" <% if c_Type=3 then response.write "selected" %>>照片集錦</option>
				          <option value="4" <% if c_Type=4 then response.write "selected" %>>文物作品</option>
                        </select>
                        </font></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="新細明體">項目</font><font size="2">&nbsp;</font></div></TD>
						<TD>
	
								  <font size="2">
								  <select name="c_TypeEx" id="c_TypeEx">
							        <option value="1" <% if c_TypeEx=1 then response.write "selected" %>>油畫</option>
							        <option value="2" <% if c_TypeEx=2 then response.write "selected" %>>水墨</option>
							        <option value="3" <% if c_TypeEx=3 then response.write "selected" %>>雕刻</option>
							        <option value="4" <% if c_TypeEx=4 then response.write "selected" %>>字畫</option>
							        <option value="5" <% if c_TypeEx=5 then response.write "selected" %>>書信</option>
							        <option value="6" <% if c_TypeEx=6 then response.write "selected" %>>素描</option>
							        <option value="8" <% if c_TypeEx=8 then response.write "selected" %>>水彩</option>									
							        <option value="7" <% if c_TypeEx=7 then response.write "selected" %>>其他</option>
						          </select>
                                  </font></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2">作品名稱關鍵字&nbsp;</font></div></TD>
						<TD><font size="2">
					    <input name="c_Name" type="text" id="c_Name" value="<%=c_Name%>" size="70">
						  </font></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2">作品年份&nbsp;關鍵字</font></div></TD>
						<TD><font size="2">
					    <input name="c_Year" type="text" id="c_Year" value="<%=c_Year%>">
						  </font></TD>
					</TR>
						<TR>
							<TD colspan="2" align="right" >
							  <div align="center">							  <font size="2">&nbsp;
						      <input type="submit" name="Submit2" value="    查    詢    ">
					      </font></div></TD>
						</TR>
			</TABLE>
	      </FONT>
			<table width="780" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#000000" bordercolorlight="#FFFFFF">
              <tr>
                <td height="27" colspan="5"><font size="2">&nbsp;</font></td>
                <td><div align="center"><font size="2"><a href="PhotoUploadAdd.asp">新增作品</a></font></div></td>
              </tr>
              <tr>
                <td width="73" height="27"><div align="center"><font size="2">類別</font></div></td>
                <td width="79"><div align="center"><font face="新細明體">項目</font></div></td>
                <td width="308"><div align="center"><font size="2">作品名稱</font></div></td>
                <td width="59"><div align="center"><font size="2">作品年份</font></div></td>
                <td width="157"><div align="center"><font size="2">作品圖片</font></div></td>
                <td width="90"><div align="center"><font size="2">功能</font></div></td>
              </tr>

			  <%=rstr%>
            </table>
		  <P align="center">
          </P>
			<div align="center"><font face="新細明體"><a href="FuncList.asp">回功能清單</a></font>
		  </div>
			<div align="center"></div>
        </form>
	</body>
</HTML>
