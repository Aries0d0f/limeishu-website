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
	if rs_c_Type=1 then c_TypeStr="�o�e�@�~"
	if rs_c_Type=2 then c_TypeStr="�ѫH�@�~"
	if rs_c_Type=3 then c_TypeStr="�Ӥ����A"
	if rs_c_Type=4 then c_TypeStr="�媫�@�~"			
	rstr=rstr & "<td height=24><font size=2>" & c_TypeStr & "&nbsp;</font></td>" & chr(13)	
	
	rs_c_TypeEx=rs("c_TypeEx") 
	if rs_c_TypeEx=1 then c_TypeExStr="�o�e"
	if rs_c_TypeEx=2 then c_TypeExStr="����"
	if rs_c_TypeEx=3 then c_TypeExStr="�J��"
	if rs_c_TypeEx=4 then c_TypeExStr="�r�e"
	if rs_c_TypeEx=5 then c_TypeExStr="�ѫH"
	if rs_c_TypeEx=6 then c_TypeExStr="���y"
	if rs_c_TypeEx=8 then c_TypeExStr="���m"	
	if rs_c_TypeEx=7 then c_TypeExStr="��L"	
	rstr=rstr & "<td><font size=2>" & c_TypeExStr & "&nbsp;</font></td>" & chr(13)	
	
	rstr=rstr & "<td><font size=2>" & rs("c_Name") & "&nbsp;</font></td>" & chr(13)
	rstr=rstr & "<td><font size=2>" & rs("c_Year") & "&nbsp;</font></td>" & chr(13)
	rstr=rstr & "<td align=center><font  size=2>  <a href='db/" & rs("c_FileName") & "' target='_blank' broder=0 > <img src='db/" & rs("c_FileName") & "' height=80  border='0' ></a></font></td>" & chr(13)
	rstr=rstr & "<td align=center><div align=center><font size=2> " & chr(13)
	rstr=rstr & "<a href='PhotoUploadMod.asp?c_Photo="& rs("c_Photo")  &"'> �ק� </a>" & chr(13)
	rstr=rstr & "/ <a href='PhotoUploadDel.asp?c_Photo="& rs("c_Photo")  &"'> �R�� </a> </font></div></td>" & chr(13)
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
			<FONT face="�s�ө���">				<BR>
				<TABLE width="780" border="1" align="center" cellPadding="0" cellspacing="0" bordercolor="#000000" borderColorDark="#ffffff" 
					 id="Table1">
					<TR>
						<TD align="center" bgColor="#edeeec" colSpan="2"><font face="�s�ө���">�@�~</font>�W��
-��ƲM��							</TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px; HEIGHT: 15px" align="right" bgColor="#edeeec"><div align="right"><font size="2">���O</font></div></TD>
						<TD style="HEIGHT: 15px">						  <font size="2">
					    <select name="c_Type" id="c_Type">
				          <option value="1" <% if c_Type=1 then response.write "selected" %>>�e�@�@�~</option>
				          <option value="2" <% if c_Type=2 then response.write "selected" %>>�ѫH�@�~</option>
				          <option value="3" <% if c_Type=3 then response.write "selected" %>>�Ӥ����A</option>
				          <option value="4" <% if c_Type=4 then response.write "selected" %>>�媫�@�~</option>
                        </select>
                        </font></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">����</font><font size="2">&nbsp;</font></div></TD>
						<TD>
	
								  <font size="2">
								  <select name="c_TypeEx" id="c_TypeEx">
							        <option value="1" <% if c_TypeEx=1 then response.write "selected" %>>�o�e</option>
							        <option value="2" <% if c_TypeEx=2 then response.write "selected" %>>����</option>
							        <option value="3" <% if c_TypeEx=3 then response.write "selected" %>>�J��</option>
							        <option value="4" <% if c_TypeEx=4 then response.write "selected" %>>�r�e</option>
							        <option value="5" <% if c_TypeEx=5 then response.write "selected" %>>�ѫH</option>
							        <option value="6" <% if c_TypeEx=6 then response.write "selected" %>>���y</option>
							        <option value="8" <% if c_TypeEx=8 then response.write "selected" %>>���m</option>									
							        <option value="7" <% if c_TypeEx=7 then response.write "selected" %>>��L</option>
						          </select>
                                  </font></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2">�@�~�W������r&nbsp;</font></div></TD>
						<TD><font size="2">
					    <input name="c_Name" type="text" id="c_Name" value="<%=c_Name%>" size="70">
						  </font></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2">�@�~�~��&nbsp;����r</font></div></TD>
						<TD><font size="2">
					    <input name="c_Year" type="text" id="c_Year" value="<%=c_Year%>">
						  </font></TD>
					</TR>
						<TR>
							<TD colspan="2" align="right" >
							  <div align="center">							  <font size="2">&nbsp;
						      <input type="submit" name="Submit2" value="    �d    ��    ">
					      </font></div></TD>
						</TR>
			</TABLE>
	      </FONT>
			<table width="780" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#000000" bordercolorlight="#FFFFFF">
              <tr>
                <td height="27" colspan="5"><font size="2">&nbsp;</font></td>
                <td><div align="center"><font size="2"><a href="PhotoUploadAdd.asp">�s�W�@�~</a></font></div></td>
              </tr>
              <tr>
                <td width="73" height="27"><div align="center"><font size="2">���O</font></div></td>
                <td width="79"><div align="center"><font face="�s�ө���">����</font></div></td>
                <td width="308"><div align="center"><font size="2">�@�~�W��</font></div></td>
                <td width="59"><div align="center"><font size="2">�@�~�~��</font></div></td>
                <td width="157"><div align="center"><font size="2">�@�~�Ϥ�</font></div></td>
                <td width="90"><div align="center"><font size="2">�\��</font></div></td>
              </tr>

			  <%=rstr%>
            </table>
		  <P align="center">
          </P>
			<div align="center"><font face="�s�ө���"><a href="FuncList.asp">�^�\��M��</a></font>
		  </div>
			<div align="center"></div>
        </form>
	</body>
</HTML>
