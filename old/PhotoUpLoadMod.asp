<!-- #include file="DBconn.asp" -->
<%
if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")
c_Photo=request("c_photo")
if request("c_photo")<>"" then 
	sql=" select * from Photo where c_Photo=" & request("c_photo")
	if rs.State=1 then rs.Close
	rs.Open sql,conn,adOpenStatic	
	if not rs.eof then
		c_Type=rs("c_Type")
		c_FileName=rs("c_FileName")		
		c_TypeEx=rs("c_TypeEx")	
		c_Width=rs("c_Width")	
		c_Height=rs("c_Height")	
		c_Year=rs("c_Year")	
		c_Name=rs("c_Name")	
	end if
end if
	

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
		<FORM METHOD="POST" ENCTYPE="multipart/form-data" id="Form1" action="PhotoUpLoadM.asp">
			
			  <font face="�s�ө���"><BR>
				<BR>
		  </font>
			  <TABLE width="780" border="1" align="center" cellPadding="1" borderColor="#999999" borderColorDark="#ffffff"
					bgColor="white" id="Table1">
				  <TR>
					  <TD align="center" bgColor="#edeeec" colSpan="2"><font face="�s�ө���">�@�~-�ק���</font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px; HEIGHT: 15px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">���O</font></div></TD>
				    <TD style="HEIGHT: 15px">
								
							    <font face="�s�ө���">
								  <select name="c_Type" id="c_Type">
							        <option value="1" <% if c_Type=1 then response.write "selected" %> >�e�@�@�~</option>
							        <option value="2" <% if c_Type=2 then response.write "selected" %> >�ѫH�@�~</option>
							        <option value="3" <% if c_Type=3 then response.write "selected" %>>�Ӥ����A</option>
							        <option value="4" <% if c_Type=4 then response.write "selected" %>>�媫�@�~</option>
                                  </select>
								  <input name="c_Photo" type="hidden" id="c_Photo" value="<%=c_Photo%>">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">����&nbsp;</font></div></TD>
					  <TD>
	
							    <font face="�s�ө���">
								  <select name="c_TypeEx" id="c_TypeEx">
							        <option value="1" <% if c_TypeEx=1 then response.write "selected" %> >�o�e</option>
							        <option value="2" <% if c_TypeEx=2 then response.write "selected" %> >����</option>
							        <option value="3" <% if c_TypeEx=3 then response.write "selected" %> >�J��</option>
							        <option value="4" <% if c_TypeEx=4 then response.write "selected" %> >�r�e</option>
							        <option value="5" <% if c_TypeEx=5 then response.write "selected" %> >�ѫH</option>
							        <option value="6" <% if c_TypeEx=6 then response.write "selected" %> >���y</option>
							        <option value="8" <% if c_TypeEx=8 then response.write "selected" %> >���m</option>									
							        <option value="7" <% if c_TypeEx=7 then response.write "selected" %> >��L</option>
						        </select>
                        </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">�@�~�W��&nbsp;</font></div></TD>
					  <TD><font face="�s�ө���">
					  <input name="c_Name" type="text" id="c_Name" value="<%=c_Name%>" size="70">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">�@�~�~��&nbsp;</font></div></TD>
					  <TD><font face="�s�ө���">
					  <input name="c_Year" type="text" id="c_Year" value="<%=c_Year%>">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">�e</font></div></TD>
				    <TD><font face="�s�ө���">
				    <input name="c_Width" type="text" id="c_Width" value="<%=c_Width%>">
				    </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec">
			        <div align="right"><font size="2" face="�s�ө���">��</font></div></TD>
				    <TD style="HEIGHT: 28px"><font face="�s�ө���">
				    <input name="c_Height" type="text" id="c_Height" value="<%=c_Width%>">
				    </font></TD>
				  </TR>
				  <TR>
				    <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec"><font size="2" face="�s�ө���">��@�~�Ϥ�</font></TD>
				    <TD style="HEIGHT: 28px"><font face="�s�ө���"><img src="db/<%=c_FileName%>" height="80"></font></TD>
			    </TR>
				  <TR>
                    <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec">
                      <div align="right"><font face="�s�ө���">�Ϥ����|</font></div></TD>
                    <TD style="HEIGHT: 28px"><font face="�s�ө���">
                    <INPUT name="c_FileName" type="file" id="c_FileName" size="70">
                    </font></TD>
			    </TR>
					  <TR>
						  <TD colspan="2" align="right" >
						    <div align="center">							  <font face="�s�ө���">&nbsp;
					        <input type="submit" name="Submit2" value="�ק�W��">
					        </font></div></TD>
					  </TR>
		      </TABLE>
			  <P align="center">
          </P>
				<div align="center"><a href="FuncList.asp">�^�\��M��</a></div>
        </form>
	</body>
</HTML>
