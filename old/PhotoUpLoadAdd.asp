
<%
	if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >

<HTML>
	<HEAD>
		<title>PhotoUpload</title>
		<meta http-equiv="Content-Type" content="text/html; charset=big5">	
	</HEAD>
	<body bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<FORM METHOD="POST" ENCTYPE="multipart/form-data" id="Form1" action="PhotoUpLoadA.asp">
			
			  <font face="�s�ө���"><BR>
				<BR>
		  </font>
			  <TABLE width="780" border="1" align="center" cellPadding="1" borderColor="#999999" borderColorDark="#ffffff"
					bgColor="white" id="Table1">
				  <TR>
				    <TD align="center" bgColor="#edeeec" colSpan="2"><font face="�s�ө���">�@�~�W�� </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px; HEIGHT: 15px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">���O</font></div></TD>
					  <TD style="HEIGHT: 15px">
								
							    <font face="�s�ө���">
								  <select name="c_Type" id="c_Type">
							        <option value="1" selected>�e�@�@�~</option>
							        <option value="2">�ѫH�@�~</option>
							        <option value="3">�Ӥ����A</option>
							        <option value="4">�媫�@�~</option>
                                  </select>
					    </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">���� &nbsp;</font></div></TD>
					  <TD>
	
							    <font face="�s�ө���">
								  <select name="c_TypeEx" id="c_TypeEx">
							        <option value="1" selected>�o�e</option>
							        <option value="2">����</option>
							        <option value="3">�J��</option>
							        <option value="4">�r�e</option>
							        <option value="5">�ѫH</option>
							        <option value="6">���y</option>
							        <option value="8">���m</option>									
							        <option value="7">��L</option>
						        </select>
                        </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">�@�~�W��&nbsp;</font></div></TD>
					  <TD><font face="�s�ө���">
					  <input name="c_Name" type="text" id="c_Name" size="70">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">�@�~�~��&nbsp;</font></div></TD>
					  <TD><font face="�s�ө���">
					  <input name="c_Year" type="text" id="c_Year">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="�s�ө���">�e&nbsp;</font></div></TD>
				    <TD><font face="�s�ө���">
				    <input name="c_Width" type="text" id="c_Width">
				    </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec">
			        <div align="right"><font size="2" face="�s�ө���">��</font></div></TD>
				    <TD style="HEIGHT: 28px"><font face="�s�ө���">
				    <input name="c_Height" type="text" id="c_Height">
				    </font></TD>
				  </TR>
				  <TR>
                    <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec">
                    <div align="right"><font size="2" face="�s�ө���">�Ϥ����|</font></div></TD>
                    <TD style="HEIGHT: 28px"><font face="�s�ө���">
                    <INPUT name="c_FileName" type="file" id="c_FileName" size="70">
                    </font></TD>
			    </TR>
					  <TR>
						  <TD colspan="2" align="right" >
						    <div align="center">							  <font face="�s�ө���">&nbsp;
					        <input type="submit" name="Submit2" value="�W��">
					        </font></div></TD>
					  </TR>
		      </TABLE>
			  <P align="center">
          </P>
				<div align="center"><font face="�s�ө���"><a href="FuncList.asp">�^�\��M��</a></font>
		  </div>
				<div align="center"></div>
    </form>
	</body>
</HTML>
