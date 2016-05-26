
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
			
			  <font face="新細明體"><BR>
				<BR>
		  </font>
			  <TABLE width="780" border="1" align="center" cellPadding="1" borderColor="#999999" borderColorDark="#ffffff"
					bgColor="white" id="Table1">
				  <TR>
				    <TD align="center" bgColor="#edeeec" colSpan="2"><font face="新細明體">作品上傳 </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px; HEIGHT: 15px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="新細明體">類別</font></div></TD>
					  <TD style="HEIGHT: 15px">
								
							    <font face="新細明體">
								  <select name="c_Type" id="c_Type">
							        <option value="1" selected>畫作作品</option>
							        <option value="2">書信作品</option>
							        <option value="3">照片集錦</option>
							        <option value="4">文物作品</option>
                                  </select>
					    </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="新細明體">項目 &nbsp;</font></div></TD>
					  <TD>
	
							    <font face="新細明體">
								  <select name="c_TypeEx" id="c_TypeEx">
							        <option value="1" selected>油畫</option>
							        <option value="2">水墨</option>
							        <option value="3">雕刻</option>
							        <option value="4">字畫</option>
							        <option value="5">書信</option>
							        <option value="6">素描</option>
							        <option value="8">水彩</option>									
							        <option value="7">其他</option>
						        </select>
                        </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="新細明體">作品名稱&nbsp;</font></div></TD>
					  <TD><font face="新細明體">
					  <input name="c_Name" type="text" id="c_Name" size="70">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="新細明體">作品年份&nbsp;</font></div></TD>
					  <TD><font face="新細明體">
					  <input name="c_Year" type="text" id="c_Year">
					  </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px" align="right" bgColor="#edeeec"><div align="right"><font size="2" face="新細明體">寬&nbsp;</font></div></TD>
				    <TD><font face="新細明體">
				    <input name="c_Width" type="text" id="c_Width">
				    </font></TD>
				  </TR>
				  <TR>
					  <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec">
			        <div align="right"><font size="2" face="新細明體">高</font></div></TD>
				    <TD style="HEIGHT: 28px"><font face="新細明體">
				    <input name="c_Height" type="text" id="c_Height">
				    </font></TD>
				  </TR>
				  <TR>
                    <TD style="WIDTH: 103px;" align="right" bgColor="#edeeec">
                    <div align="right"><font size="2" face="新細明體">圖片路徑</font></div></TD>
                    <TD style="HEIGHT: 28px"><font face="新細明體">
                    <INPUT name="c_FileName" type="file" id="c_FileName" size="70">
                    </font></TD>
			    </TR>
					  <TR>
						  <TD colspan="2" align="right" >
						    <div align="center">							  <font face="新細明體">&nbsp;
					        <input type="submit" name="Submit2" value="上傳">
					        </font></div></TD>
					  </TR>
		      </TABLE>
			  <P align="center">
          </P>
				<div align="center"><font face="新細明體"><a href="FuncList.asp">回功能清單</a></font>
		  </div>
				<div align="center"></div>
    </form>
	</body>
</HTML>
