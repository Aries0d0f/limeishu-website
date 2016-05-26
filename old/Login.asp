<meta http-equiv="Content-Type" content="text/html; charset=big5">	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>Login</title>
	</HEAD>

	<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

			<FONT size="2">			</FONT>
            <form name="form1" method="post" action="LoginCheck.asp">
              <br>
              <br>
              <table width="358" height="165" border="2" align="center" cellpadding="5" cellspacing="2" bgcolor="#cccc99">
                <tr>
                  <td colspan="2"><div align="center"><FONT size="3">系統管理後台-登入</FONT></div></td>
                </tr>
                <tr>
                  <td width="94"><div align="right"><FONT size="2">帳號</FONT></div></td>
                  <td width="258"><input name="userName" type="text" id="userName"></td>
                </tr>
                <tr>
                  <td><div align="right"><FONT size="2">密碼</FONT></div></td>
                  <td><input name="Password" type="password" id="Password"></td>
                </tr>
                <tr>
                  <td colspan="2"><div align="center">
                    <input type="submit" name="Submit" value="    登    入    ">
                  <%=request("msg")%></div></td>
                </tr>
              </table>
            </form>
    </body>
</HTML>
