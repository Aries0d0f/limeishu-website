<%

' response.write session("loginType")
if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")

%>
<meta http-equiv="Content-Type" content="text/html; charset=big5">	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>Function List</title>
	</HEAD>
	
<table width="358" height="165" border="2" align="center" cellpadding="5" cellspacing="2" bgcolor="#cccc99">
  <tr>
    <td><div align="center"><FONT size="3">�t�κ޲z��x-�\��M��</FONT></div></td>
  </tr>
  <tr>
    <td><div align="center"><a href="PhotoUpLoadList.asp">�@�~�W��</a></div></td>
  </tr>
  <tr>
    <td>        <div align="center"><a href="MessageAdmin.asp">�d�����޲z</a></div></td>
  </tr>
</table>
</HTML>