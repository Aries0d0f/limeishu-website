<%

userName="admin"
password="1912"

if request("userName")=userName  and  request("password")=password   then
	session("loginType")="yesLogin"
	response.redirect("FuncList.asp")
else
	response.redirect("Login.asp?msg=<br>±b¸¹©Î±K½X¿ù»~!")
end if


%>