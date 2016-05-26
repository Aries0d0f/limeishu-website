<%
Sub DBconn (conn)
	set conn=server.CreateObject ("ADODB.Connection")
	set rs=server.CreateObject ("ADODB.Recordset")
	conn.open "DRIVER=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("DB\LMSH1.mdb") 


	
End Sub

Sub DBClose (conn)
	conn.close
	set conn=nothing
End sub

%>
