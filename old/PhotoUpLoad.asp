<%
'if session("loginType")<>"yesLogin" then  response.Redirect("login.asp")

Set Upload = Server.CreateObject ("Persits.Upload.1") 
Upload.OverwriteFiles = true
Count = Upload.SaveVirtual ("data/") 


For Each File in Upload.Files
	c_FileName=File.OriginalFileName
Next


reStr="data/PhotoUpLoadAddSave.asp?"
reStr=reStr & "c_FileName="&c_FileName
For Each Item in Upload.Form
	reStr=reStr& "&" & Item.Name & "=" & Item.Value
Next


 
Set Upload = Nothing 

response.redirect reStr
'response.write  reStr

 
%>