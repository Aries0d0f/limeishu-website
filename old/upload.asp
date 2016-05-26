<%

Set Upload = Server.CreateObject ("Persits.Upload.1") 
Count = Upload.SaveVirtual ("/data/") 

Set Upload = Nothing 
%> 
<% = Count %> file(s) uploaded.
