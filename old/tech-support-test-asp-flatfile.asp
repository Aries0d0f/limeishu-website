<% @ LANGUAGE = VBScript %>
<% Option Explicit %>
<% ' Listing 12.1   Creating and Using a TextStream Object
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2 ' Opens the file using the system default.
Const TristateTrue = -1 ' Opens the file as Unicode.
Const TristateFalse = 0 ' Opens the file as ASCII.
Dim objFS, objTextS, strLine

' ### JUST WANT TO PRINT SOME INFO BEFORE WE DO ANYTHING

Response.Write "This test script has been uploaded to your site by web hosting technical support because of a support request (IDS-8022157) for your hosting account.<br><br>"
Response.Write "If this test script executes correctly you will see the current time and date printed below.  this information will also be written to a file.  Every time you refresh this page, you will see the updated time and date and you will also see the full contents of the flat file we have been writing our timestamp too.  Otherwise you will see an error describing the failure.<br><br>"
Response.Write "<hr>"
Response.Write "<br>"



' ### FIRST PART: APPEND TEXT TO FILE
Set objFS=Server.CreateObject("Scripting.FileSystemObject")
If objFS.FileExists("d:\webs\limeis\data\tech-support-test-asp-flatfile.txt") = True Then
  Set objTextS = objFS.OpenTextFile("d:\webs\limeis\db\tech-support-test-asp-flatfile.txt", ForAppending, False, TristateFalse)
Else
  Set objTextS = objFS.CreateTextFile("d:\webs\limeis\db\tech-support-test-asp-flatfile.txt",False, False)
End If
objTextS.WriteLine "This was written at " & Now & "."
objTextS.Close

' ### SECOND PART: READ CONTENT FROM FILE
Set objTextS = objFS.OpenTextFile("d:\webs\limeis\db\tech-support-test-asp-flatfile.txt", ForReading,TristateFalse)
Response.Write "The Content of the File: <BR><BR>" & VbCrLf
Do While objTextS.AtEndOfStream <> True
  strLine = objTextS.ReadLine
  strLine = Server.HTMLEncode(strLine)
  Response.Write strLine & "<BR>" & VbCrLf
Loop

objTextS.Close
Set objTextS = Nothing
Set objFS = Nothing
%>