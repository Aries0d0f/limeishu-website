<%@ Page Language="VB" Debug="true" %>
<%@ import Namespace="System.Data" %>
<%@ import Namespace="System.Data.OleDb" %>
<%@ import Namespace="System.IO" %>
<script runat="server">

    Sub Page_Load(sender As Object, e As EventArgs)
    
        If Page.IsPostBack = False Then ' what is here will only happen the first time the page is accessed
            listdatabases   ' call sub to list the Msaccess Databases and pop the pulldown
        end if
    End Sub
    
    ' THIS IS THE SUB WHEN THE "TEST CONNECTION" BUTTON IS PRESSED
    Sub Button1_Click(sender As Object, e As EventArgs)
    
        if databaselist.selecteditem.value <> "select" and databaselist.selectedindex <> -1then
            connecttodatabase(databaselist.selecteditem.value)
        else
            label2.visible = true
        end if
    End Sub
    
    
        ' this is the sub routine that will get the database names and pop the list box
    sub listdatabases
        DataBaseList.Items.clear    ' clears the list before repoping the list
        Dim objDirInfo As DirectoryInfo = New DirectoryInfo(Server.MapPath("/data"))
        Dim arrChildFiles As Array = objDirInfo.GetFiles("*.mdb")
        Dim arrSubFolders As Array = objDirInfo.GetDirectories()
        Dim objChildFile As FileInfo
        'add the "Select Database" to the pulldown
        DataBaseList.Items.Add(New ListItem("Select Database","select"))
        'first loop through the list of databases in the "DATA" and add to the pull down
        For Each objChildFile in arrChildFiles
            DataBaseList.Items.Add(New ListItem("D- " & objChildFile.name & " - " & FormatSize(objChildFile.Length) ,"\data\" & objChildFile.name))
        next
        'Now loop through the list of databases in the "FPDB" and add to the pull down
        objDirInfo = New DirectoryInfo(Server.MapPath("/fpdb"))
        if (objDirInfo.Exists = True) then
            arrChildFiles = objDirInfo.GetFiles("*.mdb")
            arrSubFolders = objDirInfo.GetDirectories()
            For Each objChildFile in arrChildFiles
                DataBaseList.Items.Add(New ListItem("F- " & objChildFile.name & " - " & FormatSize(objChildFile.Length) ,"\fpdb\" & objChildFile.name))
            next
        end if
    end sub
    
    
    
        ' Opens the database and binds the table's information to the datagrid
    sub connecttodatabase(dbname)
        dsContent.CurrentPageIndex = 0  'This resets the page index before attempting to bind the grid
        literal1.text = ""
        dsContent.visible = true   ' makes the datagrid visible
        label2.visible = false      ' if the error label is displayed this turns it off
        literal1.visible = false
      try
        Dim objConn
        Dim strSQL
        objConn = New OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; data source=" & server.mappath(dbname))
        objConn.Open()
        Dim schemaTable As DataTable = objconn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,New Object() {Nothing, Nothing, Nothing, nothing})
        dsContent.DataSource=schemaTable
        dsContent.DataBind()
        objConn.Close()
      catch ex as exception
      literal1.visible = true
      dsContent.visible = false
        Literal1.text = "<font color=red><b>" & ex.message & "<hr>" & ex.tostring & "<b></font>"
      end try
    
    end sub
    
    
        ' formats the database size to KB or MB
    Private Function FormatSize(ByVal dblFileSize as Double) As String
        If dblFileSize < 1024 Then
            Return String.Format("{0:N0} B", dblFileSize)
        ElseIf dblFileSize < 1024 * 1024 Then
            Return String.Format("{0:N2} KB", dblFileSize/1024)
        ElseIf dblFileSize < 1024 * 1024 * 1024 Then
            Return String.Format("{0:N2} MB", dblFileSize/(1024*1024))
        ElseIf dblFileSize >= 1024 * 1024 * 1024 Then
            Return "Size in the GB!"
        End If
    End Function
    
    Sub MasterGrid_Page(sender As Object, e As DataGridPageChangedEventArgs)
    
    End Sub

</script>
?<html>
<head>
    <title>Tech Support - Test MSAccess DataBase Connection</title>
</head>
<body>
    <form runat="server">
        <center><strong><font face="Arial" size="4">TECHNICAL SUPPORT TEST MSACCESS DATABASE
            CONNECTION</font></strong> 
            <p>
                This test script has been uploaded to your site by web hosting technical support because
                of a support request (IDS-8022157) for your hosting account. 
                <br />
                This test script is used to test for Server side issues with connecting to a MSAccess
                database. The pulldown will list Msaccess databases in the data(D-) and FPDB(F-) directories.
                Select the MSAccess database to test and click the "TEST Connection" button.&nbsp;If
                the script successfully connected to the database&nbsp;you will see a grid displaying
                information on the tables from the database, else you will see an error describing
                the failure. 
                <br />
                <br />
            </p>
            <p>
                <asp:Label id="Label2" runat="server" visible="False" forecolor="Red" font-bold="True">PLEASE
                SELECT A VALID DATABASE >>>></asp:Label>
                <asp:DropDownList id="DataBaseList" runat="server"></asp:DropDownList>
                &nbsp;&nbsp;&nbsp;&nbsp; 
                <asp:Button id="Button1" onclick="Button1_Click" runat="server" Text="TEST Connection"></asp:Button>
                <br />
                <asp:Literal id="Literal1" runat="server"></asp:Literal>
                <br />
                <asp:datagrid id="dsContent" runat="server" Visible="False" BorderWidth="3px" BorderColor="Gray" AutoGenerateColumns="False" CellPadding="3" AllowSorting="True">
                    <FooterStyle forecolor="White" backcolor="Navy"></FooterStyle>
                    <HeaderStyle font-bold="True" forecolor="White" backcolor="Navy"></HeaderStyle>
                    <PagerStyle font-bold="True" horizontalalign="Right" forecolor="White" position="TopAndBottom" backcolor="Gray" mode="NumericPages"></PagerStyle>
                    <SelectedItemStyle forecolor="White" backcolor="#9471DE"></SelectedItemStyle>
                    <ItemStyle backcolor="#DEDFDE"></ItemStyle>
                    <Columns>
                        <asp:BoundColumn DataField="TABLE_NAME" HeaderText="Name of Table"></asp:BoundColumn>
                        <asp:BoundColumn DataField="TABLE_TYPE" HeaderText="Type of Table"></asp:BoundColumn>
                        <asp:BoundColumn DataField="DATE_CREATED" HeaderText="Date Table Created"></asp:BoundColumn>
                        <asp:BoundColumn DataField="DATE_MODIFIED" HeaderText="Date Table Last Modified"></asp:BoundColumn>
                    </Columns>
                </asp:datagrid>
            </p>
            <p>
                <strong>List of common errors:</strong> 
                <br />
                These are provided for troubleshooting in case you get any errors.<br />
                <br />
                <table width="75%" align="center" border="0">
                    <tbody>
                        <tr bgcolor="#cccccc">
                            <td valign="top" width="10%">
                                <b>Error:</b></td>
                            <td width="90%">
                                System.Reflection.TargetInvocationException</td>
                        </tr>
                        <tr>
                            <td valign="top" width="10%">
                                <b>Cause:</b></td>
                            <td width="90%">
                                Look in the stack trace for <strong>'Could not find file</strong> <strong>'</strong> This
                                would indicate that the database is not in the path specified.<br />
                                • be sure the correct path/location is used<br />
                                • The spelling of the database name is correct<br />
                                • The database is uploaded 
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                            </td>
                        </tr>
                    </tbody>
                </table>
                <br />
            </p>
        </center>
    </form>
</body>
</html>

