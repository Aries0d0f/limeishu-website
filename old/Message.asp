<!-- #include file="DBconn.asp" -->
<%
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")
rs.CursorLocation=3
rs.PageSize=10
page=request("Page")
if page="" then page=1		

sql=" select * from  Message order by c_DateTime desc "
if rs.State=1 then rs.Close
rs.Open sql,conn,adOpenStatic

	pageNum=" ���� \ "
	for i=1 to rs.PageCount 
		if i=int(page) then 
			pageNum= pageNum & "[" & i & "]"
		else
			pageNum= pageNum & " <A href=Message.asp?page=" & i
			pageNum= pageNum & ">[" & i & "]</A> " 

		end if
	next
	pageNum=pageNum 

if not rs.eof then
rs.AbsolutePage=page 

list=""
rowCount=1
while not rs.eof  and rowCount<=rs.PageSize
	rowCount=rowCount+1
	list=list& " <TR bgColor='#cadde1'>" & chr(13)
	list=list& " <TD height=26><B>"&rs("c_Title")&" </B></TD>" & chr(13)
	list=list& " <TD align='center'><A href='mailto:"&rs("c_Mail")&" '> "&rs("c_User")&"  </A></TD> " & chr(13)
	list=list& " <TD align='center' height=26>"&rs("c_Datetime")&"</TD>" & chr(13)	
	list=list& " </TR>" & chr(13)
	list=list& "  <TR bgColor='#ffffff'>" & chr(13)
	list=list& " <TD colSpan=3> "&rs("c_Content")&"&nbsp; </TD>	" & chr(13)
	rs.movenext
	
wend
end if

if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 

%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>����������]==The Li Mei-shu Memorial Gallery==</title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">

		<script language="JavaScript" type="text/JavaScript">

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
		</script>
	</HEAD>
	<body bgcolor="#ffffff" background="images/background.gif" leftmargin="0" topmargin="0"
		marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/introduction%1F%1F%1F%1F_2_06.jpg','images/introduction_2_07.jpg','images/introduction_2_08.jpg','images/introduction_2_09.jpg','images/introduction_2_10.jpg','images/introduction_2_11.jpg','images/introduction_2_13.jpg','images/introduction_2_15.jpg','images/introduction_2_04.jpg','images/introduction_2_06.jpg','images/introduction_2_05.jpg')">
		<form id="Form1" method="post" runat="server">
			<table width="1004" border="0" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<td colspan="2"><table width="1004" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td colspan="2"><img src="images/introduction_01.jpg" width="1004" height="87"></td>
							</tr>
							<tr>
								<td colspan="2"><img src="images/introduction_02.jpg" width="1004" height="74"></td>
							</tr>
							<tr>
								<td width="172"><img src="images/introduction_03.jpg" width="172" height="33"></td>
							  <td width="832"><table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td><A onmouseover="MM_swapImage('Image411','','images/introduction_2_04.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="introduction.htm"><IMG src="images/introduction_04.jpg" alt="�����]²��" name="Image411" width="80" height="33" border="0" id="Image41"></A></td>
											<td><A onmouseover="MM_swapImage('Image511','','images/introduction_2_05.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="limei-shu.htm"><IMG src="images/introduction_05.jpg" alt="�{�ѧ�����" name="Image511" width="81" height="33" border="0" id="Image51"></A></td>
											<td><a href="works.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','images/introduction_2_06.jpg',1)"><img src="images/introduction_06.jpg" alt="�]�ç@�~" name="Image6" width="83" height="33"
														border="0"></a></td>
											<td><A onmouseover="MM_swapImage('Image711','','images/introduction_2_07.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="foundation.htm"><IMG src="images/introduction_07.jpg" alt="�������а���|"
														name="Image711" width="82" height="33" border="0" id="Image71"></A></td>
											<td><A onmouseover="MM_swapImage('Image811','','images/introduction_2_08.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="temple.htm"><IMG src="images/introduction_08.jpg" alt="�T�l���v�q" name="Image811" width="82" height="33" border="0" id="Image81"></A></td>
											<td><A onmouseover="MM_swapImage('Image911','','images/introduction_2_09.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="activity.htm"><IMG src="images/introduction_09.jpg" alt="���ʰT��" name="Image911" width="83" height="33" border="0" id="Image91"></A></td>
											<td><A onmouseover="MM_swapImage('Image1011','','images/introduction_2_10.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="multimedia.htm"><IMG src="images/introduction_10.jpg" alt="�h�C�鼽��" name="Image1011" width="82" height="33" border="0" id="Image101"></A></td>
											<td><A onmouseover="MM_swapImage('Image1111','','images/introduction_2_11.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="learn.htm"><IMG src="images/introduction_11.jpg" alt="���ʱоǰ�" name="Image1111" width="82" height="33" border="0" id="Image111"></A></td>
											<td><img src="images/introduction_2_12.jpg" width="74" height="33"></td>
											<td><A onmouseover="MM_swapImage('Image1311','','images/introduction_2_13.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="diary.htm"><IMG src="images/introduction_13.jpg" alt="�ڪ����[��O"
														name="Image1311" width="89" height="33" border="0" id="Image131"></A></td>
											<td><img src="images/introduction_14.jpg" width="14" height="33"></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td width="172" valign="top" bgcolor="#9dd4c5"><table width="172" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="172"><a href="index.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','images/introduction_2_15.jpg',1)"><img src="images/introduction_15.jpg" alt="�^����" name="Image15" width="172" height="33"
											border="0"></a></td>
							</tr>
							<tr>
								<td><img src="images/message_01_17.jpg" width="172" height="63"></td>
							</tr>
							<tr>
								<td><img src="images/message_01_19.jpg" width="172" height="403"></td>
							</tr>
						</table>
					</td>
					<td width="832" valign="top" background="images/message_01_22.jpg"><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td><img src="images/message_01_16.jpg" width="832" height="33"></td>
							</tr>
							<tr>
								<td height="63" background="images/message_01_18.jpg"><table width="806" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td width="132" height="43">&nbsp;</td>
											<td width="674" valign="middle"><table width="100%" border="0" cellpadding="0" cellspacing="0">
													<tr>
														<td width="3%" height="35" valign="bottom"><font size="2">&gt;&gt;</font></td>
														<td width="53%" valign="bottom"><strong><font color="#330033">�d����</font></strong></td>
														<td width="44%" valign="bottom"><a href="javascript:history.go(-1)" class="down"><img src="images/back.gif" width="49" height="17" border="0"></a></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td height="405" valign="top" style='BACKGROUND-POSITION : 0% 0%; BACKGROUND-ATTACHMENT : fixed; BACKGROUND-IMAGE : url(images/message_01_20.jpg); BACKGROUND-REPEAT : no-repeat'><table width="570" border="0" cellpadding="0" cellspacing="0" class="tdword">
										<tr>
											<td width="10">&nbsp;</td>
											<td width="560">
												<table width="100%" border="0" align="center">
													<tr>
														<td width="477" style="WIDTH: 294px"><%=pageNum%></td>
														<td width="73"><div align="right"><a href="messageAdd.asp" class="down"><img src="images/bbs_new.gif" width="67" height="17" border="0"></a>
												                <!--���X�`��,�ثe���X -->
									                  </div></td>
													</tr>
											  </table> 
												<TABLE id="Table5" cellSpacing="1" cellPadding="5" width="100%" bgColor="#999999" border="0">
                                                  <TR bgColor="#669999">
                                                    <TD align="center" width="32%"><FONT color="#ffffff">�D�D</FONT></TD>
                                                    <TD align="center" width="14%"><FONT color="#ffffff">�o��H</FONT></TD>
                                                    <TD align="center" width="20%" height="12"><FONT color="#ffffff">�i�K�ɶ�</FONT></TD>
                                                  </TR>
												  <%=list%>
                                                </TABLE>												<!--�D�D�峹start -->
												
                                                <table width="100%" border="0" align="center">
													<tr>
														<td width="80%">															    <!--�^�Ф峹�s�� -->
															    <div align="left"><a href="m7_ent_bbs_massage.htm">
																    <!--�o��s�D�D�s�� -->
															    </a><%=pageNum%></div></td>
														<td width="20%"><div align="left"><a href="messageAdd.asp" class="down"><img src="images/bbs_new.gif" width="67" height="17" border="0"></a></div></td>
													</tr>
													<tr>
														<td colspan="2"><div align="center">
																<!--���X -->
																<cc1:PagedControler id="vPagedControler" runat="server" ControlToPaginate="vDataList" BackColor="Transparent"
																	PagerStyle="NextPrevAndListNumericPages"></cc1:PagedControler>
															</div>
														</td>
													</tr>
											  </table> 
                                                <!--�D�D�峹end --></td>
										</tr>
									</table>
									<p><FONT face="�s�ө���"></FONT>&nbsp;</p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td background="images/all_down_01.gif" height="119" colspan="2"><table width="71%" border="0" align="center" cellpadding="0" cellspacing="0">
							<tr>
								<td width="40%" valign="top"><div align="center">�]�}�J�x�W�x�_���T�l���ظ��|�Q�T�ѤQ��<br>
										�q�ܡJ02-2673-2333 �ǯu�J02-2673-6077<br>
										E-mail�J<a href="mailto:lms@limeishu.org" class="top4">lms@limeishu.org</a><br>
										-----------------------------------------------------------------------------------------------------------
										<br>
										Copyrightc 2005.����������] The Li Mei-shu Memorial Gallery.All Rights Reserved
										<br>
									</div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
</body>
</HTML>
