

<Script Language=JavaScript>

function CheckData(){
	msg="";
	if (document.form1.c_Title.value=='') msg=msg+"\n*  主題";
	if (document.form1.c_User.value=='') msg=msg+"\n*  發表人";
	if (document.form1.c_Mail.value=='') msg=msg+"\n*  email";
	if (document.form1.c_City.value=='') msg=msg+"\n*  城市";
	if (document.form1.c_Content.value=='') msg=msg+"\n*  留言內容";				
	
	if (msg!="") {
		 alert("請輸入以下資料" + msg);
		 return false;
	}else{
	 	document.form1.action="MessageAddsave.asp";
		return true;
	}		


}

</Script>	

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>李梅樹紀念館==The Li Mei-shu Memorial Gallery==</title>
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
													href="introduction.htm"><IMG src="images/introduction_04.jpg" alt="紀念館簡介" name="Image411" width="80" height="33" border="0" id="Image41"></A></td>
											<td><A onmouseover="MM_swapImage('Image511','','images/introduction_2_05.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="limei-shu.htm"><IMG src="images/introduction_05.jpg" alt="認識李梅樹" name="Image511" width="81" height="33" border="0" id="Image51"></A></td>
											<td><a href="works.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','images/introduction_2_06.jpg',1)"><img src="images/introduction_06.jpg" alt="館藏作品" name="Image6" width="83" height="33"
														border="0"></a></td>
											<td><A onmouseover="MM_swapImage('Image711','','images/introduction_2_07.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="foundation.htm"><IMG src="images/introduction_07.jpg" alt="李梅樹文教基金會"
														name="Image711" width="82" height="33" border="0" id="Image71"></A></td>
											<td><A onmouseover="MM_swapImage('Image811','','images/introduction_2_08.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="temple.htm"><IMG src="images/introduction_08.jpg" alt="三峽祖師廟" name="Image811" width="82" height="33" border="0" id="Image81"></A></td>
											<td><A onmouseover="MM_swapImage('Image911','','images/introduction_2_09.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="activity.htm"><IMG src="images/introduction_09.jpg" alt="活動訊息" name="Image911" width="83" height="33" border="0" id="Image91"></A></td>
											<td><A onmouseover="MM_swapImage('Image1011','','images/introduction_2_10.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="multimedia.htm"><IMG src="images/introduction_10.jpg" alt="多媒體播放" name="Image1011" width="82" height="33" border="0" id="Image101"></A></td>
											<td><A onmouseover="MM_swapImage('Image1111','','images/introduction_2_11.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="learn.htm"><IMG src="images/introduction_11.jpg" alt="互動教學區" name="Image1111" width="82" height="33" border="0" id="Image111"></A></td>
											<td><img src="images/introduction_2_12.jpg" width="74" height="33"></td>
											<td><A onmouseover="MM_swapImage('Image1311','','images/introduction_2_13.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="diary.htm"><IMG src="images/introduction_13.jpg" alt="我的參觀日記"
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
								<td width="172"><a href="index.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','images/introduction_2_15.jpg',1)"><img src="images/introduction_15.jpg" alt="回首頁" name="Image15" width="172" height="33"
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
														<td width="53%" valign="bottom"><strong><font color="#330033">留言版</font></strong></td>
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
												<form name="form1" onSubmit="javascript:return CheckData();">
												  <TABLE id="Table5" cellSpacing="1" cellPadding="5" width="100%" border="0">
                                                    <TR bgColor="#669999">
                                                      <TD width="14%" align="center" bgcolor="#6699b0"><FONT color="#ffffff">主題</FONT></TD>
                                                      <TD width="44%" align="center" bgcolor="#6699b0"><div align="left">
                                                          <input name="c_Title" type="text" id="c_Title" size="30">
                                                      </div></TD>
                                                      <TD width="10%" height="12" align="center" bgcolor="#6699b0"><font color="#ffffff">發表人</font></TD>
                                                      <TD width="32%" align="center" bgcolor="#6699b0"><div align="left">
                                                          <input name="c_User" type="text" id="c_User">
                                                      </div></TD>
                                                    </TR>
                                                    <TR bgColor="#8ab0b0">
                                                      <TD align="center"><font color="#FFFFFF">email</font></TD>
                                                      <TD align="center"><div align="left">
                                                          <input name="c_Mail" type="text" id="c_Mail" size="30">
                                                      </div></TD>
                                                      <TD height="12" align="center" bgcolor="#8ab0b0"> <font color="#FFFFFF">城 市</font></TD>
                                                      <TD align="center"><div align="left">
                                                          <input name="c_City" type="text" id="c_City">
                                                      </div></TD>
                                                    </TR>
                                                    <TR bgColor="#8ab0b0">
                                                      <TD align="center"><font color="#FFFFFF">內容</font></TD>
                                                      <TD height="12" colspan="3" align="center"><div align="left">
                                                          <textarea name="c_Content" cols="70" rows="10" id="c_Content"></textarea>
                                                      </div></TD>
                                                    </TR>
                                                    <TR>
                                                      <TD height="12" colspan="4" align="center"><input type="submit" name="Submit" value="  送出留言  "></TD>
                                                    </TR>
                                                    
                                                  </TABLE>
											      </form>
												<!--主題文章start -->
<P align="center">&nbsp;												</P>                                                </td>
										</tr>
									</table>
									<p><FONT face="新細明體"></FONT>&nbsp;</p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td background="images/all_down_01.gif" height="119" colspan="2"><table width="71%" border="0" align="center" cellpadding="0" cellspacing="0">
							<tr>
								<td width="40%" valign="top"><div align="center">館址︰台灣台北縣三峽鎮中華路四十三巷十號<br>
										電話︰02-2673-2333 傳真︰02-2673-6077<br>
										E-mail︰<a href="mailto:lms@limeishu.org" class="top4">lms@limeishu.org</a><br>
										-----------------------------------------------------------------------------------------------------------
										<br>
										Copyrightc 2005.李梅樹紀念館 The Li Mei-shu Memorial Gallery.All Rights Reserved
										<br>
									</div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>

</body>
</HTML>
