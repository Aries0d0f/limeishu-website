<!-- #include file="DBconn.asp" -->
<%
DBconn conn
set rs=server.CreateObject ("ADODB.Recordset")

c_Type=3
colCount=6


if rtrim(request("c_Year"))="" then 
	c_Year=rtrim(session("c_Year"))
else
	session("c_Year")=rtrim(request("c_Year"))
end if
c_Year=int(session("c_Year"))

if rtrim(request("c_TypeEx"))="" then 
	c_TypeEx=rtrim(session("c_TypeEx"))
else
	session("c_TypeEx")=rtrim(request("c_TypeEx"))
end if

'response.write session("c_Year") & " : c_Year<br>"
'response.write session("c_TypeEx") & " : c_TypeEx<br>"

sql="select c_TypeEx from Photo where  c_Type= " & c_Type & " order by c_TypeEx"
if rs.State=1 then rs.Close
rs.Open sql,conn,adOpenStatic
TypeEx=""
c_TypeExTmp=-1
while not rs.eof 
	if rs("c_TypeEx")<>c_TypeExTmp then 
		if session("c_TypeEx")=1 and rs("c_TypeEx")=1 then TypeEx=TypeEx & "> "
		if rs("c_TypeEx")=1 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=1>油畫</a><br>"
		if session("c_TypeEx")=2 and rs("c_TypeEx")=2 then TypeEx=TypeEx & "> "		
		if rs("c_TypeEx")=2 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=2>水墨</a><br>"
		if session("c_TypeEx")=3 and rs("c_TypeEx")=3 then TypeEx=TypeEx & "> "
		if rs("c_TypeEx")=3 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=3>雕刻</a><br>"
		if session("c_TypeEx")=4 and rs("c_TypeEx")=4  then TypeEx=TypeEx & "> "	
		if rs("c_TypeEx")=4 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=4>字畫</a><br>"
		if session("c_TypeEx")=5 and rs("c_TypeEx")=5 then TypeEx=TypeEx & "> "
		if rs("c_TypeEx")=5 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=5>書信</a><br>"
		if session("c_TypeEx")=6 and rs("c_TypeEx")=6 then TypeEx=TypeEx & "> "
		if rs("c_TypeEx")=6 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=6>素描</a><br>"	
		if session("c_TypeEx")=8 and rs("c_TypeEx")=8 then TypeEx=TypeEx & "> "
		if rs("c_TypeEx")=8 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=8>水彩</a><br>"			
		if session("c_TypeEx")=7 and rs("c_TypeEx")=7 then TypeEx=TypeEx & "> "
		if rs("c_TypeEx")=7 then TypeEx=TypeEx & "<a href=works_01.asp?c_TypeEx=7>其他</a><br>"	
		c_TypeExTmp=rs("c_TypeEx")
	end if
	rs.movenext	
wend
'response.write sql & ":sql c_TypeEx<br>"
	

sql="select * from Photo where  c_Type= " & c_Type
if session("c_Year")="1912" then sql=sql & " and c_Year<'1922'"
if session("c_Year")="1922" then sql=sql & " and c_Year>='1922' and  c_Year<'1932' "
if session("c_Year")="1932" then sql=sql & " and c_Year>='1932' and  c_Year<'1942' "
if session("c_Year")="1942" then sql=sql & " and c_Year>='1942' and  c_Year<'1952' "
if session("c_Year")="1952" then sql=sql & " and c_Year>='1952' and  c_Year<'1962' "
if session("c_Year")="1962" then sql=sql & " and c_Year>='1962' and  c_Year<'1972' "
if session("c_Year")="1972" then sql=sql & " and c_Year>='1972' "
if session("c_TypeEx")<>""  then sql=sql & " and c_TypeEx=" & session("c_TypeEx")
sql=sql & " order by c_TypeEx,c_Year"

if rs.State=1 then rs.Close
rs.Open sql,conn,adOpenStatic
'response.write sql & ":sql  1<br>"

PhotoList="<table width='97%' border=0 cellpadding=0 cellspacing=5  >"
i=0 
isExitPic=0             

if  request("c_Photo")<>"" and  not rs.eof then
	set rs1=server.CreateObject ("ADODB.Recordset")
	sql="select * from Photo where  c_Photo= " & request("c_Photo")
	if rs1.State=1 then rs1.Close
	rs1.Open sql,conn,adOpenStatic

	c_filename=rs1("c_filename")
	if rs1("c_Type")	=1 then c_TypeStr="油畫作品"
	if rs1("c_Type")	=2 then c_TypeStr="書信作品"
	if rs1("c_Type")	=3 then c_TypeStr="照片作品"
	if rs1("c_Type")	=4 then c_TypeStr="文物作品"				
	c_NameStr=rs1("c_Name")		
	c_widthStr=rs1("c_width")
	c_HeightStr=rs1("c_Height")	
	c_YearStr=rs1("c_Year")	
else 
	if not rs.eof then 
		c_filename=rs("c_filename")
		if rs("c_Type")	=1 then c_TypeStr="油畫作品"
		if rs("c_Type")	=2 then c_TypeStr="書信作品"
		if rs("c_Type")	=3 then c_TypeStr="照片作品"
		if rs("c_Type")	=4 then c_TypeStr="文物作品"				
		c_NameStr=rs("c_Name")		
		c_widthStr=rs("c_width")
		c_HeightStr=rs("c_Height")	
		c_YearStr=rs("c_Year")	
	end if
end if
'response.write c_filename & " : c_filename<br>"
'response.write request("c_Photo") & " : c_Photo<br>"


while not rs.eof 
	isExitPic=1
	if  i mod colCount=0 then PhotoList=PhotoList & "<tr>" & chr(13)
	'PhotoList=PhotoList & "<td>&nbsp;</td>" & chr(13)
	PhotoList=PhotoList & "<td align=left width=12% ><font  size=2>  <a href='works_03.asp?c_Photo=" & rs("c_Photo") & "' broder=0 > <img src='db/" & rs("c_FileName") & "' height=60  border='0' ></a></font></td>" & chr(13)	
	

	
	
	i=i+1	
	if  i mod colCount=0 then 
		PhotoList=PhotoList & "</tr>" & chr(13)
		i=0
	end if
	rs.movenext
wend


for j=i to colCount
	PhotoList=PhotoList & "<td  width=12% >&nbsp;</td>" & chr(13)
	j=j+1
next
if isExitPic=0 then PhotoList=PhotoList & "您查詢的區間沒有圖片!"
PhotoList=PhotoList & " </table>"

if rs.State=1 then rs.Close
set rs=nothing
DBClose conn 
%>



<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>李梅樹紀念館==The Li Mei-shu Memorial Gallery==</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<style type="text/css">
		TABLE { FONT-SIZE: 13px; COLOR: #333333; LINE-HEIGHT: 150% }
		.tdword { FONT-SIZE: 13px; COLOR: #0c0c0c; LINE-HEIGHT: 180% }
		.top4 { COLOR: #3366cc; LINE-HEIGHT: 140% }
		A.top4:hover { COLOR: #ff6600; LINE-HEIGHT: 140%; TEXT-DECORATION: none }
</style>
		<script language="JavaScript" type="text/JavaScript">
<!--
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
	<body bgColor="#ffffff" leftMargin="0" background="images/background.gif" topMargin="0"
		onload="MM_preloadImages('images/introduction%1F%1F%1F%1F_2_06.jpg','images/introduction_2_07.jpg','images/introduction_2_08.jpg','images/introduction_2_09.jpg','images/introduction_2_10.jpg','images/introduction_2_11.jpg','images/introduction_2_12.jpg','images/introduction_2_13.jpg','images/introduction_2_15.jpg','images/introduction_2_04.jpg','images/introduction_2_05.jpg','images/banner_b2_04.gif','images/banner_b2_05.gif','images/works_0310_2_02.jpg','images/works_0310_2_03.jpg','images/works_0310_2_04.jpg','images/works_0310_2_05.jpg')">
		<form id="Form1" method="post" runat="server">
			<table cellSpacing="0" cellPadding="0" width="1004" align="center" border="0">
				<tr>
					<td colSpan="2">
						<table cellSpacing="0" cellPadding="0" width="1004" border="0">
							<tr>
								<td colSpan="2"><IMG height="87" src="images/introduction_01.jpg" width="1004"></td>
							</tr>
							<tr>
								<td colSpan="2"><IMG height="74" src="images/introduction_02.jpg" width="1004"></td>
							</tr>
							<tr>
								<td width="172"><IMG height="33" src="images/introduction_03.jpg" width="172"></td>
								<td width="832">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td><A onmouseover="MM_swapImage('Image4','','images/introduction_2_04.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="introduction.htm"><IMG height="33" alt="紀念館簡介" src="images/introduction_04.jpg" width="80" border="0" name="Image4"></A></td>
											<td><A onmouseover="MM_swapImage('Image5','','images/introduction_2_05.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="limei-shu.htm"><IMG height="33" alt="認識李梅樹" src="images/introduction_05.jpg" width="81" border="0" name="Image5"></A></td>
											<td><IMG height="33" src="images/introduction_2_06.jpg" width="83"></td>
											<td><A onmouseover="MM_swapImage('Image7','','images/introduction_2_07.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="foundation.htm"><IMG height="33" alt="李梅樹文教基金會" src="images/introduction_07.jpg" width="82" border="0"
														name="Image7"></A></td>
											<td><A onmouseover="MM_swapImage('Image8','','images/introduction_2_08.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="temple.htm"><IMG height="33" alt="三峽祖師廟" src="images/introduction_08.jpg" width="82" border="0" name="Image8"></A></td>
											<td><A onmouseover="MM_swapImage('Image9','','images/introduction_2_09.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="activity.htm"><IMG height="33" alt="活動訊息" src="images/introduction_09.jpg" width="83" border="0" name="Image9"></A></td>
											<td><A onmouseover="MM_swapImage('Image10','','images/introduction_2_10.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="multimedia.htm"><IMG height="33" alt="多媒體播放" src="images/introduction_10.jpg" width="82" border="0" name="Image10"></A></td>
											<td><A onmouseover="MM_swapImage('Image11','','images/introduction_2_11.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="learn.htm"><IMG height="33" alt="互動教學區" src="images/introduction_11.jpg" width="82" border="0" name="Image11"></A></td>
											<td><A onmouseover="MM_swapImage('Image12','','images/introduction_2_12.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="message.asp"><IMG height="33" alt="留言版" src="images/introduction_12.jpg" width="74" border="0" name="Image12"></A></td>
											<td><A onmouseover="MM_swapImage('Image13','','images/introduction_2_13.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="diary.htm"><IMG height="33" alt="我的參觀日記" src="images/introduction_13.jpg" width="89" border="0"
														name="Image13"></A></td>
											<td><IMG height="33" src="images/introduction_14.jpg" width="14"></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td vAlign="top" width="172" bgColor="#e2c4a0">
						<table cellSpacing="0" cellPadding="0" width="172" border="0">
							<tr>
								<td width="172"><A onmouseover="MM_swapImage('Image15','','images/introduction_2_15.jpg',1)" onmouseout="MM_swapImgRestore()"
										href="index.htm"><IMG height="33" alt="回首頁" src="images/introduction_15.jpg" width="172" border="0" name="Image15"></A></td>
							</tr>
							<tr>
								<td><IMG height="63" src="images/works_01_17.jpg" width="172"></td>
							</tr>
							<tr>
								<td>
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td><IMG height="10" src="images/works_0310_1_01.jpg" width="172"></td>
										</tr>
										<tr>
											<td><A onmouseover="MM_swapImage('Image42','','images/works_0310_2_02.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="works_01.asp"><IMG height="29" src="images/works_0310_1_02.jpg" width="172" border="0" name="Image42"></A></td>
										</tr>
										<tr>
											<td><A onmouseover="MM_swapImage('Image43','','images/works_0310_2_03.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="works_02.asp"><IMG height="31" src="images/works_0310_1_03.jpg" width="172" border="0" name="Image43"></A></td>
										</tr>
										<tr>
											<td><A onmouseover="MM_swapImage('Image44','','images/works_0310_2_04.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="works_03.asp"><IMG height="31" src="images/works_0310_1_04.jpg" width="172" border="0" name="Image44"></A></td>
										</tr>
										<tr>
											<td><A onmouseover="MM_swapImage('Image45','','images/works_0310_2_05.jpg',1)" onmouseout="MM_swapImgRestore()"
													href="works_04.asp"><IMG height="32" src="images/works_0310_1_05.jpg" width="172" border="0" name="Image45"></A></td>
										</tr>
										<tr>
											<td><IMG height="270" src="images/works_0310_1_06.jpg" width="172"></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
					<td vAlign="top" width="832" background="images/works_01_22.jpg">
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td><IMG height="33" src="images/works_01_16.jpg" width="832"></td>
							</tr>
							<tr>
								<td background="images/works_01_18.jpg" height="63"><table cellSpacing="0" cellPadding="0" width="806" border="0">
										<tr>
											<td width="132" height="43">&nbsp;</td>
											<td vAlign="middle" width="674">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td vAlign="bottom" width="3%" height="35"><font size="2">&gt;&gt;</font></td>
														<td vAlign="bottom" width="53%"><strong><font color="#330033">照片錦集</font></strong></td>
														<td vAlign="bottom" width="44%"><A class="down" href="javascript:history.go(-1)"><IMG height="17" src="images/back.gif" width="49" border="0"></A></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td style="BACKGROUND-POSITION: 0% 0%; BACKGROUND-ATTACHMENT: fixed; BACKGROUND-IMAGE: url(images/works_01_20.jpg); BACKGROUND-REPEAT: no-repeat"
									vAlign="top" height="405">
									<table class="tdword" cellSpacing="0" cellPadding="0" width="791" border="0">
										<tr>
											<td width="13">&nbsp;</td>
											<td>
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td><IMG height="14" src="images/time.jpg" width="555"></td>
													</tr>
													<tr>
														<td height="6"><table border="0" align="left" cellpadding="0" cellspacing="0">
                                                          <tr>
                                                            <td width="1%">&nbsp;</td>
                                                            <td><a href="works_03.asp?c_Year=1912"><img src="images/<%if c_Year=1912 then response.write "banner_b2_03.gif" else response.write "banner_b1_03.gif" %>" width="78" height="12" border="0"></a></td>
                                                            <td><a href="works_03.asp?c_Year=1922"><img src="images/<%if c_Year=1922 then response.write "banner_b2_03.gif" else response.write "banner_b1_05.gif" %>" width="78" height="12" border="0"></a></td>
                                                            <td><a href="works_03.asp?c_Year=1932"><img src="images/<%if c_Year=1932 then response.write "banner_b2_03.gif" else response.write "banner_b1_05.gif" %>" width="78" height="12" border="0"></a></td>
                                                            <td><a href="works_03.asp?c_Year=1942"><img src="images/<%if c_Year=1942 then response.write "banner_b2_03.gif" else response.write "banner_b1_05.gif" %>" width="78" height="12" border="0"></a></td>
                                                            <td><a href="works_03.asp?c_Year=1952"><img src="images/<%if c_Year=1952 then response.write "banner_b2_03.gif" else response.write "banner_b1_05.gif" %>" width="78" height="12" border="0"></a></td>
                                                            <td><a href="works_03.asp?c_Year=1962"><img src="images/<%if c_Year=1962 then response.write "banner_b2_03.gif" else response.write "banner_b1_05.gif" %>" width="78" height="12" border="0"></a></td>
                                                            <td><a href="works_03.asp?c_Year=1972"><img src="images/<%if c_Year=1972 then response.write "banner_b2_03.gif" else response.write "banner_b1_05.gif" %>" width="78" height="12" border="0"></a></td>
                                                          </tr>
                                                        </table></td>
													</tr>
													<tr>
													  <td><%=PhotoList%>
												      </td>
												  </tr>
												</table>
												<table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
													<tr>
													  <td width="544" height="317"><% if c_filename<>"" then response.write "<img src='DB/" & c_filename & "'>" %>&nbsp; </td>
													</tr>
											  </table>
												<br>
												<table cellSpacing="0" cellPadding="0" width="575" align="center" border="0" style="WIDTH: 575px; HEIGHT: 40px">
													<tr>
														<td><table cellSpacing="0" cellPadding="0" width="100%" align="center" border="0">
                                                          <tr>
                                                            <td width="18%" >類別︰<%=c_TypeStr%>                                                            </td>
                                                            <td width="36%" >名稱︰
                                                            <%=c_NameStr%> </td>
                                                            <td width="24%">尺寸︰<%=c_WidthStr%>   *  <%=c_HeightStr%>                                                           cm</td>
                                                            <td width="22%">年份︰
                                                                <%=c_YearStr%>年</td>
                                                          </tr>
                                                        </table></td>
													</tr>
												</table>
										  <p>&nbsp;</p>										  </td>
											<td vAlign="top" width="215">
												<div align="left"><%=TypeEx%>
										          </div></td>
										</tr>
									</table>
									<p>&nbsp;</p>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td background="images/all_down_01.gif" colSpan="2" height="119">
						<table cellSpacing="0" cellPadding="0" width="71%" align="center" border="0">
							<tr>
								<td vAlign="top" width="40%">
									<div align="center">館址︰台灣台北縣三峽鎮中華路四十三巷十號<br>
										電話︰02-2673-2333 傳真︰02-2673-6077<br>
										E-mail︰<A class="top4" href="mailto:lms@limeishu.org">lms@limeishu.org</A><br>
										-----------------------------------------------------------------------------------------------------------
										<br>
										Copyright&copy; 2005.李梅樹紀念館 The Li Mei-shu Memorial Gallery.All Rights Reserved
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
