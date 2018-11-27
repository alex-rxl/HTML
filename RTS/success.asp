<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Session.CodePage=65001%>

<%
On Error Resume Next
Dim conn,rs,v
Set conn=Server.CreateObject("Adodb.Connection")
Set rs=Server.CreateObject("Adodb.Recordset")
conn.open "driver={microsoft access driver (*.mdb)};uid=admin;pwd=????;dbq="&server.MapPath("../XiaoShou/Daily.mdb")
rs.open "select * from daily where ID=1",conn,1,2
if rs.eof and rs.bof then 
response.write "<script>alert('找不到数据.');</script>"
else
v=rs(0) & "$" & rs(1) & "$" & rs(2) & "$" & rs(3) & "$" & rs(4) & "$" & rs(5) & "$" & rs(6) & "$" & rs(7) & "$" & rs(8) & "$" & rs(9) & "$" & rs(10) & "$" & rs(11) & "$" & rs(12) & "$" & rs(13) & "$" & rs(14) & "$" & rs(15) & "$" & rs(16) & "$" & rs(17) & "$" & rs(18) & "$" & rs(19) & "$" & rs(20) & "$" & rs(21) & "$" & rs(22) & "$" & rs(23) & "$" & rs(24) & "$" & rs(25) & "$" & rs(26) & "$" & rs(27) & "$" & rs(28) & "$" & rs(29) & "$" & rs(30) & "$" & rs(31) & "$" & rs(32) & "$" & rs(33) & "$" & rs(34) & "$" & rs(35) & "$" & rs(36) & "$" & rs(37) & "$" & rs(38) & "$" & rs(39) & "$" & rs(40) & "$" & rs(41) & "$" & rs(42) & "$" & rs(43) & "$" & rs(44) & "$" & rs(45) & "$" & rs(46) & "$" & rs(47) & "$" & rs(48) & "$" & rs(49) & "$" & rs(50) & "$" & rs(51) & "$" & rs(52) & "$" & rs(53) & "$" & rs(54) & "$" & rs(55) & "$" & rs(56) & "$" & rs(57) & "$" & rs(58) & "$" & rs(59) & "$" & rs(60) & "$" & rs(61) & "$" & rs(62) & "$" & rs(63) & "$" & rs(64) & "$" & rs(65) & "$" & rs(66) & "$" & rs(67) & "$" & rs(68) & "$" & rs(69) & "$" & rs(70) & "$" & rs(71) & "$" & rs(72) & "$" & rs(73) & "$" & rs(74) & "$" & rs(75) & "$" & rs(76) & "$" & rs(77) 
end if
set rs=nothing
conn.close
set conn=nothing

Dim Str,zuida,zui2
Str=Split(v,"$")
zuida=cdbl(Str(6))
for i=6 to 75
   if cdbl(Str(i))>zuida then zuida=cdbl(Str(i))
next
zui2=cdbl(Str(3))
for i=3 to 5
   if cdbl(Str(i))>zui2 then zui2=cdbl(Str(i))
next
Dim StoreArray
StoreArray=Array(2499157,2396227,1010234,1315421,1455621,1285845,1535653,2407849,2109550,1412404,1309284,1387240,1255624,1383254,2226275,2266775,1256226,1315799,1386100,1397935,1424738,2454420,2532097,1879501,1182802,1244572,1299369,1187088,1079051,1079051,2777932,2968682,2685308,2570016,2645886,2388840,2022732,1269619,1270637,1318306,1396465,1389509,2461976,2196470,1333907,1524124,1415219,1432832,1524025,2525307,2408855,1500971,1568264,1686360,1522324,1695740,2408863,2428223,1769669,1517792,1443724,1329014,1621097,2213585,2163724,1278309,1323169,1350058,1488010,1562727,2287121,2152762,1274685,1392633,1487851,1336517,1510755,2338362,2285852,1471942,1573434,1480877,1426080,1574763,2487288,2298843,1501707,1619164,1491289,1719057,1545417,2242280,2090121,1474791,1373798,1306353,1717196,1850237,2764188,2714062,1587489,1468917,1667801,1905468,1818789,2934990,2609259,1890154,1448119,1399287,1462317,1413542,2244099,2226025,1550616,1362318,1790829,1502112,1516539,1516539,3083288,3188591,2515215,1583293,1377110,1494586,2182429,2384414,1506960,1226840,1441384,1484734,1834791,2794365,2851741,1716471,1608737,1844509,1826489,1837737,3182328,3482785,2445358,1321933,1246337,1826489,1837737,3182328,3482785,2445358,1321933,1246337,1826489,1713216,1065236,1029891,0,614595,810328,1064842,1172725,1224620,1330840,1303594,1065236,1029891,1162199,1127885,1276649,1593718,1224646,1127150,1046609,1162199,1127885,1276649,1593718,1224646,1127150,1046609,1057275,1163753,1939454,1681087,1095123,977836,1049239,1049041,1177935,1660067,1638954,1039327,956465,1165535,1116068,1109154,1838839,1647934,979226,1279826,1107755,998185,1150453,1705684,1692860,1034205,942923,1039004,888344,1142520,1640078,1493125,992533,857906,944835,944835,1888396,1855244,1450696,937520,919396,992555,988588,1080556,1446630,1638822,928953,1041981,950625,1092169,1154775,1746421,1580693,940707,951951,1023548,892843,1235983,1227370,1227370,2261339,2004758,2500000,1122301,1250986,1911195,1834937,828680,1077068,887010,1245520,1306718,1876317,1674003,1034865,1073946,1123882,992149,1033996,1757569,1619157,931931,1156484,1020040,962994,1091926,1513845,1661885,1053517,1064207,976518,925345,1192008,1554743,1544553,923049,843708,991282,768700,1842458,1741757,1548675,990966,892791,847784,994673,920186,1562162,1541666,767131,1177267,1280007,1234584,1218239,2276589,2035135,1141159,1138992,1335270,1309930,1389782,2143484,2143279,1588412,981312,1028461,991549,1067272,1669516,1699444,967110,1324882,1013274,1315886,1019051,1702470,1732413,1145913,1109545,965660,1315886,1019051,1702470,1732413,1145913,1109545,965660,1315886,1019051,1702470,1732413,1145913,943631,1527528,1100000,1100000,1800000,1800000,1050000,1050000,1050000,1050000,1050000,1800000,1800000,1050000,1050000,1050000,1050000,1050000,1800000,1800000,1050000,1050000,1050000,1050000,1050000,1800000,1800000,1050000,1050000,1050000,1050000,1050000,1800000)

dim k
k=dateDiff("d","2018-09-01",Str(1))'获取今天数组的下标

if left(formatdatetime(now,4),2)<"10" then
k=k-1
end if

%>

<html>
<head>
<meta charset="UTF-8">
<meta property="og:type" content="website" />
<meta property="og:title" content="实时销售数据">
<meta property="og:description" content="实时更新天河、佛山、番禺商场的销售数据。">
<meta property="og:image" content="http://www.ikea465.cn/XiaoShou/logo300.png">
<meta property="og:url" content="http://www.ikea465.cn/XiaoShou/index.asp">
<title>实时销售数据</title>
<link rel="stylesheet" href="css/style.css" media="screen" type="text/css" />

<script type="text/javascript">

//把用户名作为水印，避免用户宣传
window.onload=function addWaterMarker(){
var can = document.createElement('canvas');
var body = document.body;
body.appendChild(can);
can.width=150;
can.height=120;
can.style.display='none';
var cans = can.getContext('2d');
cans.rotate(-20*Math.PI/180);
cans.font = "16px Microsoft JhengHei"; 
cans.fillStyle = "rgba(17, 17, 17, 0.50)";
cans.textAlign = 'left'; 
cans.textBaseline = 'Middle';
cans.fillText(<%=Session.Contents("username")%>,can.width/3,can.height/2);
body.style.backgroundImage="url("+can.toDataURL("image/png")+")";
}
</script>

</head>

<body>
<%
if Session.Contents("login")<>"ok" then 
%>
<script>alert("请重新登陆！");window.location.href = "index.asp";</script>
<%
else
'Response.Write("欢迎登陆，"+Session.Contents("username"))
%>


<h1>数据更新时间: <%=Str(1)&"   " & Str(2)%></h1>
<h1>目标：<%=StoreArray(k)%>，实际：<%=Str(3)%>，完成率：<%=FormatNumber(Str(3)/StoreArray(k)*100,2,-2)%>%</h1>  


<div class="skillbar clearfix " data-percent="<%=Str(3)/zui2*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>Total</span><span><%=FormatNumber(Str(3)/StoreArray(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(3)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(4)/zui2*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(4)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(5)/zui2*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(5)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(6)/zuida*100%>%">
    <div class="skillbar-title" style="background: #27ae60;"><span>HFB01</span><span><%=FormatNumber(Str(6)/H01Array(k)*100,1,-2)%>%</span></div>
    <div class="skillbar-bar" style="background: #e67e22;"></div>
    <div class="skill-bar-percent"><%=Str(6)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(30)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(30)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(54)/zuida*100%>%">
	<div class="skillbar-bar" style="background:#EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(54)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(7)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB02</span><span><%=FormatNumber(Str(7)/H02Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(7)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(31)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(31)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(55)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(55)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(8)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB03</span><span><%=FormatNumber(Str(8)/H03Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background:#e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(8)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(32)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(32)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(56)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(56)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(9)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB04</span><span><%=FormatNumber(Str(9)/H04Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(9)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(33)/zuida*100%>%">
	<div class="skillbar-bar" style="background:#9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(33)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(57)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(57)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(10)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB05</span><span><%=FormatNumber(Str(10)/H05Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(10)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(34)/zuida*100%>%">
	<div class="skillbar-bar" style="background:#9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(34)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(58)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(58)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(11)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB06</span><span><%=FormatNumber(Str(11)/H06Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(11)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(35)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(35)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(59)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(59)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(12)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB07</span><span><%=FormatNumber(Str(12)/H07Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(12)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(36)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(36)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(60)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(60)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(13)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB08</span><span><%=FormatNumber(Str(13)/H08Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(13)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(37)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(37)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(61)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(61)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(14)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB09</span><span><%=FormatNumber(Str(14)/H09Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(14)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(38)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(38)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(62)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(62)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(15)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB10</span><span><%=FormatNumber(Str(15)/H10Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(15)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(39)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(39)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(63)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(63)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(16)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB11</span><span><%=FormatNumber(Str(16)/H11Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(16)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(40)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(40)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(64)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(64)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(17)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB12</span><span><%=FormatNumber(Str(17)/H12Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(17)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(41)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(41)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(65)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(65)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(18)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB13</span><span><%=FormatNumber(Str(18)/H13Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(18)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(42)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(42)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(66)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(66)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(19)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB14</span><span><%=FormatNumber(Str(19)/H14Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(19)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(43)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(43)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(67)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(67)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(20)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB15</span><span><%=FormatNumber(Str(20)/H15Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(20)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(44)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(44)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(68)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(68)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(21)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB16</span><span><%=FormatNumber(Str(21)/H16Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(21)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(45)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(45)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(69)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(69)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(22)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB17</span><span><%=FormatNumber(Str(22)/H17Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(22)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(46)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(46)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(70)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(70)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(23)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB18</span><span><%=FormatNumber(Str(23)/H18Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(23)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(47)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(47)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(71)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(71)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(24)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB19</span><span><%=FormatNumber(Str(24)/H19Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(24)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(48)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(48)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(72)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(72)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(27)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB95</span><span><%=FormatNumber(Str(27)/H95Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(27)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(51)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(51)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(75)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(75)%></div>
</div> 
<p>
<div class="skillbar clearfix " data-percent="<%=Str(28)/zuida*100%>%">
	<div class="skillbar-title" style="background: #27ae60;"><span>HFB96</span><span><%=FormatNumber(Str(28)/H96Array(k)*100,1,-2)%>%</span></div>
	<div class="skillbar-bar" style="background: #e67e22;"></div>
	<div class="skill-bar-percent"><%=Str(28)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(52)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #9DC3E6;"></div>
	<div class="skill-bar-percent"><%=Str(52)%></div>
</div> 
<div class="skillbar clearfix " data-percent="<%=Str(76)/zuida*100%>%">
	<div class="skillbar-bar" style="background: #EE7FFB;"></div>
	<div class="skill-bar-percent"><%=Str(76)%></div>
</div>

<p align="center">
<img src="Tu.PNG" width="517" height="50">
<div style="text-align:center;clear:both;">
</div>
<script src='js/jquery.js'></script>
<script src="js/index.js"></script>

<%
end if
%>
</body>
</html>