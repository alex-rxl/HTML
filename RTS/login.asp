<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Session.CodePage=65001%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>login</title>
</head> 
<body>
 
<%
username=Request.Form("username")
password=Request.Form("password")
%>
<%
db="../XiaoShou/Database.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};pwd=????;dbq=" & Server.MapPath(db) 
%>

<%
Set rs = Server.CreateObject( "ADODB.Recordset" )
sql = "select * from test where username='"+username+"';"
Set rs2 = Server.CreateObject( "ADODB.Recordset" )
sql2 = "select * from jilv"
rs.open sql,conn,1,3
if (rs.bof and rs.eof) then
%>
<script>alert("用户名错误。");window.location.href = "index.asp";</script>
<%
else
dbpwd=rs("password")
if password<>dbpwd then
rs2.open sql2,conn,1,3
rs2.Addnew
rs2(1)=now
rs2(2)=username
rs2(3)="密码错误"
rs2.Update
rs2.close
%>
<script>alert("密码错误。");window.location.href = "index.asp";</script>
<%
else
'如果5分钟没有任何操作则判定其已经退出，ok是正常登陆的标志
Session.Timeout=5
Session("username")=username
Session("login")="ok"

response.cookies("hehe")("username")=username  '记住用户名
response.cookies("hehe")("password")=password  '记住密码
response.cookies("hehe").expires=date()+360    'cookies 360日内有效

rs2.open sql2,conn,1,3
rs2.Addnew
rs2(1)=now
rs2(2)=username
rs2(3)="登陆成功"
rs2.Update
rs2.close

%>
<script>window.location.href = "success.asp";</script>
<%
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>