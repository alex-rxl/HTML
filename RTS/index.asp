<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Session.CodePage=65001%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>login</title>
<link rel="stylesheet" href="css\log.css" type="text/css" />
</head>

<body>

<div class="login-03">
	<div class="third-login">
		<h1 align="center">Log In with IKEA</h1>
		<form class="three" method="post" action="login.asp"> 
			<p align="center">Username</p>
			<li class="base">
 			<div align="center">
			    <input type="text" name="username" value="<%=request.cookies("hehe")("username")%>" required />	
			</div>
			</li>
			<p align="center">Password</p>
			<li>
			<div align="center">
			    <input type="password" name="password" value="<%=request.cookies("hehe")("password")%>" required />	
			</div>
			</li>
			<div class="submit-three">
			<div align="center">
			    <input type="submit" value="Log In" > 
			</div>
			</div>
            <p align="center"></p>
            <label>
              <div align="right"><a href="javascript:alert('账号及密码的相关事宜请联系作者：Likang.huang@ikea.com')">Help...</a></div>
            </label>
		</form>
	</div>  
</div>
</body>
</html>