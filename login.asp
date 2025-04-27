<%
LHB_GongHao=request("LHB_GongHao")  '接受管理员用户名
LHB_MiMa=request("LHB_MiMa")  '接受管理员密码
keys=request("keys")
%> 
<!--#include file="config.asp"-->
<%
if LHB_GongHao<>"" and LHB_MiMa<>"" then
set Lours=Louconn.execute("select * from LHB_USER_DATA where LHB_GongHao='"& LHB_GongHao & "'")				'判断是否存在该用户
if not (Lours.bof and Lours.eof) then
if Lours("LHB_MiMa")=LHB_MiMa then
session("LHB_MiMa")=Lours("LHB_MiMa")           	
session("LHB_GongHao")=Lours("LHB_GongHao")		'为真则将该用户的该权限放入SESSION中
session("LHB_XingMing")=Lours("LHB_XingMing")	
session("LHB_Group")=Lours("LHB_Group")

Response.Redirect "index.asp"							'跳转到管理员管理页面
End if
else
cwxx=1
end if
end if
%>

<!--#include file="LHB_head.asp"-->
<br><br><br><br>
<table width="768" border="0" cellspacing="0" align="center">
<tr>

<td width="673"></td>
</tr>
</table>
<br><br>
<center><table width="768" border="0" align="center" cellpadding="0" cellspacing="0" class="table4"></center>
<tr>
<td height="30" align="center" class=td2><!--#include file="LHB_Logo_Center.asp"--></td>
</tr>
</table><br><br>
<center>
<table width="300" border="0"><form action="login.asp" method="post" name="form2" id="form2">
<tr align="center"> </tr>
<tr>
<td colspan="2">
<%
if cwxx=1 then 
Response.Write "<p class=p2>老大，想黑我，密码错了！</p>"
end if
%>
</td>
</tr>
<tr>
<td width="197" align="right">ログインID：</td>
<td width="205"><input name="keys" type="hidden" value="submit">
<input name="LHB_GongHao" type="text" class="button1" size="20" value="admin" maxlength="20" onfocus="this.select()" onmouseover="this.focus();"></td>
</tr>
<tr>
<td align="right">パスワード：</td>
<td><input name="LHB_MiMa" type="password" class="button1" size="21" value="admin" maxlength="20" onfocus="this.select()" onmouseover="this.focus();"></td>
</tr>
<tr align="center">
<td colspan="2">
<input type="submit" name="Reg2" value="ログイン" class="button1">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type="reset" name="Submit32" value="リセット" class="button1">
</td>
</tr></form> 
</table>
</center>
<br>
<!--#include file="LHB_copy.asp"-->
