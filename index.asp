<!-- #include file="config.asp"-->
<!-- #include file="LHB_head.asp"-->
<script language="JavaScript" src="css/page.js"></script>
<!-- #include file="LHB_DHT_01.asp"-->
<!-- #include file="LHB_DHT_02.asp"-->
<%
page = CLng(request("page"))
judge=request("judge")
judge2=request("judge2")
judge3=0
L_ShuMing = Request("LouHB_ShuMing")
L_KeShiBM = Request("LouHB_KeShiBM")
L_XiangXiBM = Request("LouHB_XiangXiBM")

Set Lours=server.createobject("adodb.recordset")
If L_ShuMing <>"" Then
sql="Select * From LHB_Office_TuShu where L_ShuMing Like '%"&L_ShuMing&"%'"
Elseif L_KeShiBM <>"" Then
sql="Select * From LHB_Office_TuShu where L_TuShuFL = "&L_KeShiBM&""
Elseif L_XiangXiBM <>"" Then
sql="Select * From LHB_Office_TuShu where L_XiangXiBM = "&L_XiangXiBM&""
Else
sql="Select * From LHB_Office_TuShu order by id desc"
End If
Lours.open sql,Louconn,1,1

if Lours.EOF and Lours.BOF then
response.write("<center><br>�ǩ`���٩`���ˤȤ��礬׷�Ӥ���Ƥ��ޤ���</center>")
else
%>
<table width="768" border="0" cellspacing="0" align="center">
<tr>
<form name="From1" method="Get" action="index.asp">
<td height="30" colspan="9" align=center class="table4">��ֱ����<br>
<input name="LouHB_ShuMing" type="text" class="button1" size="30" /><input name="Search" type="submit" class="button1" value="������Ǘ�������">
</td>
</form>
</tr>
<tr>
<td width="230" height="30" align=center background="pic/DHT-bg.gif" class="table15">����</td>
<td width="100" align=center background="pic/DHT-bg.gif" class="table15">����</td>
<td width="230" align=center background="pic/DHT-bg.gif" class="table15">������</td>
<td width="77" align=center background="pic/DHT-bg.gif" class="table15">��������</td>
<td width="26" align=center background="pic/DHT-bg.gif" class="table15">�ڎ�</td>
<td width="65" align=center background="pic/DHT-bg.gif" class="table15">����</td>
<td width="40" align=center background="pic/DHT-bg.gif" class="table10">����</td>
</tr>
<%
Lours.MoveFirst
Lours.PageSize = Page_Size  '�ک`��������˱�ʾ����륨��ȥ����  
If page < 1 Then page = 1 
If page > Lours.PageCount Then page = Lours.PageCount  
Lours.AbsolutePage = page
For ipage = 1 To Lours.PageSize
%>
<tr>
<td class="table15" height="30"><a href="LHB_TianJ_XiuG.asp?LouHB_ID=<%=Lours("ID")%>&LouHB_TianJ_XiuG=2"><%=Lours("L_ShuMing")%><%=Lours("L_CeShu")%></a></td>
<td align=center class="table15"><%=Lours("L_ZuoZhe")%></td>
<td class="table15"><%=Lours("L_ChuBanShe")%></td>
<td align=center class="table15"><%=Lours("L_ChuBanRQ")%></td>
<td align=center class="table15"><%=Lours("L_TuShuZKC")%></td>
<td align=center class="table15">
<%
If Lours("L_ShuJiaHao")<>"" Then
response.write(""&Lours("L_ShuJiaHao")&"")
Else
response.write("&nbsp;")
End If
%>
</td>
<td align=center class="table10">
<%
If Lours("L_TuShuZKC")<=0 Then
response.write("&nbsp;")
Else
response.write"<a href='LHB_JieYue_New.asp?LouHB_TuShuID="&Lours("id")&"&LouHB_YuJieRQ="&Date()&"&LouHB_JieYueZT=1' onClick=return confirm('ע�⣺�����ˤ��α��򤪽�ꤷ�ޤ�����');>���E</a>"
End If
%>
</td>
</tr>
<%
Lours.MoveNext      
If Lours.EOF Then Exit For   
Next
%>
</table>
<center><table border="0" align="center" cellpadding="0" cellspacing="0" width="768" class="table4">
<form onSubmit="return test();" method="get" action="index.asp" id=form2 name=form2>
<tr> 
<td width="220" height="30" background="pic/bg.gif"> 
<%
If page = 1 Then 
Response.Write "��һҳ����һҳ" 
End If
If page <> 1 Then 
Response.Write "<a href=index.asp?page=1&LouHB_ShuMing="& L_ShuMing &"&LouHB_KeShiBM="& L_KeShiBM &"&LouHB_XiangXiBM="& L_XiangXiBM &">��һҳ</a>" 
Response.Write "��<a href=index.asp?page=" & (page - 1) & "&LouHB_ShuMing="& L_ShuMing &"&LouHB_KeShiBM="& L_KeShiBM &"&LouHB_XiangXiBM="& L_XiangXiBM &">��һҳ</a>" 
End If
If page <> Lours.PageCount Then 
Response.Write "��<a href=index.asp?page=" & (page + 1) & "&LouHB_ShuMing="& L_ShuMing &"&LouHB_KeShiBM="& L_KeShiBM &"&LouHB_XiangXiBM="& L_XiangXiBM &">��һҳ</a>" 
Response.Write "��<a href=index.asp?page=" & Lours.PageCount & "&LouHB_ShuMing="& L_ShuMing &"&LouHB_KeShiBM="& L_KeShiBM &"&LouHB_XiangXiBM="& L_XiangXiBM &">���һҳ</a>" 
End If 
If page = Lours.PageCount Then 
Response.Write "����һҳ�����һҳ" 
End If
%></td>
<td colspan="2" background="pic/bg.gif">��ת���� 
<%Response.Write "<input type=text size=5 maxlength=4 name=page class=button1><input type=hidden name=judge value=1>"  '��ʾ����ҳ���򲢽�page,judge����������ȥ%>
<input type="hidden" name="LouHB_ShuMing" value=<%=L_ShuMing%>><input type="hidden" name="LouHB_KeShiBM" value=<%=L_KeShiBM%>><input type="hidden" name="LouHB_XiangXiBM" value=<%=L_XiangXiBM%>>
ҳ</td>
</tr>
<tr> 
<td height="25" colspan="3">�i���t����<%=Lours.recordCount%>����Ӌ�ک`������<%=Lours.PageCount%>����ǰ�ک`����<%=page%></td>
</tr>
</form>
</table></center>
<%end if

Lours.Close
Louconn.Close
set Lours=nothing
set Louconn=nothing
%>
<!-- #include file="LHB_copy.asp"-->