<%
on error resume next

dim Louconn,connstr,db
db="kiclib.cn" 
Set Louconn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
Louconn.Open connstr

If Err then
Err.Clear
Set Louconn = Nothing
Response.Write "系统调整中......请稍候再试！！"
Response.End
End If

dim Page_Size : Page_Size = 20 '每页显示的图书条数
dim FuLei_KeiShi_XuHao : FuLei_KeiShi_XuHao = "2" '图书父类
dim FuLei_BuMen_XuHao : FuLei_BuMen_XuHao = "3" '部门分类
%>