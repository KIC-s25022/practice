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
Response.Write "ϵͳ������......���Ժ����ԣ���"
Response.End
End If

dim Page_Size : Page_Size = 20 'ÿҳ��ʾ��ͼ������
dim FuLei_KeiShi_XuHao : FuLei_KeiShi_XuHao = "2" 'ͼ�鸸��
dim FuLei_BuMen_XuHao : FuLei_BuMen_XuHao = "3" '���ŷ���
%>