<%
'************************************************************
'���ߣ��쳾����(SXY) ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash��������������ϵ����)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2016-08-04
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'����վ����

'�����ַ��滻
Function specialStrReplace(ByVal content)
    content = Replace(content, "\|", "[$�����ַ�A]$") 
    content = Replace(content, "\-", "[$�����ַ�B]$") 
    content = Replace(content, "\,", "[$�����ַ�C]$") 
    content = Replace(content, "\'", "[$�����ַ�D]$") 
    content = Replace(content, "\""", "[$�����ַ�E]$") 
    specialStrReplace = content 
End Function 
'�������ַ��滻
Function unSpecialStrReplace(ByVal content, startStr)
    content = Replace(content, "[$�����ַ�A]$", startStr & "|") 
    content = Replace(content, "[$�����ַ�B]$", startStr & "-") 
    content = Replace(content, "[$�����ַ�C]$", startStr & ",") 
    content = Replace(content, "[$�����ַ�D]$", startStr & "'") 
    content = Replace(content, "[$�����ַ�E]$", startStr & """") 
    unSpecialStrReplace = content 
End Function 

'��Ŀ���ʹ��� ��ҳ|�ı�|��Ʒ|����|��Ƶ|����|����|����|����|��Ƹ|����
Function handleColumnType(columnName)
    Dim s 
    Select Case columnName
        Case "��ҳ" : s = "home"
        Case "�ı�" : s = "text"
        Case "��Ʒ" : s = "product"
        Case "����" : s = "news"
        Case "��Ƶ" : s = "video"
        Case "����" : s = "download"
        Case "����" : s = "case"
        Case "����" : s = "message"
        Case "����" : s = "feedback"
        Case "��Ƹ" : s = "job"
        Case "����" : s = "order"
    End Select
    handleColumnType = s 
End Function

%> 
