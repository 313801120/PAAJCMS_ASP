<%
'************************************************************
'���ߣ������� ����ͨASP/PHP/ASP.NET/VB/JS/Android/Flash������/��������ϵ)
'��Ȩ��Դ������ѹ�����������;����ʹ�á� 
'������2018-03-13
'��ϵ��QQ313801120  ����Ⱥ35915100(Ⱥ�����м�����)    ����313801120@qq.com   ������ҳ sharembweb.com
'����������ĵ������¡����Ⱥ(35915100)�����(sharembweb.com)���
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<%

'dim t:  set t=new testClass
'call echo("",t.id)
't.id=333
'call echo("",t.id)

class testClass
	public id 
	 '���캯�� ��ʼ��
    Sub Class_Initialize()
		id=3
	end sub
    '�������� ����ֹ
    Sub Class_Terminate()
        'HtmlFolder=nothing
        'HtmlFilename=nothing
        'HtmlContent=nothing
        'Urlname=nothing
    End Sub
	
	function getID()
		getID=id
	end function
	sub setID(idStr)
		id=idStr
	end sub
end class 


%>