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
'Email����

'�������� call Send_Email("313801120@qq.com", "����", "111", "����")
function send_Email(toMail, toTitle, myEmalName, toContent)
    dim JMail, isgo, mailBody 
    response.addHeader "Content-Type", "text/html; charset=gb2312" 
    set JMail = createObject("JMail.Message")
        'JMail.ISOEncodeHeaders = False ' �Ƿ����ISO���룬Ĭ��ΪTrue
        JMail.contentTransferEncoding = "base64" 
        JMail.encoding = "base64" 
        JMail.contentType = "text/html"                                                 '������ʾ���ݵĴ���
        JMail.silent = true 
        JMail.logging = true 
        JMail.charset = "gb2312" 
        JMail.mailServerUserName = "m18251922007"                                       '��Ϊ������ĵ�¼�ʺţ�ʹ��ʱ�����Ϊ�Լ��������¼�ʺ�
        JMail.mailServerPassword = "mydd3a"                                             '��Ϊ������ĵ�¼���룬ʹ��ʱ�����Ϊ�Լ��������¼����
        JMail.from = "m18251922007@163.com"                                             '"m18251922007@163.com" '������Email
        JMail.fromName = myEmalName                                                     '����������
        JMail.addRecipient toMail                                                       '�ռ���Email
        JMail.subject = toTitle                                                         '�ʼ�����
        '�ʼ����壨HTML(ע���ż������Ӹ����ķ�ʽ)��
        mailBody = mailBody & "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>BODY {FONT-SIZE: 9pt}</style></head><body>" 
        mailBody = mailBody & toContent 
        mailBody = mailBody & "</body></html>" 
        JMail.body = mailBody                                                           '�ʼ�����
        send_Email = JMail.send("smtp.163.com")                                         'SMTP��������ַ         //���ط����Ƿ�ɹ�
        JMail.close 
    set JMail = nothing 
end function 
'���䷢�� ����Response.Write( ServerSend_Email("m18251922007","mydd3a","313801120@qq.com", "����", "11@aa.com", "����"  ) )
'���䷢��
function serverSend_Email(serverUserName, serverPassword, toMail, toTitle, myEmalName, toContent)
    dim JMail, isgo, mailBody 
    response.addHeader "Content-Type", "text/html; charset=gb2312" 
    set JMail = createObject("JMail.Message")
        'JMail.ISOEncodeHeaders = False ' �Ƿ����ISO���룬Ĭ��ΪTrue
        JMail.contentTransferEncoding = "base64" 
        JMail.encoding = "base64" 
        JMail.contentType = "text/html"                                                 '������ʾ���ݵĴ���
        JMail.silent = true 
        JMail.logging = true 
        JMail.charset = "gb2312" 
        JMail.mailServerUserName = serverUserName                                       '��Ϊ������ĵ�¼�ʺţ�ʹ��ʱ�����Ϊ�Լ��������¼�ʺ�
        JMail.mailServerPassword = serverPassword                                       '��Ϊ������ĵ�¼���룬ʹ��ʱ�����Ϊ�Լ��������¼����
        JMail.from = "m18251922007@163.com"                                             '"m18251922007@163.com" '������Email
        JMail.fromName = myEmalName                                                     '����������
        JMail.addRecipient toMail                                                       '�ռ���Email
        JMail.subject = toTitle                                                         '�ʼ�����
        '�ʼ����壨HTML(ע���ż������Ӹ����ķ�ʽ)��
        mailBody = mailBody & "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>BODY {FONT-SIZE: 9pt}</style></head><body>" 
        mailBody = mailBody & toContent 
        mailBody = mailBody & "</body></html>" 
        JMail.body = mailBody                                                           '�ʼ�����
        serverSend_Email = JMail.send("smtp.163.com")                                   'SMTP��������ַ         //���ط����Ƿ�ɹ�
        JMail.close 
    set JMail = nothing 
end function 
%> 


