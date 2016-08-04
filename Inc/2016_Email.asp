<%
'************************************************************
'作者：红尘云孙(SXY) 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流合作可联系本人)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2016-08-04
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'Email邮箱

'发送邮箱 call Send_Email("313801120@qq.com", "标题", "111", "内容")
function send_Email(toMail, toTitle, myEmalName, toContent)
    dim JMail, isgo, mailBody 
    response.addHeader "Content-Type", "text/html; charset=gb2312" 
    set JMail = createObject("JMail.Message")
        'JMail.ISOEncodeHeaders = False ' 是否进行ISO编码，默认为True
        JMail.contentTransferEncoding = "base64" 
        JMail.encoding = "base64" 
        JMail.contentType = "text/html"                                                 '正常显示内容的代码
        JMail.silent = true 
        JMail.logging = true 
        JMail.charset = "gb2312" 
        JMail.mailServerUserName = "m18251922007"                                       '此为您邮箱的登录帐号，使用时请更改为自己的邮箱登录帐号
        JMail.mailServerPassword = "mydd3a"                                             '此为您邮箱的登录密码，使用时请更改为自己的邮箱登录密码
        JMail.from = "m18251922007@163.com"                                             '"m18251922007@163.com" '发件人Email
        JMail.fromName = myEmalName                                                     '发件人姓名
        JMail.addRecipient toMail                                                       '收件人Email
        JMail.subject = toTitle                                                         '邮件主题
        '邮件主体（HTML(注意信件内链接附件的方式)）
        mailBody = mailBody & "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>BODY {FONT-SIZE: 9pt}</style></head><body>" 
        mailBody = mailBody & toContent 
        mailBody = mailBody & "</body></html>" 
        JMail.body = mailBody                                                           '邮件正文
        send_Email = JMail.send("smtp.163.com")                                         'SMTP服务器地址         //返回发送是否成功
        JMail.close 
    set JMail = nothing 
end function 
'邮箱发送 例：Response.Write( ServerSend_Email("m18251922007","mydd3a","313801120@qq.com", "标题", "11@aa.com", "内容"  ) )
'邮箱发送
function serverSend_Email(serverUserName, serverPassword, toMail, toTitle, myEmalName, toContent)
    dim JMail, isgo, mailBody 
    response.addHeader "Content-Type", "text/html; charset=gb2312" 
    set JMail = createObject("JMail.Message")
        'JMail.ISOEncodeHeaders = False ' 是否进行ISO编码，默认为True
        JMail.contentTransferEncoding = "base64" 
        JMail.encoding = "base64" 
        JMail.contentType = "text/html"                                                 '正常显示内容的代码
        JMail.silent = true 
        JMail.logging = true 
        JMail.charset = "gb2312" 
        JMail.mailServerUserName = serverUserName                                       '此为您邮箱的登录帐号，使用时请更改为自己的邮箱登录帐号
        JMail.mailServerPassword = serverPassword                                       '此为您邮箱的登录密码，使用时请更改为自己的邮箱登录密码
        JMail.from = "m18251922007@163.com"                                             '"m18251922007@163.com" '发件人Email
        JMail.fromName = myEmalName                                                     '发件人姓名
        JMail.addRecipient toMail                                                       '收件人Email
        JMail.subject = toTitle                                                         '邮件主题
        '邮件主体（HTML(注意信件内链接附件的方式)）
        mailBody = mailBody & "<html><head><META content=zh-cn http-equiv=Content-Language><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312""><style type=text/css>BODY {FONT-SIZE: 9pt}</style></head><body>" 
        mailBody = mailBody & toContent 
        mailBody = mailBody & "</body></html>" 
        JMail.body = mailBody                                                           '邮件正文
        serverSend_Email = JMail.send("smtp.163.com")                                   'SMTP服务器地址         //返回发送是否成功
        JMail.close 
    set JMail = nothing 
end function 
%> 


