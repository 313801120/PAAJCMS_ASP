<%
'************************************************************
'作者：云祥孙 【精通ASP/PHP/ASP.NET/VB/JS/Android/Flash，交流/合作可联系)
'版权：源代码免费公开，各种用途均可使用。 
'创建：2018-03-16
'联系：QQ313801120  交流群35915100(群里已有几百人)    邮箱313801120@qq.com   个人主页 sharembweb.com
'更多帮助，文档，更新　请加群(35915100)或浏览(sharembweb.com)获得
'*                                    Powered by PAAJCMS 
'************************************************************
%>
<% 
'无缓冲
sub noCache()
    response.buffer = true 
    response.expires = 0 
    response.expiresAbsolute = now() - 1 
    response.cacheControl = "no-cache" 
    response.addHeader "Expires", date() 
    response.addHeader "Pragma", "no-cache" 
    response.addHeader "Cache-Control", "private, no-cache, must-revalidate" 
end sub 

'清除缓冲(20150907)
function XY_clearCache(action)
    on error resume next                                                     '暂时这个说，应该在后台里写是否缓冲
    dim operate 
    operate = lCase(rParam(action, "operate")) 
    if operate = "true" and handleisCache = false then
        call noCache() 
        handleisCache = false 
    end if 
end function 
%> 

